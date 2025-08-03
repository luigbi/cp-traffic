VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmExportXDigital 
   Caption         =   "Export X-Digital"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportXDigital.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9615
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVeh 
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2865
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Left            =   9480
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9240
      Top             =   360
   End
   Begin VB.PictureBox pbcTextWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   255
      Left            =   435
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   4335
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Top             =   165
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "06/03/2024"
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   2580
      Width           =   3870
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   2580
      Width           =   1635
   End
   Begin VB.CommandButton cmdExportTest 
      Caption         =   "Export in Test Mode"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   5805
      Width           =   1665
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   3915
      TabIndex        =   3
      Text            =   "1"
      Top             =   165
      Width           =   405
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   7
      Top             =   4980
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportXDigital.frx":08CA
      Left            =   4215
      List            =   "AffExportXDigital.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2850
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4980
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3570
      ItemData        =   "AffExportXDigital.frx":08CE
      Left            =   6585
      List            =   "AffExportXDigital.frx":08D0
      TabIndex        =   12
      Top             =   1305
      Width           =   2820
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   615
      Top             =   4875
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6480
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   5805
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   9
      Top             =   5805
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label lblNote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   20
      Top             =   5745
      Width           =   3735
   End
   Begin VB.Label lacProcessing 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   5205
      Visible         =   0   'False
      Width           =   9240
   End
   Begin VB.Label Label2 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3030
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   285
      Left            =   135
      TabIndex        =   13
      Top             =   5385
      Width           =   9240
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7035
      TabIndex        =   11
      Top             =   975
      Width           =   1965
   End
   Begin VB.Label lacStartDate 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
   Begin VB.Menu mnuGuide 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Run Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStation 
         Caption         =   "Run Stations Only"
      End
      Begin VB.Menu mnuProgram 
         Caption         =   "Run Programs Only"
      End
      Begin VB.Menu mnuAuthorization 
         Caption         =   "Run Authorizations Only"
      End
      Begin VB.Menu mnuTestChanges 
         Caption         =   "Test updates changes"
      End
      Begin VB.Menu mnuTestErrors 
         Caption         =   "Test Errors"
      End
      Begin VB.Menu menuErrorFile 
         Caption         =   "Get Error File Name"
      End
   End
End
Attribute VB_Name = "FrmExportXDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,2003 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
'  Key:
'udcCriteria.XExportType  0 spot insertion 1 file delivery
'udcCriteriaXSpots 0 all spots  1 regional only
'udcCriteria.XProvider  dynamically set from xml.ini when section <> XDIGITAL but instead XDIGITAL-XX
'udcCriteria.XGenType  0 = send to XDS  1 = generate file
'Pass 0 = ISCI, Pass 1 = HBP, Pass 2 = HB
'tgVffInfo().sXDXMLForm
'    P = ISCI
'    A = HBP
'    S=  HB


'
'  Mode           UnitID                     TransmissionID           SiteID1
'  Cue-By AST     AstCode                    FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'  Cue-Not Ast    yyyymmdd+hb+VefCode        FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'  Cue-Not Ast    yymmdd+hbp+VefCode         FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'
'  ISCI           yyyymmdd+SeqNo+vefCode     FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'
'where hb is hhbbb
'      hbp is hhbbbpp
'      SeqNo is a unique number for each spot
'      UnitID date is the FeedDate
'      StationID is shttStationID
'      ReceiverID is attXDReceiverID
'
'Dan 10/24/14
'ISCI export ("-CU") filename must tie out with traffic->export->carts
'HB HBP export filename must tie out with affiliate->export->ISCI Cross Reference and traffic->export->Audio ISCI Title
'Dan 10/24/14 replaced mFileNameFilter with gFileNameFilter
Option Explicit
Option Compare Text
'Dan M 10/8/10 replaced dateValue with gDateValue throughout form
Private imExportMode As Integer     '0=Standard export; 1=Test Mode
Private imGenerating As Integer     '1=Spot Insertions, 2=File Delivery
Private smDate As String     'Export Date
'5/21/15 Dan only used locally
'Private smXHTDeleteDate As String
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private smExportDirectory As String
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private crf_rst As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset
Private Pff_rst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmAstAdj1() As ASTINFO
Private tmAstAdj2() As ASTINFO
Private tmXDFDInfo() As XDFDINFO
Private tmAstTimeRange() As ASTTIMERANGE
Private tmSvAstTimeRange() As ASTTIMERANGE
Private tmGameTimeRange() As ASTTIMERANGE
Private bmMgsPrevExisted As Boolean
Private lmEqtCode As Long
Private imAnyHBorHBP As Integer
Private smEventProgCodeID As String
Private lmEventGsfCode As Long
Private smAddAdvtToISCI As String
Private smMidnightBasedHours As String
Private smUnitIdByAstCodeForBreak As String
'3/23/15:  Add Send Delays to XDS
Private smSupportXDSDelay As String
Private tmDelayAstInfo() As ASTINFO
Private tmDelaySort() As DELAYSORT
Private Type DELAYSORT
    sKey As String * 30 'Date, time, sequnce number
    lAstIndex As Long
End Type
'7/9/13: Separate the generation of Transpent file from the Unit ID paramter
Private smGenTransparency As String
'4/18/13:  Handle merge vehicles.  These are those vehicles defined as Merge in the vehicle option Program ID field.
'          This is required to handle vehicles that overlay other vehicles and can't be merged by using the Log vehicle.
'          Separate agreements are required for these vehicles, therefore they can be merged into agreements
Private imMergeVefCode() As Integer
Private tmMergeAstInfo() As ASTINFO
'5457
Private tmStationInfo As XDIGITALSTATIONINFO
'5859
Private tmAgreementInfo As XDIGITALAGREEMENTINFO
Private tmAgreements() As XDIGITALAGREEMENTINFO
'dan 11/06/12 to add scrollbar to list box as needed
'also added pbcTextWidth.  Don't remove!
Private lmMaxWidth As Long
'Private Const Section As String = "XDigital"
'5896 large error, change output to user
Private bmIsError As Boolean
Private bmAstFileError As Boolean
'6082
Private rsAstFiles As ADODB.Recordset

Private smXmlErrorFile As String
Private smXmlFileFindError As String
Private myFile As FileSystemObject
'6635
Private imChunk As Integer
Private PreFeedInfo_rst As ADODB.Recordset
Private bmTestError As Boolean
'Dan 4/01/14  Special error: tried to call a service that is not defined in the xml's WebServiceURL
Private bmIsWrongServicePage As Boolean
'6788
Private bmSendStationIds As Boolean
Private bmSendAgreementIds As Boolean

Private xht_rst As ADODB.Recordset
Private xhtInfo_rst As ADODB.Recordset
Private bmAllowXMLCommands As Boolean

Private imHDAdj(0 To 3) As Integer

Private lst_rst As ADODB.Recordset

Private myExport As CLogger
Private smPathForgLogMsg As String
Private smDateForLogs As String
Private Const FILEERROR As String = "XDigitalExport"
Private Const STATIONLOG As String = "xdsStationInformationLog_"
Private Const VEHICLELOG As String = "xdsVehicleInformationLog_"
Private Const AGREEMENTLOG As String = "xdsAgreementInformationLog_"
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const CUMULUSVANTIVESECTION As String = "Vantive"
Private Const CUMULUSBACKOFFICESECTION As String = "BackOffice"
'Private Const STATIONXMLRECEIVERID As String = "ReceiverIDSource"
'6901
Private bmIsAllPorts As Boolean
'6966
Private imMaxRetries As Integer
Private bmAlertAboutReExport As Boolean
'6979 when menu Test- update changes, update xht even though generating file
Private bmTestForceUpdateXHT As Boolean
Private bmWroteTopElement As Boolean
'Dan M 11/14/14 moved to local
'Private lmReExportSent As Long
Private lmReExportDelete As Long
'7180
Private bmReExportForce As Boolean
Private Type REATINDELETIONS
    sSiteId As String * 20
    sTransmissionID As String * 20
    sUnitID As String * 20
End Type
Private tmRetainDeletions() As REATINDELETIONS
'8357
Private tmSentListForDeletionCompare() As REATINDELETIONS
Private Type PROGCODEMATCH
    sProgCodeID As String * 8
    bMatch As Boolean
End Type
Dim tmProgCodeMatch() As PROGCODEMATCH
'9114
Private smUnitIdByAstCodeForISCI As String
'7458
Dim myEnt As CENThelper
'7508
Private Type SiteIDToAtt
    sSite As String
    sAtt As String
End Type
Dim tmSiteIdToAttCode() As SiteIDToAtt
Dim bmFailedToReadReturn As Boolean
'9452
Dim bmSendNotCarried As Boolean
'9818
Dim imSharedHeadEndIsci As Integer
Dim imSharedHeadEndCue As Integer

Private Type XHT
    lCode                 As Long            ' XDS History Table Auto-Code
    lAttCode              As Long            ' Agreement reference code
    iFeedDate(0 To 1)     As Integer         ' Feed Date
    lSiteID               As Long            ' Site ID (Obtained from Station or
                                             ' Agreement)
    sTransmissionID       As String * 20     ' Transmission ID (Feed date:
                                             ' yyyymmdd)
    sUnitID               As String * 20     ' Unit ID (ISCI
                                             ' Model:yyyymmdd+SeqNo+VefCode; Cue
                                             ' Model: astCode or yymmdd+HB[or
                                             ' HBP]+vefCode. Date = Feed date)
    sISCI                 As String * 20     ' ISCI that was exported
    sProgCodeID           As String * 8      ' Program Code ID. For ISCI Mode:
                                             ' Non-Sport Vehicles this is blank;
                                             ' For Sports Vehicle it is obtained
                                             ' from the Event. For Cue Mode:
                                             ' Either obtained from the Vehicle
                                             ' option or obtained from the sport
                                             ' Event.
    sStatus               As String * 1      ' Used to indicate if the Export to
                                             ' XDS was Confirmed as successful.
                                             ' N= Not Confirmed; Blank or C = Confirmed; D=to be deleted

    lgsfCode              As Long            ' Internal game code reference
    sUnused               As String * 5      ' Unused
End Type
Dim hmXht As Integer
Dim tmXht As XHT
Private Enum XDSType
    Authorizations
    Stations
    Programs
    FILEDELIVERY
    Vehicles
    Insertions
End Enum
'8279
Private Type EventIdCueAndCode
    sCue As String
    sCode As String
End Type

Dim smEndDate As String
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private imLastVehColSorted As Integer
Private imLastVehSort As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmLastLogDate As Long
Private imVpfIndex As Integer
Private imLastTabSelected As Integer
Private imBypassAll As Integer
Private smGridTypeAhead As String
Const VEHINDEX = 0
Const LOGINDEX = 1
Const PGMINDEX = 2
Const SPLITINDEX = 3
Const SORTINDEX = 4
Const SELECTEDINDEX = 5
Const LOGSORTINDEX = 6
Const PGMSORTINDEX = 7
Const VEHCODEINDEX = 8
Const SPLITSORTINDEX = 9
'10021
Const ISCIFORM = 0
Const HBPFORM = 1
Const HBFORM = 2
Const ALLSPOTS = 0
Const REGIONALONLY = 1
Const SPOTINSERTION = 0
'11/3/17
Private smEDCDate As String
Private smTxtNumberDays As String
'10933
Dim myZoneAndDSTHelper As cDST
'11063
Dim bmCartReplaceISCI As Boolean

'Private Sub mFillVehicle()
'    'Dim iLoop As Integer
'    'lbcVehicles.Clear
'    'lbcMsg.Clear
'    'chkAll.Value = 0
'    'For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
'    '    lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
'    '    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
'    'Next iLoop
'    '8163
'  '  Dim iLoop As Integer
'    'Dim ilAnyHBorHBP As Integer
'    Dim ilVff As Integer
'    Dim ilVpf As Integer
'    Dim slXDXMLForm As String
'    Dim ilRet As Integer
'    '8163
'    Dim slNowDate As String
'    Dim llVef As Long
'
'    imAnyHBorHBP = False
'    ilRet = gPopVff()
'    ReDim imMergeVefCode(0 To 0) As Integer
'    lbcVehicles.Clear
'    lbcMsg.Clear
'    chkAll.Value = 0
''    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
''        If igExportSource = 2 Then DoEvents
''        ilVff = gBinarySearchVff(tgVehicleInfo(iLoop).iCode)
''        ilVpf = gBinarySearchVpf(CLng(tgVehicleInfo(iLoop).iCode))
''        If (ilVff <> -1) And (ilVpf <> -1) Then
''            If igExportSource = 2 Then DoEvents
''            '4/18/13: Bypass Those vehicles that woll be Merged into other vehicles.
''            '5/21/15: allow ISCI vehicles
''            'If ((tgVffInfo(ilVff).sXDProgCodeID) <> "") Then
''            If (Trim$(tgVffInfo(ilVff).sXDProgCodeID) <> "") Or (tgVpfOptions(ilVpf).iInterfaceID > 0) Then
''                If (Trim$(UCase(tgVffInfo(ilVff).sXDProgCodeID)) <> "MERGE") Then
''                    lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
''                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
''                    slXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
''                    If (slXDXMLForm = "A") Or (slXDXMLForm = "S") Then
''                        imAnyHBorHBP = True
''                    End If
''                Else
''                    imMergeVefCode(UBound(imMergeVefCode)) = tgVehicleInfo(iLoop).iCode
''                    ReDim Preserve imMergeVefCode(0 To UBound(imMergeVefCode) + 1) As Integer
''                End If
''            End If
''        End If
''    Next iLoop
''    '8163
'    slNowDate = Format(gNow(), sgSQLDateForm)
'    '8292
'    'SQLQuery = "SELECT DISTINCT attVefCode FROM att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode  WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND attExportType <> 0 AND (vatwvtVendorId = " & Vendors.XDS_Break & " OR vatwvtVendorId = " & Vendors.XDS_ISCI & ") "
'    SQLQuery = "SELECT DISTINCT attVefCode FROM att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode  WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND (vatwvtVendorId = " & Vendors.XDS_Break & " OR vatwvtVendorId = " & Vendors.XDS_ISCI & ") "
'    Set rst = gSQLSelectCall(SQLQuery)
'    Do While Not rst.EOF
'        If igExportSource = 2 Then DoEvents
'        llVef = gBinarySearchVef(CLng(rst!attvefCode))
'        If llVef <> -1 Then
'            ilVff = gBinarySearchVff(tgVehicleInfo(llVef).iCode)
'            ilVpf = gBinarySearchVpf(CLng(tgVehicleInfo(llVef).iCode))
'            '8163 added vehicle state
'            If (ilVff <> -1) And (ilVpf <> -1) And tgVehicleInfo(llVef).sState = "A" Then
'                If igExportSource = 2 Then DoEvents
'                If (Trim$(tgVffInfo(ilVff).sXDProgCodeID) <> "") Or (tgVpfOptions(ilVpf).iInterfaceID > 0) Then
'                    If (Trim$(UCase(tgVffInfo(ilVff).sXDProgCodeID)) <> "MERGE") Then
'                        lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
'                        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(llVef).iCode
'                        slXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
'                        If (slXDXMLForm = "A") Or (slXDXMLForm = "S") Then
'                            imAnyHBorHBP = True
'                        End If
'                    Else
'                        imMergeVefCode(UBound(imMergeVefCode)) = tgVehicleInfo(llVef).iCode
'                        ReDim Preserve imMergeVefCode(0 To UBound(imMergeVefCode) + 1) As Integer
'                    End If
'                End If
'            End If
'        End If
'        rst.MoveNext
'    Loop
'    '7/6/12: Moved to Form_Activate
'    'If Not ilAnyHBorHBP Then
'    '    udcCriteria.XExportType(1, "E") = False
'    '    udcCriteria.XExportType(1, "V") = vbChecked
'    'End If
'End Sub


'Private Sub chkAll_Click()
'    Dim lRet As Long
'    Dim lRg As Long
'    Dim iValue As Integer
'
'    If imAllClick Then
'        Exit Sub
'    End If
'    If chkAll.Value = vbChecked Then
'        iValue = True
'        If lbcVehicles.ListCount > 1 Then
'            edcTitle3.Visible = False
'            chkAllStation.Visible = False
'            lbcStation.Visible = False
'            lbcStation.Clear
'        Else
'            edcTitle3.Visible = True
'            chkAllStation.Visible = True
'            lbcStation.Visible = True
'        End If
'    Else
'        iValue = False
'    End If
'    If lbcVehicles.ListCount > 0 Then
'        imAllClick = True
'        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
'        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
'        imAllClick = False
'    End If
'
'End Sub

Private Sub chkAll_Click()
    
    Dim llRow As Long
    Dim iValue As Integer
    Dim ilCount As Integer
    
    On Error GoTo ErrHand
    If imAllClick Then
        Exit Sub
    End If
    If imBypassAll Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    ilCount = 0
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1
        If Trim(grdVeh.TextMatrix(llRow, VEHINDEX)) <> "" Then
            If iValue Then
                grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                ilCount = ilCount + 1
            Else
                grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
            End If
            mPaintRowColor llRow
        End If
    Next llRow
    If ilCount <> 1 Then
        edcTitle3.Visible = False
        chkAllStation.Visible = False
        lbcStation.Visible = False
        lbcStation.Clear
    Else
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-chkAll_Click"
    Exit Sub
End Sub

Private Sub chkAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    On Error GoTo ErrHand
    If imAllStationClick Then
        Exit Sub
    End If
    If chkAllStation.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcStation.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-chkAllStation_Click"
    Exit Sub
End Sub


Private Sub cmdExport_Click()
    cmdExport.Enabled = False
    '11/3/17
    tmcDelay.Enabled = False
    imExportMode = 0
    mExport
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload FrmExportXDigital
End Sub


Private Sub cmdExportTest_Click()
    imExportMode = 1
    mExport
End Sub

Private Sub edcDate_GotFocus()
    '11/3/17
    tmcDelay.Enabled = False
    'cmdExport.Enabled = False
    mSetCommands
End Sub

Private Sub edcDate_LostFocus()
    'gSetMousePointer grdVeh, grdVeh, vbHourglass
    'grdVeh.Redraw = False
    'mFindAlertsForGrdVeh
    'imLastVehColSorted = -1
    'imLastVehSort = -1
    'mVehSortCol VEHINDEX
    ''mVehSortCol LOGINDEX
    'grdVeh.Row = 0
    'grdVeh.Col = VEHCODEINDEX
    'grdVeh.Redraw = True
    'gSetMousePointer grdVeh, grdVeh, vbDefault
    'tmcDelay.Enabled = False
    'mSetLogPgmSplitColumns
    tmcDelay.Enabled = False
    tmcDelay.Interval = 500
    tmcDelay.Enabled = True
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    Dim llCol As Long
    Dim llRow As Long
    Dim llPos As Long
    
    If imFirstTime Then
        gSetMousePointer grdVeh, grdVeh, vbHourglass
        mSetGridColumns
        mSetGridTitles
        'gGrid_IntegralHeight grdVeh
        'gGrid_FillWithRows grdVeh
        ''D.S. 07-28-17
        'grdVeh.Height = grdVeh.Height + 30
        udcCriteria.Left = lacStartDate.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = lacStartDate.Top + (5 * lacStartDate.Height) / 8 '(3 * lacStartDate.Height) / 4
        'udcCriteria.Top = edcDate.Top
        udcCriteria.Action 6
        '11/9/18: Expand grid height
        llPos = udcCriteria.GetCtrlBottom()
        edcTitle1.Top = udcCriteria.Top + llPos + 60
        grdVeh.Top = edcTitle1.Top + edcTitle1.Height + 60
        grdVeh.Height = chkAll.Top - grdVeh.Top - grdVeh.RowHeight(0)
        gGrid_IntegralHeight grdVeh
        gGrid_FillWithRows grdVeh
        grdVeh.Height = grdVeh.Height + 30
        lacTitle2.Top = edcTitle1.Top
        lbcMsg.Top = grdVeh.Top
        lbcMsg.Height = grdVeh.Height
        edcTitle3.Top = edcTitle1.Top
        lbcStation.Top = grdVeh.Top
        lbcStation.Height = grdVeh.Height
        
        If UBound(tgEvtInfo) > 0 Then
            grdVeh.Redraw = False
            imBypassAll = True
            chkAll.Value = vbUnchecked
            imBypassAll = False
            lbcStation.Clear
            mClearGrid
            grdVeh.Row = 0
            For llCol = VEHINDEX To SPLITINDEX Step 1
                grdVeh.Col = llCol
                grdVeh.CellBackColor = vbHighlight
            Next llCol
            llRow = grdVeh.FixedRows
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    mAddToGrid llRow, llVef
                End If
            Next ilLoop
            mFindAlertsForGrdVeh
            mVehSortCol VEHINDEX
            'mVehSortCol LOGINDEX
            grdVeh.Row = 0
            grdVeh.Col = VEHCODEINDEX
            chkAll.Value = vbChecked
            If mGetGrdSelCount() = 1 Then
                edcTitle3.Visible = True
                chkAllStation.Visible = True
                chkAllStation.Value = vbUnchecked
                lbcStation.Visible = True
                mFillStations
                chkAllStation.Value = vbChecked
            End If
            grdVeh.Redraw = True
            chkAll.Value = vbUnchecked
            lbcStation.Clear
            'lbcVehicles.Clear
            'For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
            '    llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
            '    If llVef <> -1 Then
            '        'lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
            '        'lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
            '        mAddToGrid CLng(ilLoop), llVef
            '    End If
            'Next ilLoop
            chkAll.Value = vbChecked
'            If lbcVehicles.ListCount = 1 Then
'                imVefCode = lbcVehicles.ItemData(0)
'                edcTitle3.Visible = True
'                chkAllStation.Visible = True
'                chkAllStation.Value = vbUnchecked
'                lbcStation.Visible = True
'                mFillStations
'                chkAllStation.Value = vbChecked
'            End If
        Else
            mFillVehicle
            chkAll.Value = vbChecked
            If Not imAnyHBorHBP Then
                udcCriteria.XExportType(1, "E") = False
                udcCriteria.XExportType(1, "V") = vbUnchecked
                udcCriteria.XExportType(0, "V") = vbChecked
            End If
        End If
        If igTestSystem Then
            udcCriteria.XGenType(0, "E") = False
            udcCriteria.XGenType(0, "V") = False
            udcCriteria.XGenType(1, "V") = True
        End If
        gSetMousePointer grdVeh, grdVeh, vbDefault
        If igExportSource = 2 Then
            slNowStart = gNow()
            cmdExport.Enabled = True
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "XDSResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "XDS Result List, Started: " & slNowStart
            ' pass global so glogMsg will write messages to sgExportResultName
            hgExportResult = hlResult
            For ilLoop = grdVeh.FixedRows To grdVeh.Rows - 1
                If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
                    If Trim(grdVeh.TextMatrix(ilLoop, LOGSORTINDEX)) = "A" Then
                        gLogMsgWODT "W", hlResult, Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) & ": Log Needs Generating"
                    End If
                    If Trim(grdVeh.TextMatrix(ilLoop, PGMSORTINDEX)) = "A" Then
                        gLogMsgWODT "W", hlResult, Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) & ": Agreement Needs to be Checked as Program Structure Has Changed"
                    End If
                    If Trim(grdVeh.TextMatrix(ilLoop, SPLITSORTINDEX)) = "A" Then
                        gLogMsgWODT "W", hlResult, Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) & ": Split Copy"
                    End If
                End If
            Next ilLoop
            cmdExport_Click
            slNowEnd = gNow()
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "XDS Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            hgExportResult = 0
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.2 '1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    
    gSetFonts Me
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If

End Sub

Private Sub Form_Load()
    
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim ilValue10 As Integer
    '11063
    Dim slTemp As String
    
    smGridTypeAhead = ""
    '11/3/17
    smTxtNumberDays = ""
    smEDCDate = ""
    
    lblNote.Visible = False
    lblNote.ForeColor = vbRed
    lblNote.Caption = "* Red Box: Generate Log, Check Programming and/or Create Network Split Fills before running Export."
    imLastVehColSorted = -1
    imLastVehSort = -1
    Screen.MousePointer = vbHourglass
    FrmExportXDigital.Caption = "Export X-Digital - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    'Dan M 11/7/14 this messes up testing if date in past. move to when they have chosen a date
    'smXHTDeleteDate = gObtainStartStd(Format(gDateValue(gObtainStartStd(smDate)) - 1, "m/d/yy"))
    '10/20/17: remove setting the default date
    'edcDate.Text = smDate
    edcDate.Text = ""
    txtNumberDays.Text = 1
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    smEventProgCodeID = ""
    lmEventGsfCode = -1
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    ilRet = gOpenMKDFile(hmXht, "Xht.Mkd")
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    lbcStation.Clear
    'mFillVehicle
    smAddAdvtToISCI = "N"
    smMidnightBasedHours = "N"
    smUnitIdByAstCodeForBreak = "N"
    smUnitIdByAstCodeForISCI = "N"
    gSetMousePointer grdVeh, grdVeh, vbDefault
    imAllStationClick = False
    '7/9/13: Separate the generation of Transparent file from the Unit ID parameter
    smGenTransparency = "N"
    '9114
    SQLQuery = "Select safFeatures6 From SAF_Schd_Attributes"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue10 = Asc(rst!safFeatures6)
        If (ilValue10 And UNITIDBYASTCODEFORISCI) = UNITIDBYASTCODEFORISCI Then
            smUnitIdByAstCodeForISCI = "Y"
        End If
    End If
    SQLQuery = "Select spfUsingFeatures10 From SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue10 = Asc(rst!spfUsingFeatures10)
        If (ilValue10 And ADDADVTTOISCI) = ADDADVTTOISCI Then
            smAddAdvtToISCI = "Y"
        End If
        If (ilValue10 And MIDNIGHTBASEDHOUR) = MIDNIGHTBASEDHOUR Then
            smMidnightBasedHours = "Y"
        End If
        '7/9/13: Separate the generation of Transpent file from the Unit ID parameter
        '        Reinstate Unit ID
        'If (ilValue10 And UNITIDBYASTCODE) = UNITIDBYASTCODE Then
        '    smUnitIdByAstCode = "Y"
        'End If
        If (ilValue10 And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK Then
            smUnitIdByAstCodeForBreak = "Y"
        End If
    End If
    '3/17/16: Build array of Head End zone adjustments
    mBuildHeadEndZoneAdjTable
    '6247 transparency file independent of the Unit ID
    '3/23/15: Add Send Delays to XDS
    SQLQuery = "Select siteGenTransparent as Trans, siteSupportXDSDelay From site where sitecode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!Trans = "Y" Then
            smGenTransparency = "Y"
        End If
        '3/23/15: Add Send Delays to XDS
        smSupportXDSDelay = rst!siteSupportXDSDelay
    Else
        '3/23/15: Add Send Delays to XDS
        smSupportXDSDelay = "N"
    End If
    '9818
    gGetSharedHeadEnd imSharedHeadEndIsci, imSharedHeadEndCue
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        cmdExportTest.Visible = True
    Else
        cmdExportTest.Visible = False
        cmdExport.Left = cmdCancel.Left
        cmdCancel.Left = cmdExportTest.Left
    End If
    chkAll.Value = vbChecked
    ilRet = gPopAvailNames()
    '2/16/18: Verify if vpf needs to be reloaded to obtain vpfLLD
    ilRet = gPopVehicleOptions()
    '11063
    bmCartReplaceISCI = False
    slTemp = "Locations"
    If igTestSystem Then
        slTemp = "TestLocations"
    End If
    If gLoadOption(slTemp, "XDSReplaceISCI", slTemp) Then
        If UCase(slTemp) = "YES" Then
            bmCartReplaceISCI = True
        End If
    End If
    'dan 11/06/12
    lmMaxWidth = lbcMsg.Width
    Screen.MousePointer = vbDefault
    'Dan 03/26/14
    'csi internal guide-for testing help
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuGuide.Visible = True
    End If
    'with menu above...if true, don't erase jeff's error message, so I can test how xds errors are handled.
    bmTestError = False
    bmTestForceUpdateXHT = False
    Set myExport = New CLogger
    myExport.LogPath = myExport.CreateLogName(sgMsgDirectory & FILEERROR)
    myExport.CleanThisFolder = messages
    myExport.CleanFolder
    smDateForLogs = Format(gNow(), "mm-dd-yy")
    smPathForgLogMsg = FILEERROR & "Log_" & smDateForLogs & ".txt"
    '10933
    Set myZoneAndDSTHelper = New cDST
    myZoneAndDSTHelper.StartGetSite XDS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    rst_Gsf.Close
    crf_rst.Close
    Pff_rst.Close
    lst_rst.Close

    cprst.Close
    xht_rst.Close
    rsAstFiles.Close
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    ilRet = gCloseMKDFile(hmXht, "Xht.Mkd")
    
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmAstAdj1
    Erase tmAstAdj2
    Erase tmXDFDInfo
    Erase imMergeVefCode
    Erase tmMergeAstInfo
    Erase tmDelayAstInfo
    Erase tmDelaySort
    Erase tmAstTimeRange
    Erase tmSvAstTimeRange
    Erase tmGameTimeRange
    Erase tmRetainDeletions
    '8357
    Erase tmSentListForDeletionCompare
    Erase tmProgCodeMatch
    Erase tmAgreements
    Erase tmSiteIdToAttCode
    mClosePreFeedInfo
    mCloseXHTInfo
    'Dan M
    Set myExport = Nothing
    Set FrmExportXDigital = Nothing
End Sub

Private Sub grdVeh_Click()

    'Moved to mouse up
    'lbcStation.Clear
    'If mGetGrdSelCount() = 1 Then
    '    edcTitle3.Visible = True
    '    chkAllStation.Visible = True
    '    lbcStation.Visible = True
    '    mFillStations
    'Else
    '    edcTitle3.Visible = False
    '    chkAllStation.Visible = False
    '    lbcStation.Visible = False
    'End If
    'imBypassAll = True
    'chkAll.Value = vbUnchecked
    'imBypassAll = False
End Sub

Private Sub grdVeh_GotFocus()
    cmdCancel.Caption = "&Cancel"
    smGridTypeAhead = ""
End Sub

Private Sub grdVeh_KeyPress(KeyAscii As Integer)

    Dim llRowIndex As Long
    Dim llRow As Long
    
    
    If (KeyAscii = 8) Then
        If (smGridTypeAhead <> "") Then
            smGridTypeAhead = Left(smGridTypeAhead, Len(smGridTypeAhead) - 1)
        End If
    Else
        smGridTypeAhead = smGridTypeAhead & Chr(KeyAscii)
    End If
    
    If (KeyAscii = 0) Then
        Exit Sub
    End If
    
    llRowIndex = gGrid_RowSearch(grdVeh, 0, smGridTypeAhead)
    If (llRowIndex > 0) Then
        For llRow = grdVeh.FixedRows To grdVeh.Rows - 1
            If grdVeh.TextMatrix(llRow, VEHINDEX) <> "" Then
                If (grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1" And llRow <> llRowIndex) Then
                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
                    mPaintRowColor llRow
                ElseIf (llRow = llRowIndex) Then
                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                    mPaintRowColor llRow
                End If
            End If
        Next llRow
        'D.S. 7-28-17
        If grdVeh.TopRow + grdVeh.Height \ grdVeh.RowHeight(llRowIndex) - 2 = llRowIndex Then
            grdVeh.TopRow = grdVeh.TopRow + 1
        End If
        If Not grdVeh.RowIsVisible(llRowIndex) Then
            grdVeh.TopRow = grdVeh.FixedRows
            llRow = grdVeh.FixedRows
            Do
                If grdVeh.RowIsVisible(llRowIndex) Then
                    Exit Do
                End If
                grdVeh.TopRow = grdVeh.TopRow + 1
                llRow = llRow + 1
            Loop While llRow < grdVeh.Rows
            'D.S. 7-28-17
            If grdVeh.TopRow + grdVeh.Height \ grdVeh.RowHeight(llRowIndex) - 2 = llRowIndex Then
                grdVeh.TopRow = grdVeh.TopRow + 1
            End If
        End If
        lmLastClickedRow = llRowIndex
        mShowStations
    End If
End Sub

Private Sub grdVeh_Scroll()
    cmdCancel.Caption = "&Cancel"
    lmScrollTop = grdVeh.TopRow
End Sub

Private Sub lbcStation_Click()
    On Error GoTo ErrHand
    lbcMsg.Clear
    cmdCancel.Caption = "&Cancel"
    If imAllStationClick Then
        Exit Sub
    End If
    If cmdExport.Enabled = False And IsDate(edcDate.Text) And (txtNumberDays.Text <> "") Then
        cmdExport.Enabled = True
        cmdExportTest.Enabled = True
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
    Exit Sub
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigitallbcStation_Click"
    Exit Sub
End Sub

Private Sub grdVeh_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub grdVeh_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub edcDate_Change()
    '8163
    tmcDelay.Enabled = False
    tmcDelay.Interval = 3000
    lbcMsg.Clear
    'If cmdExport.Enabled = False Then
    '    cmdExport.Enabled = True
    '    cmdExportTest.Enabled = True
    '    cmdCancel.Caption = "&Cancel"
    'End If
    tmcDelay.Enabled = True

End Sub

Private Sub edcDate_Validate(Cancel As Boolean)
    '8163
    'tmcDelay_Timer
End Sub

Private Function mExportSpotInsertions(ilPassForm As Integer, slSection As String) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim slStr As String
    Dim ilOkStation As Integer
    Dim ilOkVehicle As Integer
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim ilIncludeSpot As Integer
    Dim llIndex As Long
    Dim slSeqNo As String
    Dim slShortTitle As String
    Dim slLen As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slRCart As String
    Dim slRISCI As String
    Dim slRCreative As String
    Dim slRProd As String
    Dim llRCrfCsfCode As Long
    Dim llRCrfCode As Long
    Dim llCrfCode As Long
    Dim llODate As Long
    Dim ilSeqNo As Integer
    Dim slTransmissionID As String
    Dim slPrevTransmissionID As String
    Dim slRotTime As String
    Dim ilRegionExist As Integer
    Dim slVehicleName As String
    Dim slStationName As String
    Dim llTotalExport As Long
   ' Dim tlXmlStatus As CSIRspGetXMLStatus
    Dim llAdf As Long
    Dim llLstCode As Long
    Dim slISCIPrefix As String
    Dim slXDXMLForm As String
    Dim ilVff As Integer
    Dim slHour As String
    Dim slPrevHour As String
    Dim ilBreakNumber As Integer
    Dim ilPositionNumber As Integer
    Dim slFeedTime As String
    Dim slPrevFeedTime As String
    Dim slHB As String
    Dim slHBP As String
    Dim llIndexLoop As Long
    Dim llIndexStart As Long
    Dim llIndexEnd As Long
    Dim slRotStartDT As String
    Dim slRotEndDT As String
    Dim slVefCode As String
    '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
    Dim slVefCode5 As String
    Dim slStationID As String
    Dim slLength As String
    Dim ilVoiceTracked As Integer
    Dim slXDReceiverID As String
    Dim slProgCodeID As String
    Dim slUnitHBP As String
    Dim slUnitHB As String
    Dim ilStartHour As Integer
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
    Dim slEventProgCodeID As String
    Dim blRetStatus As Boolean
    Dim llAst As Long
    Dim ilLocalAdj As Integer
    Dim ilFirstCue As Integer
    Dim slSendProgAndSiteBy As String   'B=Break, D=Day
    Dim slSendWriteBy As String  'B=Break, D=Day
    Dim llAstCode As Long
    Dim ilPrevVefCode As Integer
    Dim ilLastVefCodeExported As Integer
    
    '6/19/14: ttp 6944 Separate out events by Program Code ID if a sport event
    Dim ilFound As Integer
    Dim slCodeID As String
    Dim ilMatch As Integer
    Dim ilProgCode As Integer
    '7256 and Dan moved these to be local 11/14/14
    Dim blisReExport As Boolean
    Dim llReExportSent As Long
    Dim llNewExportSent As Long
    '6082
    Dim blFailedBecauseNotRegional As Boolean
    '6796
    Dim ilVefIndex As Integer
    Dim blISCIGame As Boolean
    '6882
    Dim ilNeedExport As Integer
    Dim ilSafeChunkSize As Integer
    '7219
    Dim blisCUInHeader As Boolean
    On Error GoTo ErrHand
    '7236  if we hit an error or warning when sending a chunk that is not vehicle, then we don't want to update the xht for the entire vehicle. "Y", "N", or "" means partial wasn't done.
    'Dim slPartialExportMayContinue As String
    Dim slConfirmAtts As String
    Dim slTempAtt As String
    Dim slPreviousAtts As String
    Dim slLogMessage As String
    '7508
    Dim slDoNotReturn As String
    Dim slDoNotMaster As String
    Dim blSentThisVehicle As Boolean
    '8236
    ' stop the adding of info to cue
    Dim blLeaveCueAlone As Boolean
    ' the main determiner if we should get codes/cues from programming
    Dim blIsNonGameEvent As Boolean
    ' so we only open once
    Dim blCEFIsOpen As Boolean
'    '7952 so function will work
'    Dim blUnneeded As Boolean
    '8279
    'cue and code can be multiple.  Split in here.  Per break
    ReDim tlEventIds(0 To 0) As EventIdCueAndCode
    'most of the info we need to copy and add for multiple cues and codes is single piece of info, but not Isci.  Store multiple here.  Per Break, so may contain multiple astcodes
    ReDim slEventIdsIscis(0 To 0) As String
    'index into tlEventIds
    Dim ilEventIdIndex As Integer
    '8299  store the whole 'cef comment' here. Done at the agreement level when filtering.  Use the cefCode as the key
    Dim dlCueAndCodes As Dictionary
    ' false for first time in vehicle...so we only set dictionary one time
    Dim blEventIDVehicleAlreadyRan As Boolean
    '9114
    Dim slUnitIsci As String
    '9452
    Dim ilPledgeStatus As Integer
    Dim tlAstInfo() As ASTINFO
    ReDim tmRetainDeletions(0 To 0) As REATINDELETIONS
    Dim slProp As String
    '9629
    Dim slTempDate As String
    Dim slTempTime As String
    '10933 if true, blIsNonGameEvent must be true as well.
    Dim blIsEventZone As Boolean
    Dim ilAcknowledgeDaylight As Integer
    ReDim slEventZones(0 To 0) As String
    Dim slZoneAdjustment As String
    'udcCriteria.XSpots(1) = Regional apots only
'note on slSendWriteBy since xht added, only vehicle has been tested. 5/7/15, I commented out 'B' and 'D' code
'6882 - add 'V' to slSendWriteBy...send by vehicle.
'6635 Dan M. 'grouping' spots rather than sending one at a time. Turns out, slSendWriteBy = 'D' groups spots on same date (and agreement)
' slSendWriteBy = 'B' sends each spot.  Also, slSendProgAndSiteBy was not implemented, so code below always set to "B"
    'slSendWriteBy = "D"
    slSendWriteBy = "V"
    slSendProgAndSiteBy = "B"
    '6979 I have to write '<Insertions> ' or '<Sites>' at the start of the file, but only at the top
    bmWroteTopElement = False
    bmFailedToReadReturn = False
    blCEFIsOpen = False
    '6979 for messages at end.
    llReExportSent = 0
    lmReExportDelete = 0
    llNewExportSent = 0
    ilNeedExport = 0
    blSentThisVehicle = False
    ilLastVefCodeExported = -1
    'slPartialExportMayContinue = ""
    slConfirmAtts = ""
    ReDim tmSiteIdToAttCode(0)
    ilSafeChunkSize = mSafeChunkSize()
    Set dlCueAndCodes = New Dictionary
    '7219
    If InStr(1, UCase(Trim(slSection)), "-CU", vbBinaryCompare) > 0 Then
        blisCUInHeader = True
    Else
        blisCUInHeader = False
    End If
'    If (ilPassForm = 0) Or (udcCriteria.XSpots(1)) Then
'        slSendWriteBy = "B"
'        slSendProgAndSiteBy = "B"
'        If (ilPassForm = 0) Then
'            smUnitIdByAstCode = "N"
'        End If
'    Else
'        '7/9/13: Dan-should this be testing smGenTransparency or is it correct in testing the Unit ID?
'        If smUnitIdByAstCode = "N" Then
'            slSendWriteBy = "D"
'        Else
'            slSendWriteBy = "B"
'        End If
'        slSendProgAndSiteBy = "B"
'    End If
    'Dan 3/15/13 default to false
    mExportSpotInsertions = False
    llTotalExport = 0
    imExporting = True
    imGenerating = 1
    slPrevTransmissionID = ""
    blRetStatus = True
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    slSDate = smDate
    slEDate = gObtainNextSunday(slSDate)
    If gDateValue(gAdjYear(sEndDate)) < gDateValue(gAdjYear(slEDate)) Then
        slEDate = sEndDate
    End If
    '7180
    bmReExportForce = udcCriteria.XReExport
    '10933 10938
    myZoneAndDSTHelper.isDSTActive sMoDate, sEndDate
    '4/18/13: Build array of spots to be merged into other vehicles
    ilRet = mBuildMergeAstInfo(sMoDate, sEndDate, ilPassForm)
    imVefCode = 0
    For ilLoop = 1 To grdVeh.Rows - 1 Step 1
            If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
                If grdVeh.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                    imVefCode = grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)
                    smVefName = grdVeh.TextMatrix(ilLoop, VEHINDEX)
            Else
                imVefCode = -1
                Exit For
            End If
        End If
    Next ilLoop
    
'    For ilVef = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilVef) Then
'            If imVefCode = 0 Then
'                imVefCode = lbcVehicles.ItemData(ilVef)
'            Else
'                imVefCode = -1
'                Exit For
'            End If
'        End If
'    Next ilVef
    Do
        If igExportSource = 2 Then DoEvents
        '3/23/15:  Add Send Delays to XDS  '9452 add "send not carried" field
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attVoiceTracked, attXDReceiverID, attSendDelayToXDS, vefName,attXDSSendNotCarry "
'        '7952 don't send 'slave' values added
'        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, shttclustergroupId,shttMasterCluster, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attVoiceTracked, attXDReceiverID, attSendDelayToXDS, attMulticast, vefName"
        '7701
        SQLQuery = SQLQuery & " FROM shtt, cptt,vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode "
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        If ilPassForm = ISCIFORM Then
            '7701
            SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.XDS_ISCI
            'SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'X'"
        Else
            '7701
            SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.XDS_Break
           ' SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'B'"
        End If
        SQLQuery = SQLQuery & " AND vefCode = cpttVefCode"
        If imVefCode > 0 Then
            SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
        End If
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            '9452
            bmSendNotCarried = False
            If cprst!attXDSSendNotCarry = "Y" Then
                bmSendNotCarried = True
            End If
            slPreviousAtts = ""
            lacProcessing.Caption = "Checking " & Trim$(cprst!vefName)
            If ilPrevVefCode <> cprst!cpttvefcode Then
                ilPrevVefCode = cprst!cpttvefcode
                '7256 reset per vehicle
                blisReExport = False
                '8299 reset
                blEventIDVehicleAlreadyRan = False
                ilRet = gGetProgramTimes(ilPrevVefCode, smDate, sEndDate, tmAstTimeRange())
                ReDim tmSvAstTimeRange(0 To UBound(tmAstTimeRange)) As ASTTIMERANGE
                For ilLoop = 0 To UBound(tmAstTimeRange) - 1 Step 1
                    tmSvAstTimeRange(ilLoop) = tmAstTimeRange(ilLoop)
                Next ilLoop
                ReDim tmGameTimeRange(0 To UBound(tmAstTimeRange)) As ASTTIMERANGE
                For ilLoop = 0 To UBound(tmAstTimeRange) - 1 Step 1
                    tmGameTimeRange(ilLoop) = tmAstTimeRange(ilLoop)
                Next ilLoop
                '3/25/14: PreFeed
                ilRet = mBuildPreFeedInfo(ilPrevVefCode)
                '6882
                If slSendWriteBy = "V" Then
                    '7/25/14: If only deletes defined, export to x-digital
                    If (ilNeedExport = 0) And (udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked) And (UBound(tmRetainDeletions) > LBound(tmRetainDeletions)) Then
                        If bmWroteTopElement = False Then
                            mAddSurroundingElement ilPassForm, True
                            bmWroteTopElement = True
                            ilNeedExport = 1
                        End If
                    End If
                   If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                        ' at least one to send and chose spot insertions
                        If ilNeedExport > 0 Then
                            blSentThisVehicle = True
                            '6979
                            mAddSurroundingElement ilPassForm, False
                            '7/24/14: Send delete commands
                            ilRet = mSendDeleteCommands()
                            ilNeedExport = 0
                            'All vehicles other than last(or only).
                            '7508
                            If Not mSendAndTestReturn(XDSType.Insertions, slDoNotReturn) Then
                                blRetStatus = False
                                slDoNotMaster = slDoNotMaster & slDoNotReturn
                                If bmIsError Then
                                    '7508
                                    If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                                        myExport.WriteWarning myEnt.ErrorMessage
                                    End If
                                    Exit Function
                                End If
                                myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName, False
                                Call mSetResults("Export not completely successful. see " & smPathForgLogMsg, MESSAGERED)
                            Else
                                If udcCriteria.XGenType(0, slProp) Then
                                    gUpdateLastExportDate imVefCode, slEDate
                                End If
                            End If
                        End If
                        'block nothing to go out. Could be set at the 'partial'
                        If blSentThisVehicle Then
                            blSentThisVehicle = False
                            If bmFailedToReadReturn Then
                                slConfirmAtts = ""
                                slDoNotMaster = ""
                            Else
                                'changes slDoNotMaster to bad atts
                                slConfirmAtts = mAdjustAtts(slConfirmAtts, slDoNotMaster)
                                If Len(slDoNotMaster) = 0 Then
                                    slDoNotMaster = "0,"
                                End If
                            End If
                            'confirm the xht.  The 'unconfirmed' will be wiped out next time through
                            mConfirmXHT slConfirmAtts, slSDate, sEndDate
                            If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
                                ' 'bad' get error; all others get success. this handles normal and forced reexport where atts aren't saved
                                'if bmFailedToReadReturn, slDoNotMaster is blank, so all will get error
                                If Not myEnt.UpdateIncompleteByFilename(EntError, , , slDoNotMaster) Then
                                    myExport.WriteWarning myEnt.ErrorMessage
                                End If
                            End If
                        End If
                        'reset per vehicle because 'partial' may have had issue
                        bmFailedToReadReturn = False
                    End If
                End If  'send by vehicle
                '7236 new vehicle reset, but after send above
                'slPartialExportMayContinue = ""
                slConfirmAtts = ""
                slDoNotMaster = ""
                ReDim tmSiteIdToAttCode(0)
            Else
                ReDim tmAstTimeRange(0 To UBound(tmSvAstTimeRange)) As ASTTIMERANGE
                For ilLoop = 0 To UBound(tmSvAstTimeRange) - 1 Step 1
                    tmAstTimeRange(ilLoop) = tmSvAstTimeRange(ilLoop)
                Next ilLoop
            End If
            If igExportSource = 2 Then DoEvents
            '9/6/11: Wrong place
            'slVefCode = Trim$(Str$(imVefCode))
            slStationID = Trim$(Str$(cprst!shttStationId))
            '6806  don't use station?  set to blank!  if both are false (which would be the case when a dual provider and the xml doesn't have field)
            'we block sending in stations and agreements.  But NOT in spot insertions or file delivery.
            If bmSendAgreementIds = True And (bmSendAgreementIds <> bmSendStationIds) Then
                slStationID = ""
            End If
            If Trim$(slStationID) <> "" Then
                If Val(slStationID) <= 0 Then
                    slStationID = ""
                End If
            End If
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If lbcStation.Selected(ilLoop) Then
                        If lbcStation.ItemData(ilLoop) = cprst!shttCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next ilLoop
            Else
                ilOkStation = True
            End If
'            '7952
'            If ilOkStation Then
'                If gSlave(cprst!shttclustergroupId, cprst!shttMasterCluster, cprst!attMulticast, blUnneeded, "") Then
'                    ilOkStation = False
'                End If
'            End If
            If ilOkStation Then
                ilOkVehicle = False
'                For ilVef = 0 To lbcVehicles.ListCount - 1
'                    If igExportSource = 2 Then DoEvents
'                    If lbcVehicles.Selected(ilVef) Then
                For ilVef = 1 To grdVeh.Rows - 1 Step 1
                    If Trim(grdVeh.TextMatrix(ilVef, VEHINDEX)) <> "" Then
                        If grdVeh.TextMatrix(ilVef, SELECTEDINDEX) = "1" Then
                            imVefCode = grdVeh.TextMatrix(ilVef, VEHCODEINDEX)
                            smVefName = grdVeh.TextMatrix(ilVef, VEHINDEX)
                            'If lbcVehicles.ItemData(ilVef) = cprst!cpttvefcode Then
                            If imVefCode = cprst!cpttvefcode Then
                                'imVefCode = lbcVehicles.ItemData(ilVef)
                                slISCIPrefix = ""
                                ilVff = gBinarySearchVff(imVefCode)
                                ilVpf = gBinarySearchVpf(CLng(imVefCode))
                                If (ilVff <> -1) And (ilVpf <> -1) Then
                                    If ilPassForm = ISCIFORM Then
                                        slISCIPrefix = Trim$(tgVffInfo(ilVff).sXDSISCIPrefix)
                                        If tgVpfOptions(ilVpf).iInterfaceID > 0 Then
                                            slXDXMLForm = "P"
                                        Else
                                            slXDXMLForm = ""
                                        End If
                                    Else
                                        slISCIPrefix = Trim$(tgVffInfo(ilVff).sXDISCIPrefix)
                                        slXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
                                    End If
                                    slProgCodeID = Trim$(tgVffInfo(ilVff).sXDProgCodeID)
                                    '8236 10933
                                    blIsEventZone = False
                                    If UCase(Trim$(slProgCodeID)) = "EVENT" Then
                                        blIsNonGameEvent = True
                                        If tgVffInfo(ilVff).sXDEventZone = "Y" Then
                                            blIsEventZone = True
                                        End If
                                    Else
                                        blIsNonGameEvent = False
                                    End If
                                    If ((ilPassForm = ISCIFORM) And (tgVpfOptions(ilVpf).iInterfaceID > 0)) Then
                                        ilOkVehicle = True
                                    End If
                                    If ((ilPassForm = HBPFORM) And (slXDXMLForm = "A")) Then
                                        ilOkVehicle = True
                                    End If
                                    If ((ilPassForm = HBFORM) And (slXDXMLForm = "S")) Then
                                        ilOkVehicle = True
                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Next ilVef
            End If
            '6547 move test of voice tracking
            If IsNull(cprst!attVoiceTracked) Then
                ilVoiceTracked = False
            Else
                If cprst!attVoiceTracked <> "Y" Then
                    ilVoiceTracked = False
                Else
                    ilVoiceTracked = True
                End If
            End If
            slXDReceiverID = cprst!attXDReceiverId
            '6806  don't use agreement?  set to blank!  if both are false (which would be the case when a dual provider and the xml doesn't have field)
            'we block sending in stations and agreements.  But NOT in spot insertions or file delivery.
            If bmSendStationIds = True And (bmSendAgreementIds <> bmSendStationIds) Then
                slXDReceiverID = ""
            End If
            If Trim$(slXDReceiverID) <> "" Then
                If Val(slXDReceiverID) <= 0 Then
                    slXDReceiverID = ""
                End If
            End If
            If ilOkStation And ilOkVehicle And ilVoiceTracked = 0 Then
                If (ilLastVefCodeExported > 0) And (ilLastVefCodeExported <> imVefCode) Then
                    If igTimes = 1 Then
                        If (lbcStation.ListCount <= 0) Or (chkAllStation.Value = vbChecked) Then
                            'gClearAbf ilLastVefCodeExported, 0, sMoDate, gObtainNextSunday(sMoDate)
                        End If
                    End If
                    ilLastVefCodeExported = imVefCode
                End If
'            If ilOkStation And ilOkVehicle Then
'                If IsNull(cprst!attVoiceTracked) Then
'                    ilVoiceTracked = False
'                Else
'                    If cprst!attVoiceTracked <> "Y" Then
'                        ilVoiceTracked = False
'                    Else
'                        ilVoiceTracked = True
'                    End If
'                End If
'                slXDReceiverID = cprst!attXDReceiverID
'                If Trim$(slXDReceiverID) <> "" Then
'                    If Val(slXDReceiverID) <= 0 Then
'                        slXDReceiverID = ""
'                    End If
'                End If
                slVehicleName = mGetVehicleName(imVefCode)
                '10933 added ilAcknowledgeDaylight
                slStationName = mGetStationName(cprst!shttCode, ilAcknowledgeDaylight)
                If ilPassForm = ISCIFORM Then
                    If Trim$(slXDReceiverID) = "" Then
                        If Trim$(slStationID) = "" Then
                            Call mSetResults(slStationName & " missing XDS Station ID", MESSAGEBLACK)
                            'gLogMsg slStationName & " missing XDS Station ID", smPathForgLogMsg, False
                            myExport.WriteWarning slStationName & " missing XDS Station ID"
                            ilOkStation = False
                        End If
                    Else
                        slStationID = slXDReceiverID
                    End If
                    'Dan M 4/03/14 block if sending ISCI and network id = 0
                    blISCIGame = False
                    ilVpf = gBinarySearchVpf(CLng(imVefCode))
                    If ilVpf <> -1 Then
                        If tgVpfOptions(ilVpf).iInterfaceID = 0 Then
                            Call mSetResults(slVehicleName & " missing XDS Network ID", MESSAGEBLACK)
                           ' gLogMsg slVehicleName & " Missing XDS Network ID", smPathForgLogMsg, False
                            myExport.WriteWarning slVehicleName & " Missing XDS Network ID"
                            ilOkVehicle = False
                        Else
                            '6796
                            ilVefIndex = gBinarySearchVef(CLng(imVefCode))
                            If ilVefIndex <> -1 Then
                                If tgVehicleInfo(ilVefIndex).sVehType = "G" Then
                                    blISCIGame = True
                                Else
                                    blISCIGame = False
                                End If
                            Else
                                mSetResults ilVef & " missing in vehicle array", MESSAGERED
                                myExport.WriteError ilVef & " missing in vehicle array-mExportSpotInsertions", True, False
                                ilOkVehicle = False
                            End If 'vefInfo ok to search
                        End If 'interface id > 0
                    End If 'vpf is set
                Else
                    If Trim$(slXDReceiverID) = "" Then
                        If Trim$(slStationID) = "" Then
                            Call mSetResults(slStationName & " airing " & slVehicleName & " Missing X-Digital Station ID", MESSAGEBLACK)
                            'gLogMsg slStationName & " airing " & slVehicleName & " missing X-Digital Station ID", smPathForgLogMsg, False
                            myExport.WriteWarning slStationName & " airing " & slVehicleName & " missing X-Digital Station ID"
                            ilOkStation = False
                        Else
                            slXDReceiverID = slStationID
                        End If
                    End If
                End If
                '6547 change above means don't need to test here.
                If ilOkStation And ilOkVehicle Then
                    '8357 store each siteid and unit id and test against when deleting
                    ReDim tmSentListForDeletionCompare(0 To 0) As REATINDELETIONS
                    '7508 isci is slStationID; hb is slXD
                    If ilPassForm = ISCIFORM Then
                        mAddSiteAndAtt cprst!cpttatfCode, slStationID
                    Else
                        mAddSiteAndAtt cprst!cpttatfCode, slXDReceiverID
                    End If
                    '7458
                    With myEnt
                        .Agreement = cprst!cpttatfCode
                        .Station = cprst!shttCode
                        .Vehicle = imVefCode
                        .ProcessStart
                    End With
'                If ilOkStation And ilOkVehicle And (Not ilVoiceTracked) Then
                    If igExportSource = 2 Then DoEvents
                    '9/6/11: Moved vehicle setting here because it can change when looking thru CPTT
                    slVefCode = Trim$(Str$(imVefCode))
                    '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
                    slVefCode5 = slVefCode
                    Do While Len(slVefCode5) < 5
                        slVefCode5 = "0" & slVefCode5
                    Loop
                    'Jeff
                    ilVpf = gBinarySearchVpf(CLng(imVefCode))
                    If ilVpf <> -1 Then
                        On Error GoTo ErrHand
                        
                        '12/20/14: Remove XHT records regardless of agreements that are two or more weeks old
                        ''7/26/14: Remove old XHT
                        'mRemoveOldXHT cprst!cpttatfCode
                        
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                        tgCPPosting(0).iStatus = cprst!cpttStatus
                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                        tgCPPosting(0).iVefCode = imVefCode
                        tgCPPosting(0).iShttCode = cprst!shttCode
                        tgCPPosting(0).sZone = cprst!shttTimeZone
                        '3/25/14: PreFeed
                        If PreFeedInfo_rst.RecordCount > 0 Then
                            tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
                        Else
                            tgCPPosting(0).sDate = Format$(smDate, sgShowDateForm) 'Format$(sMoDate, sgShowDateForm)
                        End If
                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                        tgCPPosting(0).iNumberDays = imNumberDays
                        'Create AST records
                        '3/25/14: PreFeed
                        If PreFeedInfo_rst.RecordCount > 0 Then
                            igTimes = 1 'By Week
                        Else
                            If (smSupportXDSDelay = "Y") And (cprst!attSendDelayToXDS = "Y") Then
                                igTimes = 1
                            Else
                                igTimes = 3 'By Date
                            End If
                        End If
                        imAdfCode = -1
                        If igExportSource = 2 Then DoEvents
                        llODate = -1
                        ilStartHour = -1
                        slPrevHour = ""
                        ilBreakNumber = -1
                        ilPositionNumber = -1
                        slPrevFeedTime = ""
'                        '10938 replaced how to get ilLocalAdj
'                        myZoneAndDSTHelper.StationZone = tgCPPosting(0).sZone
'                        myZoneAndDSTHelper.StationHonorDaylight ilAcknowledgeDaylight
'                        ilLocalAdj = myZoneAndDSTHelper.FindZoneDifference
                        '9629 moved above isci/hb-hbp split
                        ilLocalAdj = mStationAdj(tgCPPosting(0).sZone)
                        '10/20/12: Form 0 requires RegionType to be set to 2
                        If ilPassForm = ISCIFORM Then
                            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , , , , True)
                            gFilterAstExtendedTypes tmAstInfo
                            ilRet = mBuildDelayAst(ilPassForm, cprst)
                            mRemoveExtraAirplays
                        Else
'                            '10933
                            myZoneAndDSTHelper.StationZone = tgCPPosting(0).sZone
                            myZoneAndDSTHelper.StationHonorDaylight ilAcknowledgeDaylight
                            'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                            If (smMidnightBasedHours = "Y") And (ilPassForm <> ISCIFORM) Then
                                '6082 change first 'false' to 'true to get rid of 0 in astcode
                                ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , True, , , True)
                                gFilterAstExtendedTypes tmAstInfo
                                ilRet = mBuildDelayAst(ilPassForm, cprst)
                                mRemoveExtraAirplays
                                '3/17/16: Handle any head end zone
                                '9629 moved above
'                                ilLocalAdj = mStationAdj(tgCPPosting(0).sZone)
                                'If Left(tgCPPosting(0).sZone, 1) <> "E" Then
                                If ilLocalAdj <> 0 Then
                                    ReDim tmAstAdj1(LBound(tmAstInfo) To UBound(tmAstInfo)) As ASTINFO
                                    For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                                        tmAstAdj1(llAst) = tmAstInfo(llAst)
                                        '3/17/16: Handle any head end zone
                                        'ilLocalAdj = mAdjustToEstZone(tgCPPosting(0).sZone, tmAstAdj1(llAst).sFeedDate, tmAstAdj1(llAst).sFeedTime)
                                        mAdjustToHeadendZone ilLocalAdj, tmAstAdj1(llAst).sFeedDate, tmAstAdj1(llAst).sFeedTime
                                    Next llAst
                                    'If gDateValue(sMoDate) = gDateValue(slSDate) Then
                                    '    tgCPPosting(0).sDate = DateAdd("d", -7, tgCPPosting(0).sDate)
                                    '    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                                    '    ReDim tmAstAdj2(LBound(tmAstInfo) To UBound(tmAstInfo)) As ASTINFO
                                    '    For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                                    '        tmAstAdj2(llAst) = tmAstInfo(llAst)
                                    '        ilLocalAdj = mAdjustToEstZone(tgCPPosting(0).sZone, tmAstAdj2(llAst).sFeedDate, tmAstAdj2(llAst).sFeedTime)
                                    '    Next llAst
                                    'Else
                                        ReDim tmAstAdj2(0 To 0) As ASTINFO
                                    'End If
                                    If UBound(tmAstAdj2) > LBound(tmAstAdj2) Then
                                        ReDim tmAstInfo(LBound(tmAstAdj1) To UBound(tmAstAdj1) + UBound(tmAstAdj2)) As ASTINFO
                                        For llAst = LBound(tmAstAdj2) To UBound(tmAstAdj2) - 1 Step 1
                                            tmAstInfo(llAst) = tmAstAdj2(llAst)
                                        Next llAst
                                        For llAst = LBound(tmAstAdj1) To UBound(tmAstAdj1) - 1 Step 1
                                            tmAstInfo(UBound(tmAstAdj2) + llAst) = tmAstAdj1(llAst)
                                        Next llAst
                                    Else
                                        ReDim tmAstInfo(LBound(tmAstAdj1) To UBound(tmAstAdj1)) As ASTINFO
                                        For llAst = LBound(tmAstAdj1) To UBound(tmAstAdj1) - 1 Step 1
                                            tmAstInfo(llAst) = tmAstAdj1(llAst)
                                        Next llAst
                                    End If
                                End If
                                
                                '5/15/13: Converting the Program merge times NOT required as all times are in Eastern zone for ESPN
                                
                            Else
                                '6082 change first 'false' to 'true to get rid of 0 in astcode
                                ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , , , , True)
                                gFilterAstExtendedTypes tmAstInfo
                                ilRet = mBuildDelayAst(ilPassForm, cprst)
                                mRemoveExtraAirplays
                                '5/15/13: Convert program times to station zone
                                '         Will not be done at this time as only ESPN will use the merge
                            End If
                            '8299
                            If blIsNonGameEvent Then
                                If Not blCEFIsOpen Then
                                    mOpenCEFFile
                                    blCEFIsOpen = True
                                End If
                                mFilterLibraryEventIds blEventIDVehicleAlreadyRan, tmAstInfo, dlCueAndCodes
                                blEventIDVehicleAlreadyRan = True
                            End If

                        End If
                        
                        '4/25/14:PreFeed
                        mCreateAstPreFeedSpots
                        
                        '4/18/13: Merge spots
                        mMergeAsts ilVpf
                        
                        llIndex = LBound(tmAstInfo)
                        'Dan M not handling merge's attcodes
'                        '6/24/14: Obtain previously generated records
'                        If UBound(tmAstInfo) > LBound(tmAstInfo) Then
'                            ilRet = mBuildXHTInfo(tmAstInfo(llIndex).lAttCode, slSDate, sEndDate, ilPassForm)
'                        End If
                        'Dan 2/13/15 added this as a precaution.  see 7397 (database had cptt for a week it shouldn't have)
                        ReDim tmProgCodeMatch(0 To 0) As PROGCODEMATCH
                        For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1
                            If InStr(1, slPreviousAtts, "," & tmAstInfo(llAst).lAttCode & ",") = 0 Then
                                slPreviousAtts = slPreviousAtts & "," & tmAstInfo(llAst).lAttCode & ","
                                ilRet = mBuildXHTInfo(tmAstInfo(llAst).lAttCode, slSDate, sEndDate, ilPassForm)
                                If ilRet Then
                                    blisReExport = True
                                End If
                            End If
                            '7458
                            If Not myEnt.Add(tmAstInfo(llAst).sFeedDate, tmAstInfo(llAst).lgsfCode, Asts) Then
                                myExport.WriteWarning myEnt.ErrorMessage
                            End If
                        Next llAst
                        '6/19/14: ttp 6944 Separate out events by Program Code ID if a sport event
                        'Dan M 6/27/14  this is done above, so not needed here
'                        slVehicleName = mGetVehicleName(tmAstInfo(llIndex).iVefCode)
'                        slStationName = mGetStationName(tmAstInfo(llIndex).iShttCode)
                        'Call mSetResults("Exporting " & slStationName & ", " & slVehicleName, 0)
                        lacProcessing.Caption = "Exporting " & slVehicleName & ", " & slStationName
                        DoEvents
                        If (ilPassForm = ISCIFORM) And (Not blISCIGame) Then
                            ReDim slLoopProgCodeID(0 To 1) As String
                            slLoopProgCodeID(0) = ""
                        Else
                            ReDim slLoopProgCodeID(0 To 0) As String
                            For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                                If igExportSource = 2 Then DoEvents
                                blSpotOk = True
                                '6082 reset
                                blFailedBecauseNotRegional = False
                                ilAnf = gBinarySearchAnf(tmAstInfo(llAst).iAnfCode)
                                If ilAnf <> -1 Then
                                    If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                                        blSpotOk = False
                                    End If
                                End If
                                If igExportSource = 2 Then DoEvents
                                If (blSpotOk) And ((gDateValue(gAdjYear(tmAstInfo(llAst).sFeedDate)) >= gDateValue(gAdjYear(slSDate))) And (gDateValue(gAdjYear(tmAstInfo(llAst).sFeedDate)) <= gDateValue(gAdjYear(sEndDate)))) Then
                                    '9452
                                    If (tgStatusTypes(gGetAirStatus(tmAstInfo(llAst).iStatus)).iPledged <> 2 Or bmSendNotCarried) Then
                                        If ilPassForm = ISCIFORM Then
                                            slCodeID = mGetProgCode("EVENT", tmAstInfo(llAst).lgsfCode)
                                        Else
                                            slCodeID = mGetProgCode(slProgCodeID, tmAstInfo(llAst).lgsfCode)
                                        End If
                                        If (slCodeID <> slProgCodeID) Or (ilPassForm = ISCIFORM) Then
                                            ilFound = False
                                            For ilLoop = 0 To UBound(slLoopProgCodeID) - 1 Step 1
                                                If slLoopProgCodeID(ilLoop) = slCodeID Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                slLoopProgCodeID(UBound(slLoopProgCodeID)) = slCodeID
                                                ReDim Preserve slLoopProgCodeID(0 To UBound(slLoopProgCodeID) + 1) As String
                                            End If
                                        End If
                                    End If
                                End If
                            Next llAst
                            If UBound(slLoopProgCodeID) = 0 Then
                                ReDim slLoopProgCodeID(0 To 1) As String
                                slLoopProgCodeID(0) = ""
                            End If
                        End If
                        
                        For ilProgCode = 0 To UBound(slLoopProgCodeID) - 1 Step 1
                            For ilMatch = 0 To UBound(tmProgCodeMatch) - 1 Step 1
                                If Trim$(slLoopProgCodeID(ilProgCode)) = Trim$(tmProgCodeMatch(ilMatch).sProgCodeID) Then
                                    tmProgCodeMatch(ilMatch).bMatch = True
                                End If
                            Next ilMatch
                        Next ilProgCode
                        
                        '6/19/14: ttp 6944 Separate out events by Program Code ID if a sport event
                        For ilProgCode = 0 To UBound(slLoopProgCodeID) - 1 Step 1
                            llIndex = LBound(tmAstInfo)
                            slVehicleName = mGetVehicleName(tmAstInfo(llIndex).iVefCode)
                            '10933 added ilAcknowledgeDaylight, but don't use
                            slStationName = mGetStationName(tmAstInfo(llIndex).iShttCode, ilAcknowledgeDaylight)
                            slPrevTransmissionID = ""
                            llODate = -1
                            ilStartHour = -1
                            slPrevHour = ""
                            ilBreakNumber = -1
                            ilPositionNumber = -1
                            slPrevFeedTime = ""
'START OF SPOTS LOOP
                            Do While llIndex < UBound(tmAstInfo)
                                If igExportSource = 2 Then DoEvents
                                blSpotOk = True
                                '6082 reset
                                blFailedBecauseNotRegional = False
                                ilAnf = gBinarySearchAnf(tmAstInfo(llIndex).iAnfCode)
                                If ilAnf <> -1 Then
                                    If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                                        blSpotOk = False
                                    End If
                                End If
                                If igExportSource = 2 Then DoEvents
                                
                                '6/19/14: ttp 6944 Separate out events by Program Code ID if a sport event
                                If (blSpotOk) Then
                                    If slLoopProgCodeID(ilProgCode) <> "" Then
                                        If (ilPassForm = ISCIFORM) Then
                                            If mGetProgCode("EVENT", tmAstInfo(llIndex).lgsfCode) <> slLoopProgCodeID(ilProgCode) Then
                                                blSpotOk = False
                                            End If
                                        Else
                                            If mGetProgCode(slProgCodeID, tmAstInfo(llIndex).lgsfCode) <> slLoopProgCodeID(ilProgCode) Then
                                                blSpotOk = False
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If (blSpotOk) And ((gDateValue(gAdjYear(tmAstInfo(llIndex).sFeedDate)) >= gDateValue(gAdjYear(slSDate))) And (gDateValue(gAdjYear(tmAstInfo(llIndex).sFeedDate)) <= gDateValue(gAdjYear(sEndDate)))) Then
                                    'Check if Date cghanged
                                    If llODate <> gDateValue(gAdjYear(tmAstInfo(llIndex).sFeedDate)) Then
                                        'If llODate <> -1 Then
                                        '    csiXMLData "CT", "Site", ""
                                        '    csiXMLData "CT", "Sites", ""
                                        '    ilRet = csiXMLWrite(1)
                                        'End If
                                            '5/7/15 Dan always B
'                                        If llODate <> -1 Then
'                                            If slSendProgAndSiteBy = "D" Then
'                                                mXMLSiteTags ilPassForm, slXDReceiverID, slTransmissionID, slUnitHB, slUnitHBP, slVefCode5, 0
'                                            End If
                                            '5/7/15 Dan we don't use this, so comment out
'                                            If slSendWriteBy = "D" Then
'                                                If udcCriteria.XExportType(0, "V") = vbChecked Then
'                                                    '6979
'                                                    mAddSurroundingElement ilPassForm, False
'                                                    '6635 same as below, but now uses Jeff's error log
'                                                    '7508
'                                                    'If Not mSendAndWriteReturn("Spot Insertions") Then
'                                                    If Not mSendAndTestReturn(XDSType.Insertions, slDoNotReturn) Then
'                                                        blRetStatus = False
'    '                                                    If llTotalExport > 0 Then
'    '                                                        llTotalExport = llTotalExport - 1
'    '                                                    End If
'                                                        If bmIsError Then
'                                                            Exit Function
'                                                        End If
'                                                       ' gLogMsg "Station being sent: " & slStationName & " , Vehicle being sent: " & slVehicleName, smPathForgLogMsg, False
'                                                        myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName
'                                                        Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", MESSAGERED)
'    '                                                ilRet = csiXMLWrite(1)
'    '                                                If ilRet <> True Then
'    '                                                     blRetStatus = False
'    '                                                     'dan 3/13/13 remove from count
'    '                                                     If llTotalExport > 0 Then
'    '                                                         llTotalExport = llTotalExport - 1
'    '                                                     End If
'    '                                                    ilRet = csiXMLStatus(tlXmlStatus)
'    '                                                    '5896  log error, change name of log to show user in display
'    '                                                     If mIsXmlError(tlXmlStatus.sStatus) Then
'    '                                                         Exit Function
'    '                                                     End If
'
'    '                                                    '11/26/12: Continue to next CPTT
'    '                                                    'imTerminate = True
'    '                                                    'imExporting = False
'    '                                                    blRetStatus = False
'    '                                                    ilRet = csiXMLStatus(tlXmlStatus)
'    '                                                    '5896 3/13/13
'    '
'    '                                                    Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", RGB(155, 0, 0))
'    '                                                    '5896
'    '                                                    If gIsNull(tlXmlStatus.sStatus) Then
'    '                                                        gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR", "XDigitalExportLog.Txt", False
'    '                                                    Else
'    '                                                        gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
'    '                                                    End If
'    '                                                    'gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " Error: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
'    '                                                    'Exit Function
'                                                        'dan 3/13/13 remove from count
'                                                        '6689 remove
'                                                        'Exit Do
'                                                        '7236 not implemented here
''                                                    ' Dan M 7/13/12 add UpdateLastExportDate, only when transmitting
''                                                         '7236 if not a successful send, don't allow to update later
''                                                        blPartialExportMayContinue = False
'                                                   Else
'                                                        '7236 don't update partial as its too complicated. grab when done with vehicle
'                                                        'ilRet = mUpdateXHT()
'                                                        If udcCriteria.XGenType(0) Then
'                                                            gUpdateLastExportDate imVefCode, slEDate  'or sledate?
'                                                        End If
'                                                    End If
'                                                    '11/4/11: HB was counting breaks, count move up in llIndexLoop
'                                                    'llTotalExport = llTotalExport + 1
'                                                End If
'                                            End If
'                                        End If
                                        ilFirstCue = True
                                        llODate = gDateValue(gAdjYear(tmAstInfo(llIndex).sFeedDate))
                                        slFeedTime = Format$(tmAstInfo(llIndex).sFeedTime, "hh:mm:ss")
                                        ilStartHour = Val(Left(slFeedTime, 2))
                                        slPrevHour = ""
                                        ilBreakNumber = -1
                                        ilPositionNumber = -1
                                        slPrevFeedTime = ""
                                        ilSeqNo = 1
                                        slTransmissionID = Format$(llODate, "yyyymmdd")
                                        If slPrevTransmissionID <> slTransmissionID Then
                                            slPrevTransmissionID = slTransmissionID
                                            
                                            ' COMMENTED OUT THE DELETE COMMAND
        '                                    ' Due to a change on XDigital side, the Sites tag must not be sent
        '                                    ' with the delete command.
        '                                    Call csiXMLSetMethod("SetInsertions", "", slTransmissionID, "")
        '                                    csiXMLData "OT", "Deletes", ""
        '                                    csiXMLData "CA", "Delete TransmissionID='" & slTransmissionID & "'", ""
        '                                    csiXMLData "CT", "Deletes", ""
        '                                    ilRet = csiXMLWrite(1)
        '                                    'MsgBox "Step 1 - Results = " & ilRet
        '                                    Call csiXMLSetMethod("SetInsertions", "Sites", slTransmissionID, "")
        '                                    If ilRet <> True Then
        '                                        'MsgBox "Step 2 - Going into failed section"
        '                                        imTerminate = True
        '                                        imExporting = False
        '                                        Call mSetResults("Export Failed. Error=1", RGB(155, 0, 0))
        '                                        'MsgBox "Step 3 Getting error message"
        '                                        ilRet = csiXMLStatus(tlXMLStatus)
        '                                        'MsgBox "Step 5. Error message being written to file " & sgMsgDirectory & "\XDigitalExportLog.Txt"
        '                                        'MsgBox "Step 4. Error Msg = " & tlXMLStatus.sStatus
        '                                        gLogMsg "ERROR: " & tlXMLStatus.sStatus, "XDigitalExportLog.Txt", False
        '                                        Exit Function
        '                                    End If
        '                                    'MsgBox "Delete was ok!"
        '
                                            ' END OF DELETE COMMENTED OUT
                                        End If
                                    End If
                                    slFeedTime = Format$(tmAstInfo(llIndex).sFeedTime, "hh:mm:ss")
                                    If slFeedTime <> slPrevFeedTime Then
                                        If smMidnightBasedHours = "Y" Then
                                            slHour = Trim$(Str$(Val(Left(slFeedTime, 2))))
                                        Else
                                            slHour = Trim$(Str$(Val(Left(slFeedTime, 2)) - ilStartHour + 1))
                                        End If
                                        If slHour <> slPrevHour Then
                                            ilBreakNumber = 1
                                        Else
                                            ilBreakNumber = ilBreakNumber + 1
                                        End If
                                        ilPositionNumber = 1
                                        slPrevHour = slHour
                                        slPrevFeedTime = slFeedTime
                                    Else
                                        ilPositionNumber = ilPositionNumber + 1
                                    End If
                                    '6217
                                    If smMidnightBasedHours = "Y" Then
                                        If Len(slHour) = 1 Then
                                            slHour = "0" & slHour
                                        End If
                                    End If
                                    slHB = "H" & slHour & "B" & Trim$(Str$(ilBreakNumber))
                                    slHBP = "H" & slHour & "B" & Trim$(Str$(ilBreakNumber)) & "P" & Trim$(Str$(ilPositionNumber))
                                    slUnitHB = slHour
                                    If Len(slUnitHB) = 1 Then
                                        slUnitHB = "0" & slUnitHB
                                    End If
                                    slStr = Trim$(Str$(ilBreakNumber))
                                    Do While Len(slStr) < 3
                                        slStr = "0" & slStr
                                    Loop
                                    slUnitHB = slUnitHB & slStr
                                    slStr = Trim$(Str$(ilPositionNumber))
                                    If Len(slStr) = 1 Then
                                        slStr = "0" & slStr
                                    End If
                                    slUnitHBP = slUnitHB & slStr
                                    slRotStartDT = ""
                                    slRotEndDT = ""
                                    ''Check if regional copy defined with spots within same avail
                                    ''ilRet = gGetRegionCopy(tmAstInfo(llIndex).iShttCode, tmAstInfo(llIndex).lSdfCode, tmAstInfo(llIndex).iVefCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                                    'ilRet = gGetRegionCopy(tmAstInfo(llIndex), slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                                    'If ilRet Then
                                    If igExportSource = 2 Then DoEvents
                                    '9452 - allow iPledged = 2 if 'Treat Pledge Status of "Not Carried" same as "Aired"'
                                    'If (tgStatusTypes(gGetAirStatus(tmAstInfo(llIndex).iStatus)).iPledged <> 2) Then
                                    ilPledgeStatus = tgStatusTypes(gGetAirStatus(tmAstInfo(llIndex).iStatus)).iPledged
                                    If (ilPledgeStatus <> 2 Or bmSendNotCarried) Then
                                        If ilPledgeStatus = 2 Then
'                                            ReDim tlAstInfo(0 To 0)
'                                            tlAstInfo(0) = tmAstInfo(llIndex)
                                            gGetAndAssignRegionToAst hmAst, tmAstInfo(llIndex)
                                        End If
                                        If tmAstInfo(llIndex).iRegionType > 0 Then
                                            slRISCI = Trim$(tmAstInfo(llIndex).sRISCI)
                                            slRCreative = gXMLNameFilter(Trim$(tmAstInfo(llIndex).sRCreativeTitle))
                                            llRCrfCode = tmAstInfo(llIndex).lRCrfCode
                                            ilIncludeSpot = True
                                            ilRegionExist = True
                                            '11/14/11: Use Feed date instead of the rotation date
                                            '6/6/11: Allow multi-days, set StartDate and EndDate for HB and HBP to Feed date
                                            'If ilPassForm = 0 Or ilPassForm = 1 Then
                                            If ilPassForm = ISCIFORM Then
                                                If igExportSource = 2 Then DoEvents
                                                SQLQuery = "Select * from CRF_Copy_Rot_Header"
                                                SQLQuery = SQLQuery & " Where (crfCode = " & llRCrfCode & ")"
                                                Set crf_rst = gSQLSelectCall(SQLQuery)
                                                If crf_rst.EOF Then
                                                    ' slRotStartDT = slTransmissionID & " " & "00:00:00"
                                                    ' slRotEndDT = slTransmissionID & " " & "23:59:59"
                                                    slRotStartDT = Format$(llODate, "mm/dd/yyyy") & " " & "00:00:00"
                                                    slRotEndDT = Format$(llODate, "mm/dd/yyyy") & " " & "23:59:59"
                                                Else
                                                    'slRotStartDT = Format$(crf_rst!crfStartDate, "mm/dd/yyyy") & " " & Format$(crf_rst!crfStartTime, "hh:mm:ss")
                                                    slRotTime = Format$(crf_rst!crfEndTime, "hh:mm:ss")
                                                    If slRotTime = "00:00:00" Then
                                                        slRotTime = "23:59:59"
                                                    End If
                                                    'slRotEndDT = Format$(crf_rst!crfEndDate, "mm/dd/yyyy") & " " & slRotTime
                                                    slRotStartDT = Format$(llODate, "mm/dd/yyyy") & " " & Format$(crf_rst!crfStartTime, "hh:mm:ss")
                                                    slRotEndDT = Format$(llODate, "mm/dd/yyyy") & " " & slRotTime
                                                End If
                                            Else
                                                slRotStartDT = Format$(gAdjYear(tmAstInfo(llIndex).sFeedDate), "mm/dd/yyyy")
                                                slRotEndDT = slRotStartDT
                                                slRotStartDT = slRotStartDT & " " & "00:00:00"
                                                slRotEndDT = slRotEndDT & " " & "23:59:59"
                                            End If
                                        'not regional
                                        Else
                                            '6082 lose this! Let if fall into later code, and mark as false there to mean don't write out
                                            'ilIncludeSpot = False
                                            '7/9/13: Transparent file generatrion is independent of the Unit ID setting
                                            'If smUnitIdByAstCode = "Y" Then
                                            If smGenTransparency = "Y" Then
                                                ilIncludeSpot = True
                                            Else
                                                ilIncludeSpot = False
                                            End If
                                            ilRegionExist = False
                                            If udcCriteria.XSpots(ALLSPOTS) Then
                                                ilIncludeSpot = True
                                                '11/14/11: Use Feed date instead of the rotation date
                                                '6/6/11: Allow multi-days, set StartDate and EndDate for HB and HBP to Feed date
                                                If ilPassForm = ISCIFORM Then
                                                    llCrfCode = gGetSdfCrfCode(tmAstInfo(llIndex).lSdfCode)
                                                    If llCrfCode <> 0 Then
                                                        SQLQuery = "Select * from CRF_Copy_Rot_Header"
                                                        SQLQuery = SQLQuery & " Where (crfCode = " & llCrfCode & ")"
                                                        Set crf_rst = gSQLSelectCall(SQLQuery)
                                                        If crf_rst.EOF Then
                                                            'slRotStartDT = slTransmissionID & " " & "00:00:00"
                                                            'slRotEndDT = slTransmissionID & " " & "23:59:59"
                                                            slRotStartDT = Format$(llODate, "mm/dd/yyyy") & " " & "00:00:00"
                                                            slRotEndDT = Format$(llODate, "mm/dd/yyyy") & " " & "23:59:59"
                                                        Else
                                                            'slRotStartDT = Format$(crf_rst!crfStartDate, "mm/dd/yyyy") & " " & Format$(crf_rst!crfStartTime, "hh:mm:ss")
                                                            slRotTime = Format$(crf_rst!crfEndTime, "hh:mm:ss")
                                                            If slRotTime = "00:00:00" Then
                                                                slRotTime = "23:59:59"
                                                            End If
                                                            'slRotEndDT = Format$(crf_rst!crfEndDate, "mm/dd/yyyy") & " " & slRotTime
                                                            slRotStartDT = Format$(llODate, "mm/dd/yyyy") & " " & Format$(crf_rst!crfStartTime, "hh:mm:ss")
                                                            slRotEndDT = Format$(llODate, "mm/dd/yyyy") & " " & slRotTime
                                                        End If
                                                    End If
                                                Else
                                                    slRotStartDT = Format$(gAdjYear(tmAstInfo(llIndex).sFeedDate), "mm/dd/yyyy")
                                                    slRotEndDT = slRotStartDT
                                                    slRotStartDT = slRotStartDT & " " & "00:00:00"
                                                    slRotEndDT = slRotEndDT & " " & "23:59:59"
                                                End If ' which form
                                            '6082  chose regional but not regional copy.  Need this!
                                            Else
                                                blFailedBecauseNotRegional = True
                                            End If ' chose 'all'
                                        End If 'regional
                                    Else
                                        ilIncludeSpot = False
                                    End If  'pledged type ok
                                    If igExportSource = 2 Then DoEvents
                                    'don't include if missing id
                                    If ilIncludeSpot Then
                                        If ((ilPassForm = ISCIFORM) And (Val(slStationID) <= 0)) Then
                                            ilIncludeSpot = False
                                            'Station
                                            Call mSetResults(slStationName & " Missing XDS Station ID", MESSAGEBLACK)
                                            If Not bmMgsPrevExisted Then
                                                'gLogMsg slStationName & " Missing XDS Station ID", smPathForgLogMsg, False
                                                myExport.WriteWarning slStationName & " Missing XDS Station ID"
                                            End If
                                        ElseIf ((ilPassForm <> ISCIFORM) And (Val(slXDReceiverID) <= 0)) Then
                                            ilIncludeSpot = False
                                            'Station and Vehicle
                                            Call mSetResults(slStationName & " airing " & slVehicleName & " Missing X-Digital Station ID", MESSAGEBLACK)
                                            If Not bmMgsPrevExisted Then
                                               ' gLogMsg slStationName & " airing " & slVehicleName & " Missing X-Digital Station ID", smPathForgLogMsg, False
                                                myExport.WriteWarning slStationName & " airing " & slVehicleName & " Missing X-Digital Station ID"
                                            End If
                                        End If
                                    End If
                                    If igExportSource = 2 Then DoEvents
                                    If ilIncludeSpot Then
                                        '6082 I need to come in here even if not writing out--user chose regional and not regional.
                                        ' used to set ilIncludeSpot to 0 in that case.  Set it to that now so won't write out.
                                        If blFailedBecauseNotRegional Then
                                            ilIncludeSpot = False
                                        End If
                                        If slRotStartDT = "" Then
                                            slRotStartDT = Format$(llODate, "mm/dd/yyyy") & " " & "00:00:00"
                                        End If
                                        If slRotEndDT = "" Then
                                            slRotEndDT = Format$(llODate, "mm/dd/yyyy") & " " & "23:59:59"
                                        End If
                                        'Get Rotation
                                        slSeqNo = ilSeqNo
                                        Do While Len(slSeqNo) < 4
                                            slSeqNo = "0" & slSeqNo
                                        Loop
                                        ilSeqNo = ilSeqNo + 1
                                        If ilPassForm = ISCIFORM Then
                                            '7/24/14: add ProgCodeID to search
                                            ''7/6/15: Set flag indicating if XML commands should be executed
                                            'bmAllowXMLCommands = mSetXMLCommandPass0(llIndex, slStationID, slTransmissionID, slSeqNo, slVefCode5)
                                            '10/9/14: astInfo().sISCI references the Blackout copy. Replace with Generic
                                            If tmAstInfo(llIndex).iRegionType = 2 Then
                                                SQLQuery = "SELECT lstBkoutLstCode"
                                                SQLQuery = SQLQuery & " FROM lst"
                                                SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tmAstInfo(llIndex).lLstCode)
                                                Set lst_rst = gSQLSelectCall(SQLQuery)
                                                If Not lst_rst.EOF Then
                                                    llLstCode = lst_rst!lstBkoutLstCode
                                                    ilRet = mGetSdfCopy(llLstCode, tmAstInfo(llIndex).sCart, tmAstInfo(llIndex).sISCI, slCreative, slLen)
                                                End If
                                            End If
                                            '9114  slUnitIsci
                                            If smUnitIdByAstCodeForISCI = "Y" Then
                                                slUnitIsci = Trim$(Str$(tmAstInfo(llIndex).lCode))
                                                Do While Len(slUnitIsci) < 9
                                                    slUnitIsci = "0" & slUnitIsci
                                                Loop
                                            Else
                                                slUnitIsci = slTransmissionID & slSeqNo & slVefCode5
                                            End If
                                            '9818
                                            If imSharedHeadEndIsci > 0 Then
                                                slUnitIsci = imSharedHeadEndIsci & slUnitIsci
                                            End If
                                            '9114, and don't see the need for the 'if' below
'                                            If slLoopProgCodeID(ilProgCode) <> "" Then
'                                                bmAllowXMLCommands = mSetXMLCommandPass0(llIndex, slStationID, slTransmissionID, slSeqNo, slVefCode5, slLoopProgCodeID(ilProgCode))
'                                            Else
'                                                bmAllowXMLCommands = mSetXMLCommandPass0(llIndex, slStationID, slTransmissionID, slSeqNo, slVefCode5, "")
'                                            End If
                                            If slLoopProgCodeID(ilProgCode) <> "" Then
                                                bmAllowXMLCommands = mSetXMLCommandPass0(llIndex, slStationID, slTransmissionID, slUnitIsci, slLoopProgCodeID(ilProgCode))
                                            Else
                                                bmAllowXMLCommands = mSetXMLCommandPass0(llIndex, slStationID, slTransmissionID, slUnitIsci, "")
                                            End If
                                            '6082
                                            If ilIncludeSpot Then
                                                'csiXMLData "OT", "Sites", ""
                                                'Dan 7/14/14 first time writing to xml
                                                '6979 had to split for deletions
                                                If bmWroteTopElement = False Then
                                                    mAddSurroundingElement ilPassForm, True
                                                    bmWroteTopElement = True
                                                End If
                                                mCSIXMLData "OT", "Site", "SiteId = " & """" & slStationID & """"
                                                '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
                                                'mCSIXMLData "OT", "Insert", "UnitID=" & """" & slTransmissionID & slSeqNo & """"
                                                '9114
                                               ' mCSIXMLData "OT", "Insert", "UnitID=" & """" & slTransmissionID & slSeqNo & slVefCode5 & """"
                                                mCSIXMLData "OT", "Insert", "UnitID=" & """" & slUnitIsci & """"
                                                mCSIXMLData "CA", "AiringNetwork", "NetworkId=" & """" & tgVpfOptions(ilVpf).iInterfaceID & """"
                                                '8357
                                                '9114
'                                                mRetainSiteIdAndUnitId slStationID, slTransmissionID & slSeqNo & slVefCode5
                                                mRetainSiteIdAndUnitId slStationID, slUnitIsci
                                                '6796  Dan add program code(program Number with 6835) here...only if game. Passing "EVENT" forces it to get game
                                                If blISCIGame Then
                                                    slEventProgCodeID = mGetProgCode("EVENT", tmAstInfo(llIndex).lgsfCode)
                                                    '6835
                                                    mCSIXMLData "CD", "ProgramNumber", slEventProgCodeID
                                                End If
                                            End If
                                            llIndexStart = llIndex
                                            llIndexEnd = llIndex
                                        Else
                                            '7/6/15: Set flag indicating if XML commands should be executed
                                            If ilPassForm = HBPFORM Then
                                                '7/24/14: add ProgCodeID to search
                                                'bmAllowXMLCommands = mSetXMLCommandPass1(llIndex, slXDReceiverID, slTransmissionID, slUnitHBP, slVefCode5)
'                                                If slLoopProgCodeID(ilProgCode) <> "" Then
'                                                    bmAllowXMLCommands = mSetXMLCommandPass1(llIndex, slXDReceiverID, slTransmissionID, slUnitHBP, slVefCode5, slLoopProgCodeID(ilProgCode))
'                                                Else
'                                                    bmAllowXMLCommands = mSetXMLCommandPass1(llIndex, slXDReceiverID, slTransmissionID, slUnitHBP, slVefCode5, slProgCodeID)
'                                                End If
                                                If slLoopProgCodeID(ilProgCode) <> "" Then
                                                    bmAllowXMLCommands = mSetXMLCommandPassHBP(llIndex, slXDReceiverID, slTransmissionID, slUnitHBP, slVefCode5, slLoopProgCodeID(ilProgCode))
                                                Else
                                                    bmAllowXMLCommands = mSetXMLCommandPassHBP(llIndex, slXDReceiverID, slTransmissionID, slUnitHBP, slVefCode5, slProgCodeID)
                                                End If
                                            Else
                                            '8236  adjust slProgCodeID and slHB if slProgCodeID = "Event" and not a game
                                                blLeaveCueAlone = False
                                                If ilProgCode = 0 Then
                                                    'still testing: if not blank, it's a game
                                                    If slLoopProgCodeID(ilProgCode) = "" Then
                                                        If blIsNonGameEvent Then
                                                            'for below, when writing
                                                            blLeaveCueAlone = True
                                                            '8299 moved up in code
'                                                            If Not blCEFIsOpen Then
'                                                                mOpenCEFFile
'                                                                blCEFIsOpen = True
'                                                            End If
                                                            '10933 building the slEventZones array
                                                            If blIsEventZone Then
                                                                slProgCodeID = mParseEventIdForZone(tmAstInfo(llIndex).lEvtIDCefCode, slProgCodeID, slEventZones, dlCueAndCodes)
                                                                '11015
                                                                ReDim slEventIdsIscis(0)
                                                            Else
                                                            '8279
                                                                slHB = mParseEventId(tmAstInfo(llIndex).lEvtIDCefCode, slProgCodeID, slHB, tlEventIds, dlCueAndCodes)
                                                                ReDim slEventIdsIscis(0)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                llIndexStart = llIndex - (ilPositionNumber - 1)
                                                llIndexEnd = llIndexStart + 1
                                                Do While llIndexEnd < UBound(tmAstInfo)
                                                    If igExportSource = 2 Then DoEvents
                                                    If (gDateValue(gAdjYear(tmAstInfo(llIndexStart).sFeedDate)) = gDateValue(gAdjYear(tmAstInfo(llIndexEnd).sFeedDate))) Then
                                                        If gTimeToLong(tmAstInfo(llIndexStart).sFeedTime, False) = gTimeToLong(tmAstInfo(llIndexEnd).sFeedTime, False) Then
                                                            llIndexEnd = llIndexEnd + 1
                                                        Else
                                                            Exit Do
                                                        End If
                                                    Else
                                                        Exit Do
                                                    End If
                                                Loop
                                                llIndexEnd = llIndexEnd - 1
                                                '7/24/14: add ProgCodeID to search
                                                'bmAllowXMLCommands = mSetXMLCommandPass2(llIndexStart, llIndexEnd, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5)
                                                '10021
'                                                If slLoopProgCodeID(ilProgCode) <> "" Then
'                                                    bmAllowXMLCommands = mSetXMLCommandPass2(llIndexStart, llIndexEnd, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5, slLoopProgCodeID(ilProgCode))
'                                                Else
'                                                    bmAllowXMLCommands = mSetXMLCommandPass2(llIndexStart, llIndexEnd, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5, slProgCodeID)
'                                                End If
                                                If slLoopProgCodeID(ilProgCode) <> "" Then
                                                    bmAllowXMLCommands = mSetXMLCommandPassHB(llIndexStart, llIndexEnd, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5, slLoopProgCodeID(ilProgCode))
                                                Else
                                                    bmAllowXMLCommands = mSetXMLCommandPassHB(llIndexStart, llIndexEnd, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5, slProgCodeID)
                                                End If
                                            End If
                                            If (slSendProgAndSiteBy = "B") Or ((slSendProgAndSiteBy = "D") And (ilFirstCue)) Then
                                                slEventProgCodeID = mGetProgCode(slProgCodeID, tmAstInfo(llIndex).lgsfCode)
                                                '6082
                                                If ilIncludeSpot Then
                                                    mGetRotDTForGames slProgCodeID, tmAstInfo(llIndex).lgsfCode, slRotStartDT, slRotEndDT
                                                    'Dan 7/14/14 first time writing to xml
                                                    '6979 had to split for deletions
                                                    If bmWroteTopElement = False Then
                                                        mAddSurroundingElement ilPassForm, True
                                                        bmWroteTopElement = True
                                                    End If
                                                    mCSIXMLData "OT", "Insert", ""
                                                    mCSIXMLData "CD", "ProgramCode", slEventProgCodeID
                                                    '5760 Dan
                                                    If Len(slVehicleName) > 0 Then
                                                        mCSIXMLData "CD", "ProgramName", gXMLNameFilter(slVehicleName)
                                                    End If
                                                    ilFirstCue = False
                                                End If
                                            End If
                                            If ilPassForm = HBPFORM Then
                                                '6082
                                                If ilIncludeSpot Then
                                                    If (tmAstInfo(llIndex).iRegionType > 0) And (tmAstInfo(llIndex).lIrtCode > 0) And (Trim$(tmAstInfo(llIndex).sReplacementCue) <> "") Then
                                                        mCSIXMLData "CD", "Cue", Trim$(tmAstInfo(llIndex).sReplacementCue)
                                                    Else
                                                        mCSIXMLData "CD", "Cue", slEventProgCodeID & slHBP
                                                    End If
                                                End If
                                                llIndexStart = llIndex
                                                llIndexEnd = llIndex
                                            ElseIf ilPassForm = HBFORM Then
                                                '6082
                                                If ilIncludeSpot Then
                                                    'mCSIXMLData "CD", "Cue", slEventProgCodeID & slHB
                                                    '8236 added first 'if'
                                                    If blLeaveCueAlone Then
                                                        '10933 alter slHB by zone
                                                        If blIsEventZone Then
                                                            If slProgCodeID <> "EVENT" Then
                                                                slZoneAdjustment = myZoneAndDSTHelper.ZoneByMethod(Trim$(tmAstInfo(llIndex).sFeedDate), Trim$(tmAstInfo(llIndex).sFeedTime))
                                                                slHB = slEventZones(slZoneAdjustment)
                                                            End If
                                                        End If
                                                        mCSIXMLData "CD", "Cue", slHB
                                                    ElseIf (tmAstInfo(llIndex).iRegionType > 0) And (tmAstInfo(llIndex).lIrtCode > 0) And (Trim$(tmAstInfo(llIndex).sReplacementCue) <> "") Then
                                                        mCSIXMLData "CD", "Cue", Trim$(tmAstInfo(llIndex).sReplacementCue)
                                                    Else
                                                        mCSIXMLData "CD", "Cue", slEventProgCodeID & slHB
                                                    End If
                                                End If
                                                llIndexStart = llIndex - (ilPositionNumber - 1)
                                                'Set Length
                                                '6191 comment out and get length different way
                                                'ilRet = mGetSdfCopy(tmAstInfo(llIndexStart).lLstCode, slCart, slISCI, slCreative, slLength)
                                                slLength = tmAstInfo(llIndex).iLen
                                                '6082 - dan to do spot length?
                                                llIndexEnd = llIndexStart + 1
                                                Do While llIndexEnd < UBound(tmAstInfo)
                                                    If igExportSource = 2 Then DoEvents
                                                    If (gDateValue(gAdjYear(tmAstInfo(llIndexStart).sFeedDate)) = gDateValue(gAdjYear(tmAstInfo(llIndexEnd).sFeedDate))) Then
                                                        If gTimeToLong(tmAstInfo(llIndexStart).sFeedTime, False) = gTimeToLong(tmAstInfo(llIndexEnd).sFeedTime, False) Then
                                                            '6191
                                                            'ilRet = mGetSdfCopy(tmAstInfo(llIndexEnd).lLstCode, slCart, slISCI, slCreative, slLen)
                                                            slLen = tmAstInfo(llIndexEnd).iLen
                                                            slLength = Trim$(Str$(Val(slLength) + Val(slLen)))
                                                            llIndexEnd = llIndexEnd + 1
                                                        Else
                                                            Exit Do
                                                        End If
                                                    Else
                                                        Exit Do
                                                    End If
                                                Loop
                                                llIndexEnd = llIndexEnd - 1
                                            End If
                                        End If
                                        If igExportSource = 2 Then DoEvents
                                        '6082
                                        If ilIncludeSpot Then
                                            '9629
                                            If ilLocalAdj <> 0 Then
                                                slTempDate = Format(slRotStartDT, "mm/dd/yyyy")
                                                slTempTime = Format(slRotStartDT, "hh:mm:ss")
                                                mAdjustToHeadendZone ilLocalAdj, slTempDate, slTempTime
                                                slRotStartDT = slTempDate & " " & Format(slTempTime, "hh:mm:ss")
                                                slTempDate = Format(slRotEndDT, "mm/dd/yyyy")
                                                slTempTime = Format(slRotEndDT, "hh:mm:ss")
                                                mAdjustToHeadendZone ilLocalAdj, slTempDate, slTempTime
                                                slRotEndDT = slTempDate & " " & Format(slTempTime, "hh:mm:ss")
                                            End If
                                            mCSIXMLData "CD", "StartDate", slRotStartDT
                                            mCSIXMLData "CD", "EndDate", slRotEndDT
                                            mCSIXMLData "CD", "TransmissionID", slTransmissionID
                                        End If
                                        
                                        'Get Short Title
                                        For llIndexLoop = llIndexStart To llIndexEnd Step 1
                                            If igExportSource = 2 Then DoEvents
                                            llAstCode = tmAstInfo(llIndexLoop).lCode
                                            '6191
    '                                        slShortTitle = gGetShortTitle(tmAstInfo(llIndexLoop).lSdfCode)
                                            If (sgSpfUseProdSptScr <> "P") Then
                                                llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                '7429
                                                If llAdf = -1 Then
                                                    gPopAdvertisers
                                                    llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                End If
                                                If llAdf <> -1 Then
                                                    slShortTitle = Trim$(Left(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6))
                                                    If slShortTitle = "" Then
                                                        slShortTitle = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
                                                    End If
                                                '7429
                                                ElseIf smAddAdvtToISCI = "Y" Or ilPassForm = ISCIFORM Then
                                                    blRetStatus = False
                                                    slShortTitle = "ADV_MISSING-" & tmAstInfo(llIndexLoop).iAdfCode
                                                    mSetResults "Advertiser missing!  Can't find advertister with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode, MESSAGERED
                                                    If ilPassForm = ISCIFORM Then
                                                        myExport.WriteWarning "Advertiser missing. Can't find advertiser with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode & ".  Filename will not be written correctly.  Search for 'ADV_MISSING' to see the issue.", False
                                                    Else
                                                        myExport.WriteWarning "Advertiser missing. Can't find advertiser with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode & ".  ISCI will not be written correctly.  Search for 'ADV_MI' to see the issue.", False
                                                    End If
                                                End If
                                                slShortTitle = slShortTitle & "," & Trim$(tmAstInfo(llIndexLoop).sProd)
                                            Else
                                                slShortTitle = gGetShortTitle(tmAstInfo(llIndexLoop).lSdfCode)
                                            End If
                                            '12/27/13: Replace ShortTitle by advertiser name so that the Traffic Cart Export (Copy Export) will match
                                            If ilPassForm = ISCIFORM Then
                                                '6744 Dan "CUMULUS" is never returned.  Must've wanted when Cumulus on double head end
                                               ' If InStr(1, UCase(Trim(slSection)), "CUMULUS", vbBinaryCompare) > 0 Then
                                               ' look for "-CU" at top
                                               ' If InStr(1, UCase(Trim(slSection)), "-CU", vbBinaryCompare) > 0 Then
                                                If blisCUInHeader Then
                                                    llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                    If llAdf <> -1 Then
                                                        slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                                                    End If
                                                End If
                                            End If
                                            slShortTitle = UCase$(gFileNameFilter(slShortTitle))
                                            'Get National Copy
                                            If ilPassForm = HBFORM Then
                                                ilRegionExist = False
                                            End If
                                            If igExportSource = 2 Then DoEvents
                                            If tmAstInfo(llIndexLoop).iRegionType = 2 Then
                                                If ilPassForm = HBFORM Then
                                                    ilRegionExist = True
                                                    slRISCI = Trim$(tmAstInfo(llIndexLoop).sRISCI)
                                                    slRCreative = gXMLNameFilter(Trim$(tmAstInfo(llIndexLoop).sRCreativeTitle))
                                                End If
                                                SQLQuery = "SELECT lstBkoutLstCode"
                                                SQLQuery = SQLQuery & " FROM lst"
                                                SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tmAstInfo(llIndexLoop).lLstCode)
                                                Set rst = gSQLSelectCall(SQLQuery)
                                                If Not rst.EOF Then
                                                    llLstCode = rst!lstBkoutLstCode
                                                    ilRet = mGetOrigSpotShortTitle(ilPassForm, slSection, llLstCode, slShortTitle)
                                                    ilRet = mGetSdfCopy(llLstCode, slCart, slISCI, slCreative, slLen)
                                                Else
                                                    slCart = ""
                                                    slISCI = ""
                                                    slCreative = ""
                                                End If
                                            ElseIf tmAstInfo(llIndexLoop).iRegionType = 1 Then
                                                If ilPassForm = ISCIFORM Then
                                                    ilRet = mGetSdfCopy(tmAstInfo(llIndexLoop).lLstCode, slCart, slISCI, slCreative, slLen)
                                                End If
                                                If ilPassForm = HBPFORM Then
                                                    slLen = Trim$(Str$(tmAstInfo(llIndexLoop).iLen))
                                                End If
                                                If ilPassForm = HBFORM Then
                                                    ilRegionExist = True
                                                    slRISCI = Trim$(tmAstInfo(llIndexLoop).sRISCI)
                                                    slRCreative = gXMLNameFilter(Trim$(tmAstInfo(llIndexLoop).sRCreativeTitle))
                                                End If
                                            Else
                                                '6191 replace, and use cpf to get creative title
                                                'ilRet = mGetSdfCopy(tmAstInfo(llIndexLoop).lLstCode, slCart, slISCI, slCreative, slLen)
                                                slLen = tmAstInfo(llIndexLoop).iLen
                                                slISCI = Trim$(tmAstInfo(llIndexLoop).sISCI)
                                                slCart = tmAstInfo(llIndexLoop).sCart
                                                slCreative = Trim$(mGetCreative(tmAstInfo(llIndexLoop).lCpfCode))
                                            End If
                                            If igExportSource = 2 Then DoEvents
                                            'RD_ to ISCI becuase ABC and Disney are using the same database and need to not have a confit of names
                                            'slISCI = "RD_" & UCase$(mFileNameFilter(slISCI))
                                            slISCI = slISCIPrefix & UCase$(gFileNameFilter(Trim$(slISCI)))
                                            If (smAddAdvtToISCI = "Y") And (ilPassForm <> 0) Then
                                                '2/7/13: Use only advertiser abbreviation because ShortTitle with product is too long (over 32)
                                                llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                If llAdf <> -1 Then
                                                    slShortTitle = Trim$(Left(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6))
                                                    If slShortTitle = "" Then
                                                        slShortTitle = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
                                                    End If
                                                '7429
'                                                Else
'                                                    slShortTitle = Trim$(Left(tmAstInfo(llIndexLoop).sProd, 6))
                                                End If
                                                slShortTitle = UCase$(gFileNameFilter(slShortTitle))
                                                slISCI = Left$(slShortTitle, 6) & "(" & slISCI & ")"
                                            End If
                                            '11063
                                            If bmCartReplaceISCI Then
                                                slISCI = Trim$(slCart)
                                                'temp
'                                                If slISCI = "" Then
'                                                    slISCI = "MISSING"
'                                                End If
                                            End If
                                            '6082
                                            If ilIncludeSpot Then
                                                If ilPassForm = ISCIFORM Then
                                                    mCSIXMLData "OT", "NationalSpot", ""
                                                    mCSIXMLData "CD", "ISCI", slISCI
                                                    '7496
                                                    'mCSIXMLData "CD", "FileName", slShortTitle & "(" & slISCI & ")" & ".MP2"
                                                    mCSIXMLData "CD", "FileName", slShortTitle & "(" & slISCI & ")" & UCase(sgAudioExtension)
                                                    mCSIXMLData "CD", "Duration", slLen
                                                    mCSIXMLData "CT", "NationalSpot", ""
                                                    If Not ilRegionExist Then
                                                        mSaveFD slVefCode, slStationID, slISCI, slCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), slShortTitle, slXDXMLForm, ""
                                                    End If
                                                ElseIf ilPassForm = HBPFORM Then
                                                    If Not ilRegionExist And udcCriteria.XSpots(ALLSPOTS) Then
                                                        mCSIXMLData "OT", "SpotSet", "duration=" & """" & slLen & """"
                                                        mCSIXMLData "CD", "ISCI", slISCI
                                                        mSaveFD slVefCode, slXDReceiverID, slISCI, slCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), "", slXDXMLForm, slEventProgCodeID
                                                        mCSIXMLData "CT", "SpotSet", ""
                                                    End If
                                                ElseIf (ilPassForm = HBFORM) And (Not ilRegionExist) Then
                                                    If llIndexLoop = llIndexStart Then
                                                        mCSIXMLData "OT", "SpotSet", "duration=" & """" & slLength & """"
                                                    End If
                                                    mCSIXMLData "CD", "ISCI", slISCI
                                                    '8279
                                                    mEventIdsAddIsci slISCI, slEventIdsIscis
                                                    mSaveFD slVefCode, slXDReceiverID, slISCI, slCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), "", slXDXMLForm, slEventProgCodeID
                                                    If llIndexLoop = llIndexEnd Then
                                                        mCSIXMLData "CT", "SpotSet", ""
                                                    End If
                                                End If
                                                If igExportSource = 2 Then DoEvents
                                                If ilRegionExist Then
                                                    'RD_ to ISCI becuase ABC and Disney are using the same database and need to not have a confit of names
                                                    'slRISCI = "RD_" & UCase$(mFileNameFilter(slRISCI))
                                                    slRISCI = slISCIPrefix & UCase$(gFileNameFilter(Trim$(slRISCI)))
                                                    If (smAddAdvtToISCI = "Y") And (ilPassForm <> 0) Then
                                                        If tmAstInfo(llIndexLoop).iRegionType = 2 Then
                                                            llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                            '7429
                                                            If llAdf = -1 Then
                                                                gPopAdvertisers
                                                                llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
                                                            End If
                                                            If llAdf <> -1 Then
                                                                slShortTitle = Trim$(Left$(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6)) '& ", " & Trim$(tmAstInfo(llIndexLoop).sRProduct)
                                                                If slShortTitle = "" Then
                                                                    slShortTitle = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
                                                                End If
                                                            '7429
'                                                            Else
'                                                                slShortTitle = Trim$(Left(tmAstInfo(llIndexLoop).sRProduct, 6))
                                                            Else
                                                                slShortTitle = "ADV_MISSING-" & tmAstInfo(llIndexLoop).iAdfCode
                                                                mSetResults "Advertiser missing!  Can't find advertister with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode, MESSAGERED
                                                                myExport.WriteWarning "Advertiser missing. Can't find advertiser with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode & " for regional copy.  Filename or ISCI will not be written correctly.  Search for 'ADV_MISSING' to see the issue.", False
                                                            End If
                                                            slShortTitle = UCase$(gFileNameFilter(slShortTitle))
                                                            'slShortTitle = Left$(slShortTitle, 15)
                                                        End If
                                                        slRISCI = Left$(slShortTitle, 6) & "(" & slRISCI & ")"
                                                    End If
                                                    '11063
                                                    If bmCartReplaceISCI Then
                                                        slRISCI = Trim$(tmAstInfo(llIndexLoop).sRCart)  'not loop  llIndex
                                                        If slRISCI = "" Then
                                                            slRISCI = "MISSING REGIONAL CART"
                                                        End If
                                                    End If
                                                    If ilPassForm = ISCIFORM Then
                                                        mCSIXMLData "OT", "RegionalSpot", ""
                                                        mCSIXMLData "CD", "ISCI", slRISCI
                                                        If tmAstInfo(llIndexLoop).iRegionType = 2 Then
                                                            '7557
                                                            slShortTitle = gXDSShortTitle(CLng(tmAstInfo(llIndexLoop).iAdfCode), tmAstInfo(llIndexLoop).sRProduct, blisCUInHeader, True)
                                                            If InStr(slShortTitle, "ADV_MISSING-") = 1 Then
                                                                mSetResults "Advertiser missing!  Can't find advertister with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode, MESSAGERED
                                                                myExport.WriteWarning "Advertiser missing. Can't find advertiser with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode & " for regional copy.  Filename or ISCI will not be written correctly.  Search for 'ADV_MISSING' to see the issue.", False
                                                            End If
'                                                            llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
'                                                            '7429
'                                                            If llAdf = -1 Then
'                                                                gPopAdvertisers
'                                                                llAdf = gBinarySearchAdf(CLng(tmAstInfo(llIndexLoop).iAdfCode))
'                                                            End If
'                                                            If llAdf <> -1 Then
'                                                                slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtAbbr) & ", " & Trim$(tmAstInfo(llIndexLoop).sRProduct)
'                                                            '7429
'                                                            'Else
'                                                            '    slShortTitle = Trim$(tmAstInfo(llIndexLoop).sRProduct)
'                                                            Else
'                                                               slShortTitle = "ADV_MISSING-" & tmAstInfo(llIndexLoop).iAdfCode
'                                                               mSetResults "Advertiser missing!  Can't find advertister with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode, MESSAGERED
'                                                               myExport.WriteWarning "Advertiser missing. Can't find advertiser with adfCode of " & tmAstInfo(llIndexLoop).iAdfCode & " for regional copy.  Filename or ISCI will not be written correctly.  Search for 'ADV_MISSING' to see the issue.", False
'                                                           End If
'                                                            '12/27/13: Replace ShortTitle by advertiser name so that the Traffic Cart Export (Copy Export) will match
'                                                            '6744 Dan "CUMULUS" is never returned.  Must've wanted when Cumulus on double head end
'                                                           ' If InStr(1, UCase(Trim(slSection)), "CUMULUS", vbBinaryCompare) > 0 Then
'                                                           ' Dan M 10/27/14 look for "-CU" at top
'                                                           ' If InStr(1, UCase(Trim(slSection)), "-CU", vbBinaryCompare) > 0 Then
'                                                            If blisCUInHeader Then
'                                                                If llAdf <> -1 Then
'                                                                    slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtName)
'                                                                End If
'                                                            End If
'                                                            slShortTitle = UCase$(gFileNameFilter(slShortTitle))
                                                        End If
                                                        mSaveFD slVefCode, slStationID, slRISCI, slRCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), slShortTitle, slXDXMLForm, ""
                                                        '7496
                                                        'mCSIXMLData "CD", "FileName", slShortTitle & "(" & slRISCI & ")" & ".MP2"
                                                        mCSIXMLData "CD", "FileName", slShortTitle & "(" & slRISCI & ")" & UCase(sgAudioExtension)
                                                        mCSIXMLData "CD", "Duration", slLen
                                                        mCSIXMLData "CT", "RegionalSpot", ""
                                                    ElseIf ilPassForm = HBPFORM Then  'HBP
                                                        mCSIXMLData "OT", "SpotSet", "duration=" & """" & slLen & """"
                                                        mCSIXMLData "CD", "ISCI", slRISCI
                                                        mSaveFD slVefCode, slXDReceiverID, slRISCI, slRCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), "", slXDXMLForm, slEventProgCodeID
                                                        mCSIXMLData "CT", "SpotSet", ""
                                                    ElseIf ilPassForm = HBFORM Then  'HB
                                                        'Loop on all spots within same break
                                                        If llIndexLoop = llIndexStart Then
                                                            mCSIXMLData "OT", "SpotSet", "duration=" & """" & slLength & """"
                                                        End If
                                                        mCSIXMLData "CD", "ISCI", slRISCI
                                                        '8279
                                                        mEventIdsAddIsci slRISCI, slEventIdsIscis
                                                        mSaveFD slVefCode, slXDReceiverID, slRISCI, slRCreative, gDateValue(gAdjYear(Format(slRotStartDT, "mm/dd/yy"))), gDateValue(gAdjYear(Format(slRotEndDT, "mm/dd/yy"))), "", slXDXMLForm, slEventProgCodeID
                                                        If llIndexLoop = llIndexEnd Then
                                                            mCSIXMLData "CT", "SpotSet", ""
                                                        End If
                                                    End If 'pass form choice
                                                End If 'region?
                                                
                                                If igExportSource = 2 Then DoEvents
                                                ' 6082 get the regionals
                                                '7/9/13: Transparent file generatrion is independent of the Unit ID setting
                                                'If smUnitIdByAstCode = "Y" Then
                                                If smGenTransparency = "Y" Then
                                                    If ilRegionExist Then
                                                        If Not mAstFileSave(slXDReceiverID, slStationName, slRISCI, slRCreative, llIndexLoop, slTransmissionID, slEventProgCodeID, True, slVehicleName) Then
                                                            bmAstFileError = True
                                                        End If
                                                    Else
                                                        If Not mAstFileSave(slXDReceiverID, slStationName, slISCI, slCreative, llIndexLoop, slTransmissionID, slEventProgCodeID, False, slVehicleName) Then
                                                            bmAstFileError = True
                                                        End If
                                                    End If
                                                End If
                                                '11/4/11: HB was counting breaks, so count was move here to count spots
                                                If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                                                    llTotalExport = llTotalExport + 1
                                                    'Dan M 11/4/14 moved to below
                                                    '6882
'                                                    ilNeedExport = ilNeedExport + 1
                                                    '7180 added bmReExportForce
                                                    If bmAllowXMLCommands Or bmReExportForce Then
                                                        ilNeedExport = ilNeedExport + 1
                                                        '7256
                                                        If blisReExport Then
                                                            llReExportSent = llReExportSent + 1
                                                        Else
                                                            llNewExportSent = llNewExportSent + 1
                                                        End If
'                                                        '7458
                                                        If Not myEnt.Add(tmAstInfo(llIndexLoop).sFeedDate, tmAstInfo(llIndexLoop).lgsfCode, , , True) Then
                                                            myExport.WriteWarning myEnt.ErrorMessage
                                                        End If
'                                                    Else
'                                                        '7458
'                                                        If Not myEnt.Add(tmAstInfo(llIndexLoop).sFeedDate, tmAstInfo(llIndexLoop).lGsfCode, Asts) Then
'                                                            myExport.WriteWarning myEnt.ErrorMessage
'                                                        End If
                                                    End If
                                                End If
                                            End If ' 6082 ilInclude
                                        Next llIndexLoop
                                        '6082
                                        If ilIncludeSpot Then
                                            llIndex = llIndexEnd
                                        'If ilPassForm = 0 Then
                                        '    mCSIXMLData "CT", "Insert", ""
                                        '    mCSIXMLData "CT", "Site", ""
                                        'ElseIf (ilPassForm = 1) Or (ilPassForm = 2) Then
                                        '    mCSIXMLData "OT", "Sites", ""
                                        '    If ilPassForm = 1 Then
                                        '        '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
                                        '        'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHBP & """"
                                        '        mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5 & """"
                                        '    Else
                                        '        '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
                                        '        'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHB & """"
                                        '        mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHB & slVefCode5 & """"
                                        '    End If
                                        '    mCSIXMLData "CT", "Site", ""
                                        '    mCSIXMLData "CT", "Sites", ""
                                        '    mCSIXMLData "CT", "Insert", ""
                                        'End If
                                            If slSendProgAndSiteBy = "B" Then
                                                mXMLSiteTags ilPassForm, slXDReceiverID, slTransmissionID, slUnitHB, slUnitHBP, slVefCode5, llAstCode
                                            End If
                                            '8279
                                            If blLeaveCueAlone Then
                                                'skip first one!
                                                For ilEventIdIndex = 1 To UBound(tlEventIds) - 1 Step 1
                                                    'false if empty code and cue but it's never empty, so it's useless!
                                                    If mEventIdsWriteExtraInserts(tlEventIds(ilEventIdIndex).sCode, tlEventIds(ilEventIdIndex).sCue, slVehicleName, slRotStartDT, slRotEndDT, slTransmissionID, slLength, slEventIdsIscis) Then
                                                        If Not mEventIDXMLSiteTags(ilEventIdIndex, slXDReceiverID, slTransmissionID, slUnitHB, slVefCode5) Then
                                                            blRetStatus = False
                                                            mSetResults "Event ID issue.  Could not alter unit ids as needed", MESSAGERED
                                                            myExport.WriteWarning "Event ID issue.  Could not alter unit ids as needed", False
                                                        End If
                                                        'ubound is count because the upper is not used!
                                                        mAddToCount blisReExport, UBound(slEventIdsIscis), llTotalExport, llReExportSent, llNewExportSent, ilNeedExport
                                                    End If
                                                Next ilEventIdIndex
                                            End If
                                            '5/7/14 we don't use this anymore, so comment out!
'                                            If slSendWriteBy = "B" Then
'                                                'csiXMLData "CT", "Sites", ""
'                                                If udcCriteria.XExportType(0, "V") = vbChecked Then
'                                                    '6979
'                                                    mAddSurroundingElement ilPassForm, False
'                                                    '6635 same as below, but now uses Jeff's error log
'                                                    '7508
'                                                    'If Not mSendAndWriteReturn("Spot Insertions") Then
'                                                    If Not mSendAndTestReturn(XDSType.Insertions, slDoNotReturn) Then
'                                                        blRetStatus = False
'    '                                                    If llTotalExport > 0 Then
'    '                                                        llTotalExport = llTotalExport - 1
'    '                                                    End If
'                                                        If bmIsError Then
'                                                            Exit Function
'                                                        End If
'                                                        'gLogMsg "Station being sent: " & slStationName & " , Vehicle being sent: " & slVehicleName, smPathForgLogMsg, False
'                                                        myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName
'                                                        Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", MESSAGERED)
'    '                                                ilRet = csiXMLWrite(1)
'    '                                                If ilRet <> True Then
'    '                                                    '11/26/12: Continue to next CPTT
'    '                                                    'imTerminate = True
'    '                                                    'imExporting = False
'    '                                                    blRetStatus = False
'    '                                                    'dan 3/13/13 remove from count
'    '                                                    If llTotalExport > 0 Then
'    '                                                        llTotalExport = llTotalExport - 1
'    '                                                    End If
'    '                                                   ilRet = csiXMLStatus(tlXmlStatus)
'    '                                                   '5896  log error, change name of log to show user in display
'    '                                                    If mIsXmlError(tlXmlStatus.sStatus) Then
'    '                                                        Exit Function
'    '                                                    End If
'                                                    'Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", RGB(155, 0, 0))
'        '                                                    '5896
'        '                                                If gIsNull(tlXmlStatus.sStatus) Then
'        '                                                    gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR", "XDigitalExportLog.Txt", False
'        '                                                Else
'        '                                                    gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
'        '                                                End If
'                                                        'gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " Error: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
'                                                        'Exit Function
'                                                        '6689 remove!
'                                                        'Exit Do
'                                                        '7236 not yet implemented here
''                                                        '7236 if not a successful send, clear rs without writing to xht.
''                                                        blPartialExportMayContinue = False
'                                                    '7236
'                                                    Else
''                                                        ilRet = mUpdateXHT()
'                                                        ' Dan M 7/13/12 add UpdateLastExportDate, only when transmitting
'                                                        If udcCriteria.XGenType(0) Then
'                                                            gUpdateLastExportDate imVefCode, slEDate  'or sledate?
'                                                        End If
'                                                    End If
'                                                    '11/4/11: HB was counting breaks, count move up in llIndexLoop
'                                                    'llTotalExport = llTotalExport + 1
'                                                End If ' send to xds
'                                            '6882 send if by vehicle and at max
'                                            ElseIf slSendWriteBy = "V" Then
                                            If slSendWriteBy = "V" Then
                                                ' at least one to send and chose spot insertions
                                                If ilNeedExport > ilSafeChunkSize And udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                                                    blSentThisVehicle = True
                                                    ilNeedExport = 0
                                                    '6979
                                                    mAddSurroundingElement ilPassForm, False
                                                    '7508
                                                    'If Not mSendAndWriteReturn("Spot Insertions") Then
                                                    If Not mSendAndTestReturn(Insertions, slDoNotReturn) Then
                                                        blRetStatus = False
                                                        If bmIsError Then
                                                            '7508
                                                            If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                                                                myExport.WriteWarning myEnt.ErrorMessage
                                                            End If
                                                            Exit Function
                                                        End If
                                                        myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName
                                                        Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", MESSAGERED)
                                                        '7508 don't need slPartial anymore
                                                         '7236 if not a successful send, stop later update
                                                       ' slPartialExportMayContinue = "N"
                                                        slDoNotMaster = slDoNotMaster & slDoNotReturn
                                                   Else
                                                        '7236 will update lower down
                                                        'ilRet = mUpdateXHT()
                                                        ' Dan M 7/13/12 add UpdateLastExportDate, only when transmitting
                                                        If udcCriteria.XGenType(0, slProp) Then
                                                            gUpdateLastExportDate imVefCode, slEDate
                                                        End If
'                                                        If slPartialExportMayContinue <> "N" Then
'                                                            slPartialExportMayContinue = "Y"
'                                                        End If
                                                    End If
                                                End If
                                            End If 'write by B
                                        End If 'include - for 6082
                                        If igExportSource = 2 Then DoEvents
                                    ' 6082 get non regionals
                                        '7/9/13: Transparent file generatrion is independent of the Unit ID setting
                                        'If smUnitIdByAstCode = "Y" Then
                                        If smGenTransparency = "Y" Then
                                            If blFailedBecauseNotRegional Then
                                                If Not mAstFileSave(slXDReceiverID, slStationName, slISCI, slCreative, llIndex, slTransmissionID, slEventProgCodeID, False, slVehicleName) Then
                                                    bmAstFileError = True
                                                End If
                                            End If
        '                                    If Len(slRISCI) > 0 Then
        '                                        If Not mAstFileSave(slVehicleName, slStationName, slRISCI, slRCreative, llIndex, slTransmissionID, slEventProgCodeID) Then
        '                                            bmAstFileError = True
        '                                        End If
        '                                    Else
        '                                        If Not mAstFileSave(slVehicleName, slStationName, slISCI, slCreative, llIndex, slTransmissionID, slEventProgCodeID) Then
        '                                            bmAstFileError = True
        '                                        End If
        '                                    End If
                                        End If
                                    End If 'include spot?
                                End If ' spot ok
                                llIndex = llIndex + 1
                                If imTerminate Then
                                    imExporting = False
                                    '8236
                                    If blCEFIsOpen Then
                                        mCloseCEFFile
                                    End If
                                    Exit Function
                                End If
                            Loop 'BASIC SPOT LOOP
                            
                        '6/19/14: ttp 6944 Separate out events by Program Code ID if a sport event
                        Next ilProgCode
                        
                        '7/24/14: Remove records not found in current export
                        'ilRet = mSendDeleteCommands()
                        mRetainDeletions
                        '7/24/14: Update records
                        '7236 we update by agreement, but we don't 'confirm' since we actually send by vehicle. Confirm (status = Y) after vehicle
                        'ilRet = mUpdateXHT()
                        slTempAtt = mUpdateXHT()
                        If Len(slTempAtt) > 0 Then
                            slConfirmAtts = slConfirmAtts & slTempAtt
                        End If
                        '7458
                        If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
                            If Not myEnt.CreateEnts(Incomplete) Then
                                myExport.WriteWarning myEnt.ErrorMessage
                            End If
                        Else
                            myEnt.ClearWhenDontSend
                        End If
                        'csiXMLData "CT", "Site", ""
                        'csiXMLData "CT", "Sites", ""
                        'ilRet = csiXMLWrite(1)
                        '5/7/15 Dan always B
'                        If slSendProgAndSiteBy = "D" Then
'                            mXMLSiteTags ilPassForm, slXDReceiverID, slTransmissionID, slUnitHB, slUnitHBP, slVefCode5, 0
'                        End If
                        '5/7/15 Dan don't use, so comment out
'                        If slSendWriteBy = "D" Then
'                            If udcCriteria.XExportType(0, "V") = vbChecked Then
'                                '6979
'                                mAddSurroundingElement ilPassForm, False
'                                '6635 same as below, but now uses Jeff's error log
'                                '7508
'                                'If Not mSendAndWriteReturn("Spot Insertions") Then
'                                If Not mSendAndTestReturn(Insertions, slDoNotReturn) Then
'                                    blRetStatus = False
''                                    If llTotalExport > 0 Then
''                                        llTotalExport = llTotalExport - 1
''                                    End If
'                                    If bmIsError Then
'                                        'Dan 6/27/14 write out the vehicle where halting occured.
'                                        myExport.WriteFacts "Program halted on Vehicle: " & slVehicleName, True
'                                        Exit Function
'                                    End If
'                                    'gLogMsg "Station being sent: " & slStationName & " , Vehicle being sent: " & slVehicleName, smPathForgLogMsg, False
'                                    'Dan 6/27/14 cleaned this up.  This is not an 'error', but a warning.
'                                   ' myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName
'                                    myExport.WriteFacts "Warning above for Vehicle: " & slVehicleName, True
'                                    Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", MESSAGERED)
''                                ilRet = csiXMLWrite(1)
''                                If ilRet <> True Then
''                                    '5896
''                                     blRetStatus = False
''                                     'dan 3/13/13 remove from count
''                                     If llTotalExport > 0 Then
''                                         llTotalExport = llTotalExport - 1
''                                     End If
''                                    ilRet = csiXMLStatus(tlXmlStatus)
''                                    '5896  log error, change name of log to show user in display
''                                     If mIsXmlError(tlXmlStatus.sStatus) Then
''                                         Exit Function
''                                     End If
''                                    '11/26/12: Continue to next CPTT
''                                    'imTerminate = True
''                                    'imExporting = False
''                                    blRetStatus = False
''                                    ilRet = csiXMLStatus(tlXmlStatus)
''                                    Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt", RGB(155, 0, 0))
''                                    '5896
''                                    If gIsNull(tlXmlStatus.sStatus) Then
''                                        gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR", "XDigitalExportLog.Txt", False
''                                    Else
''                                        gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                                    End If
''                                    'gLogMsg "Vehicle " & slVehicleName & " Station " & slStationName & " Error: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                                    'Exit Function
'                                    '6689 remove
'                                    'Exit Do
'                                ' Dan M 7/13/12 add UpdateLastExportDate, only when transmitting
'                                '7236 not yet implemented
'                                Else
'                                    If udcCriteria.XGenType(0) Then
'                                        gUpdateLastExportDate imVefCode, slEDate  'or sledate?
'                                    End If
'                                '11/4/11: HB was counting breaks, count move up in llIndexLoop
'                                'llTotalExport = llTotalExport + 1
'                                End If
'                            End If
'                        End If 'send by day
                        llODate = -1
                    End If 'vpfCode valid
                    If igExportSource = 2 Then DoEvents
                End If 'station vehicle ok
            End If 'station vehicle not voice track
            cprst.MoveNext
        Wend 'getting next agreement
        '6882 final for send by vehicle
        If slSendWriteBy = "V" Then
            '7/25/14: If only deletes defined, export to x-digital
            If (ilNeedExport = 0) And (udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked) And (UBound(tmRetainDeletions) > LBound(tmRetainDeletions)) Then
                If bmWroteTopElement = False Then
                    mAddSurroundingElement ilPassForm, True
                    bmWroteTopElement = True
                    ilNeedExport = 1
                End If
            End If
            If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                ' at least one to send and chose spot insertions
                If ilNeedExport > 0 Then
                    blSentThisVehicle = True
                    '6979
                    mAddSurroundingElement ilPassForm, False
                    '7/24/14: Send delete commands
                    ilRet = mSendDeleteCommands()
                    ilNeedExport = 0
                    'The last vehicle gets sent from here.  If only one, it's here too
                    '7508
                    If Not mSendAndTestReturn(XDSType.Insertions, slDoNotReturn) Then
                        blRetStatus = False
                        slDoNotMaster = slDoNotMaster & slDoNotReturn
                        If bmIsError Then
                            '7508
                            If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                                myExport.WriteWarning myEnt.ErrorMessage
                            End If
                            Exit Function
                        End If
                        myExport.WriteError "Error above for Station: " & slStationName & " , Vehicle: " & slVehicleName, False
                        Call mSetResults("Export not completely successful. see " & smPathForgLogMsg, MESSAGERED)
                    Else
                        If udcCriteria.XGenType(0, slProp) Then
                            gUpdateLastExportDate imVefCode, slEDate
                        End If
                    End If
                End If
                'block nothing to go out. Could be set at the 'partial'
                If blSentThisVehicle Then
                    blSentThisVehicle = False
                    If bmFailedToReadReturn Then
                        slConfirmAtts = ""
                        slDoNotMaster = ""
                    Else
                        'changes slDoNotMaster to bad atts
                        slConfirmAtts = mAdjustAtts(slConfirmAtts, slDoNotMaster)
                        If Len(slDoNotMaster) = 0 Then
                            slDoNotMaster = "0,"
                        End If
                    End If
                    'confirm the xht.  The 'unconfirmed' will be wiped out next time through
                    mConfirmXHT slConfirmAtts, slSDate, sEndDate
                    If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
                        ' 'bad' get error; all others get success. this handles normal and forced reexport where atts aren't saved
                        'if bmFailedToReadReturn, slDoNotMaster is blank, so all will get error
                        If Not myEnt.UpdateIncompleteByFilename(EntError, , , slDoNotMaster) Then
                            myExport.WriteWarning myEnt.ErrorMessage
                        End If
                    End If
                End If
                'reset per vehicle because 'partial' may have had issue
                bmFailedToReadReturn = False
            End If
        End If
        If imTerminate Then
            imExporting = False
            '8236
            If blCEFIsOpen Then
                mCloseCEFFile
            End If
            Exit Function
        End If
        If igExportSource = 2 Then DoEvents
        llODate = -1
        
        '12/11/17: Clear abf
        If igTimes = 1 Then
            If lbcStation.ListCount <= 0 Or chkAllStation.Value = vbChecked Then
                'gClearAbf imVefCode, 0, sMoDate, gObtainNextSunday(sMoDate)
            End If
        End If
        
        If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
            gClearASTInfo True
        Else
            gClearASTInfo False
        End If
        sMoDate = DateAdd("d", 7, sMoDate)
        slSDate = sMoDate
        slEDate = gObtainNextSunday(slSDate)
        If gDateValue(gAdjYear(sEndDate)) < gDateValue(gAdjYear(slEDate)) Then
            slEDate = sEndDate
        End If
    Loop While gDateValue(gAdjYear(sMoDate)) < gDateValue(gAdjYear(sEndDate))
    '7/6/14: Allow XML
    bmAllowXMLCommands = True
    '8280 moved to below
'    If udcCriteria.XExportType(0, "V") = vbChecked Then
'        ilRet = csiXMLWrite(1)  ' Call one last time to flush the send queue.
'    End If
    '5/21/15 Dan always 2 passes, so this gets called twice. Not necessary. Move to mExport
    '12/20/14: Remove XHT records regardless of agreements that are two or more weeks old
   ' mRemoveOldXHT
    'Dan 3/13/13 only show if chose spots exported
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
        '8280 only if sending
        If udcCriteria.XGenType(0, slProp) Then
            ilRet = csiXMLWrite(1)
        End If
'        'dan 4/23/13
'        Select Case ilPassForm
'            Case 0
'                mSetResults "Total ISCI Spots Exported = " & llTotalExport, 0
'            Case 1
'                mSetResults "Total HBP Spots Exported = " & llTotalExport, 0
'            Case 2
'                mSetResults "Total HB Spots Exported = " & llTotalExport, 0
'        End Select
        '6979  reexports
        Select Case ilPassForm
            Case 0
                slLogMessage = "Total ISCI Spots Found = " & llTotalExport
               ' mSetResults "Total ISCI Spots Found = " & llTotalExport, 0
            Case 1
                slLogMessage = "Total HBP Spots Found = " & llTotalExport
               ' mSetResults "Total HBP Spots Found = " & llTotalExport, 0
            Case 2
                slLogMessage = "Total HB Spots Found = " & llTotalExport
                'mSetResults "Total HB Spots Found = " & llTotalExport, 0
        End Select
        mSetResults slLogMessage, MESSAGEBLACK, True
        '7256
        myExport.WriteFacts slLogMessage, True
        ' 6979 > because getting regional and national
          '  mSetResults "Re-exports. " & lmReExportNew - llReExportSent & " were sent as new.", 0
        mSetResults "New exports. " & llNewExportSent & " were sent.", 0, True
        mSetResults "Re-exports. " & llReExportSent & " were sent.", 0, True
        mSetResults "Re-exports. " & lmReExportDelete & " were deleted.", 0, True
        myExport.WriteFacts "New exports-" & llNewExportSent & vbCrLf & "Re-exports-" & llReExportSent & vbCrLf & "Deletes-" & lmReExportDelete
'        Call mSetResults("Total Spots Exported = " & llTotalExport, 0)
    End If
    
    '4/18/13: Check if any spots not merged
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
        ilVef = -1
        For llAst = 0 To UBound(tmMergeAstInfo) - 1 Step 1
            If tmMergeAstInfo(llAst).lCode > 0 Then
                If ilVef <> tmMergeAstInfo(llAst).iVefCode Then
                    slVehicleName = mGetVehicleName(tmMergeAstInfo(llAst).iVefCode)
                    Call mSetResults("Spots from " & slVehicleName & " not Merged into any Export", MESSAGEBLACK)
                    'gLogMsg "Spots from " & slVehicleName & " not Merged into any Export", smPathForgLogMsg, False
                    myExport.WriteWarning "Spots from " & slVehicleName & " not Merged into any Export"
                    ilVef = tmMergeAstInfo(llAst).iVefCode
                End If
            End If
        Next llAst
    End If
    '8236
    If blCEFIsOpen Then
        mCloseCEFFile
    End If
    mExportSpotInsertions = blRetStatus
    Exit Function
mExportSpotInsertionsErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mExportSpotInsertions"
    mExportSpotInsertions = False
    Exit Function
    
End Function
Private Function mGetCreative(llCpfCode As Long) As String
    Dim slRet As String
    Dim slSql As String
    Dim rstCpf As Recordset
    
    On Error GoTo ERRORBOX
    slRet = ""
    If llCpfCode > 0 Then
        slSql = "select cpfCreative from CPF_Copy_Prodct_ISCI where cpfCode = " & llCpfCode
        Set rstCpf = gSQLSelectCall(slSql)
        If Not rstCpf.EOF Then
            If Not IsNull(rstCpf!cpfCreative) Then
                slRet = gXMLNameFilter(Trim$(rstCpf!cpfCreative))
            End If
        End If
    End If
Cleanup:
    If Not rstCpf Is Nothing Then
        If (rstCpf.State And adStateOpen) <> 0 Then
            rstCpf.Close
        End If
        Set rstCpf = Nothing
    End If
    mGetCreative = slRet
    Exit Function
ERRORBOX:
    slRet = ""
    gHandleError smPathForgLogMsg, "frmExportXDigital-mGetCreative"
    GoTo Cleanup
End Function

Private Function mIsXmlError(slMessage As String, Optional slRoutine As String = "SpotInsertions", Optional blHalt = True) As Boolean
    'is error as opposed to warning?  Shut down.  Assume not an error <msgs><msg code="-2" ID="20136" name="SiteID">Invalid SiteID=20136, item ignored</msg><msg code="0" ID="" name="">OK</msg></msgs>
    '6966 added optional 'halt' to help with resends
    Dim blRet As Boolean
   ' Dim ilPos As Integer
    Dim ilPos As Long
    Dim ilEnd As Long
    Dim slRet As String
  On Error GoTo ERRORBOX
    slRet = ""
    blRet = False
    ilPos = InStr(slMessage, "<FAULTSTRING>")
    If ilPos > 0 Then
        ilPos = ilPos + 13
        ilEnd = InStr(slMessage, "</FAULTSTRING>")
        If ilEnd > ilPos Then
            slRet = Mid(slMessage, ilPos, ilEnd - ilPos)
        End If
        If InStr(1, slRet, "HTTP HEADER SOAPACTION:") > 0 Then
            bmIsWrongServicePage = True
        Else
            blRet = True
        End If
    '6973 look for bad request
    ElseIf InStr(slMessage, "BAD REQUEST") > 0 Then
        blRet = True
        slRet = slMessage
    '<MSGS><MSG CODE=5 ID=0
    Else
'        ilPos = InStr(slMessage, "<msg code=5")
'        If ilPos > 0 Then
'            blRet = True
'            ilPos = InStr(ilPos, slMessage, ">")
'            If ilPos > 0 Then
'                ilEnd = InStr(ilPos, slMessage, "</msg>")
'                If ilEnd > ilPos Then
'                    slRet = Mid(slMessage, ilPos + 1, ilEnd - ilPos - 1)
'                End If
'            End If
'            If Len(slRet) = 0 Then
'                slRet = " message code = 5 "
'            End If
'        ' -5  could not connect to database
'        Else
'            ilPos = InStr(slMessage, "<msg code=-5")
'            If ilPos > 0 Then
'                blRet = True
'                ilPos = InStr(ilPos, slMessage, ">")
'                If ilPos > 0 Then
'                    ilEnd = InStr(ilPos, slMessage, "</msg>")
'                    If ilEnd > ilPos Then
'                        slRet = Mid(slMessage, ilPos + 1, ilEnd - ilPos - 1)
'                    End If
'                End If
'                If Len(slRet) = 0 Then
'                    slRet = " message code = -5 "
'                End If
'            End If
'        End If
        '6922 I used to read the first message and see if there was a -4,-5,or -6.  Now I look for "OK"
'        '6754
'        ilPos = InStr(slMessage, "<msg code=")
'        If ilPos > 0 Then
'            Dim slNum As String
'            slNum = Mid(slMessage, ilPos + 10, 2)
'            If slNum = "-5" Or slNum = "-6" Or slNum = "-4" Then
'                blRet = True
'                ilPos = InStr(ilPos, slMessage, ">")
'                If ilPos > 0 Then
'                    ilEnd = InStr(ilPos, slMessage, "</msg>")
'                    If ilEnd > ilPos Then
'                        slRet = Mid(slMessage, ilPos + 1, ilEnd - ilPos - 1)
'                    End If
'                End If
'                slRet = " message code = " & slNum & " " & slRet
'            End If
'        End If
        'some messages are not complete. If not, assume NOT an error, but just a warning
        If InStr(slMessage, "</MSGS>") > 0 Then
            ilPos = InStr(slMessage, ">OK</MSG")
            If ilPos = 0 Then
                blRet = True
                slRet = "Ok not returned: " & slMessage
            End If
        Else
            myExport.WriteWarning "Return message not complete", True
        End If
    End If
    If blRet Then
        bmIsError = True
        '6966
        If blHalt Then
             Call mSetResults("Xml error in " & slRoutine & ". see XDigitalExportLog.Txt", MESSAGERED)
             mSetResults "Export halted.", MESSAGERED
            ' gLogMsg " ERROR: " & slRet, smPathForgLogMsg, False  '"Vehicle " & slVehicleName & " Station " & slStationName &
             myExport.WriteError "Export halted: " & slRet
             lacProcessing.Visible = False
             lacResult.Caption = "Check '" & smPathForgLogMsg & "' in Messages folder for issue."
             gCustomEndStatus lmEqtCode, igExportReturn, ""
             imExporting = False
             cmdExport.Enabled = False
             cmdExportTest.Enabled = False
             cmdCancel.Caption = "&Done"
             Screen.MousePointer = vbDefault
             cmdCancel.SetFocus
        Else
             myExport.WriteWarning "Xml error in " & slRoutine & ":" & slRet
        End If
    End If
    mIsXmlError = blRet
    Exit Function
ERRORBOX:
    mIsXmlError = False
End Function


Private Sub mFillStations()
    
    Dim ilRet As Integer
    Dim llVef As Long
    Dim ilVff As Integer
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim slDate As String
        
    On Error GoTo ErrHand
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    'Only get agreements that are to be sent to web and are active
    If edcDate.Text = "" Then
        slDate = Format(gNow(), sgSQLDateForm)
    ElseIf gIsDate(edcDate.Text) = False Then
        slDate = Format(gNow(), sgSQLDateForm)
    Else
        slDate = Format(edcDate.Text, sgSQLDateForm)
    End If
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery & " AND (vatWvtVendorId = " & Vendors.XDS_Break & " OR vatWvtVendorId = " & Vendors.XDS_ISCI & ") "
    SQLQuery = SQLQuery & " AND attonAir <= '" & slDate & "' AND attdropdate > '" & slDate & "' AND attoffair > '" & slDate & "'"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcStation.AddItem Trim$(rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
        rst.MoveNext
    Wend
    'If Log vehicle, check all vehicles that are part of that log vehicle
    llVef = gBinarySearchVef(CLng(imVefCode))
    If llVef <> -1 Then
        If tgVehicleInfo(llVef).sVehType = "L" Then
            For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If imVefCode = tgVehicleInfo(ilLoop).iVefCode Then
                    For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                        If tgVehicleInfo(ilLoop).iCode = tgVffInfo(ilVff).iVefCode Then
                            If tgVffInfo(ilVff).sMergeWeb <> "S" Then
                                SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
                                SQLQuery = SQLQuery & " FROM shtt, att"
                                SQLQuery = SQLQuery & " WHERE (attVefCode = " & tgVehicleInfo(ilLoop).iCode
                                SQLQuery = SQLQuery & " AND attExportType = 1"
                                SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
                                SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
                                Set rst = gSQLSelectCall(SQLQuery)
                                While Not rst.EOF
                                    'Check that name has not been previously added
                                    llRow = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, Trim$(rst!shttCallLetters))
                                    If llRow < 0 Then
                                        lbcStation.AddItem Trim$(rst!shttCallLetters)
                                        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                    End If
                                    rst.MoveNext
                                Wend
                            End If
                            Exit For
                        End If
                    Next ilVff
                End If
            Next ilLoop
        End If
    End If
    chkAllStation.Value = vbChecked
    gSetMousePointer grdVeh, grdVeh, vbDefault
    chkAllStation_Click
    
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mFillStations"
    Exit Sub
End Sub

'Private Sub mFillStations()
'
'   Dim slNowDate As String
'
'    On Error GoTo ErrHand
'    Screen.MousePointer = vbHourglass
'    '8162
'    slNowDate = Format(gAdjYear(edcDate.Text), sgSQLDateForm)
'    If Not IsDate(slNowDate) Then
'        slNowDate = Format(gNow(), sgSQLDateForm)
'    End If
'    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
'    SQLQuery = SQLQuery & " FROM shtt, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
'    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
'    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
'    SQLQuery = SQLQuery & " AND (vatWvtVendorId = " & Vendors.XDS_Break & " OR vatWvtVendorId = " & Vendors.XDS_ISCI & ") "
'    SQLQuery = SQLQuery & " AND attonAir <= '" & slNowDate & "' AND attdropdate > '" & slNowDate & "' AND attoffair > '" & slNowDate & "'"
'    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
'    Set rst = gSQLSelectCall(SQLQuery)
'    While Not rst.EOF
'        lbcStation.AddItem Trim$(rst!shttCallLetters)
'        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
'        rst.MoveNext
'    Wend
'    chkAllStation.Value = vbChecked
'    Screen.MousePointer = vbDefault
'    Exit Sub
'ErrHand:
'    gSetMousePointer grdVeh, grdVeh, vbDefault
'    gHandleError smPathForgLogMsg, "frmExportXDigital-mFillStations"
'End Sub

'Private Sub mFillStations()
''8163 all new
'   Dim slNowDate As String
'
'    On Error GoTo ErrHand
'    Screen.MousePointer = vbHourglass
'    '8162
'    slNowDate = Format(gAdjYear(edcDate.Text), sgSQLDateForm)
'    If Not IsDate(slNowDate) Then
'        slNowDate = Format(gNow(), sgSQLDateForm)
'    End If
'    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
'    SQLQuery = SQLQuery & " FROM shtt, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
'    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
'    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
'    SQLQuery = SQLQuery & " AND (vatWvtVendorId = " & Vendors.XDS_Break & " OR vatWvtVendorId = " & Vendors.XDS_ISCI & ") "
'    SQLQuery = SQLQuery & " AND attonAir <= '" & slNowDate & "' AND attdropdate > '" & slNowDate & "' AND attoffair > '" & slNowDate & "'"
'    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
'    Set rst = gSQLSelectCall(SQLQuery)
'    While Not rst.EOF
'        lbcStation.AddItem Trim$(rst!shttCallLetters)
'        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
'        rst.MoveNext
'    Wend
'    chkAllStation.Value = vbChecked
'    Screen.MousePointer = vbDefault
'    Exit Sub
''    On Error GoTo ErrHand
''    Screen.MousePointer = vbHourglass
''    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
''    SQLQuery = SQLQuery & " FROM shtt, att"
''    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
''    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
''    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
''    Set rst = gSQLSelectCall(SQLQuery)
''    While Not rst.EOF
''        If igExportSource = 2 Then DoEvents
''        lbcStation.AddItem Trim$(rst!shttCallLetters)
''        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
''        rst.MoveNext
''    Wend
''    chkAllStation.Value = vbChecked
''    Screen.MousePointer = vbDefault
''    Exit Sub
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gHandleError smPathForgLogMsg, "frmExportXDigital-mFillStations"
'End Sub

Private Function mGetSdfCopy(llLstCode As Long, slCart As String, slISCI As String, slCreative As String, slLen As String) As Integer
    
    Dim lst_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    slCart = ""
    slISCI = ""
    slCreative = ""
    SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstLen, cpfCreative"
    'SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
    SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode)"
    SQLQuery = SQLQuery & " WHERE lstCode =" & Str(llLstCode)
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If Not lst_rst.EOF Then
        'If IsNull(rst!adfName) = True Then
        '    slAdvt = "Missing"
        'Else
        '    slAdvt = Trim$(rst!adfName)
        'End If
        'If IsNull(rst!lstProd) = True Then
        '    slProd = ""
        'Else
        '    slProd = Trim$(rst!lstProd)
        'End If
        If igExportSource = 2 Then DoEvents
        If IsNull(lst_rst!lstCart) Or Left$(lst_rst!lstCart, 1) = Chr$(0) Then
            slCart = ""
        Else
            slCart = Trim$(lst_rst!lstCart)
        End If
        If IsNull(lst_rst!lstISCI) = True Then
            slISCI = ""
        Else
            slISCI = Trim$(lst_rst!lstISCI)
        End If
        If IsNull(lst_rst!cpfCreative) = True Then
            slCreative = ""
        Else
            slCreative = gXMLNameFilter(Trim$(lst_rst!cpfCreative))
        End If
        slLen = Trim$(Str$(lst_rst!lstLen))
    End If
    lst_rst.Close
    mGetSdfCopy = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mGetSdfCopy"
    mGetSdfCopy = False
    Exit Function
End Function

Public Sub mSetResults(slMsg As String, llFGC As Long, Optional blAllowDuplicate = False)
    '7256 allow duplicates Dan 11/14/14
    Dim llLoop As Long
    bmMgsPrevExisted = False
    For llLoop = 0 To lbcMsg.ListCount - 1 Step 1
        If slMsg = lbcMsg.List(llLoop) Then
            bmMgsPrevExisted = True
            If blAllowDuplicate = False Then
                Exit Sub
            End If
        End If
    Next llLoop
    'dan 11/06/12
    'lbcMsg.AddItem slMsg
    gAddMsgToListBox FrmExportXDigital, lmMaxWidth, slMsg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
   ' lbcMsg.ForeColor = llFGC
    If lbcMsg.ForeColor <> MESSAGERED Then
        lbcMsg.ForeColor = llFGC
    End If

    If igExportSource = 2 Then DoEvents
End Sub

Private Function mGetVehicleName(iVefCode As Integer) As String
    Dim llLoop As Integer
    mGetVehicleName = ""
    For llLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgVehicleInfo(llLoop).iCode = iVefCode Then
            mGetVehicleName = Trim(tgVehicleInfo(llLoop).sVehicle)
            Exit For
        End If
    Next
End Function

Private Function mGetStationName(iShttCode As Integer, iAckDaylight As Integer) As String
    '10933 added acknowledge daylight
    Dim llLoop As Integer
    mGetStationName = ""
    For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgStationInfo(llLoop).iCode = iShttCode Then
            mGetStationName = Trim(tgStationInfo(llLoop).sCallLetters)
            iAckDaylight = tgStationInfo(llLoop).iAckDaylight
            Exit For
        End If
    Next
End Function

'Private Function mFileNameFilter(slInName As String) As String
'    Dim slName As String
'    Dim ilPos As Integer
'    Dim ilFound As Integer
'    slName = slInName
'    'Remove " and '
'    Do
'        If igExportSource = 2 Then DoEvents
'        ilFound = False
'        ilPos = InStr(1, slName, "'", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        If igExportSource = 2 Then DoEvents
'        ilFound = False
'        ilPos = InStr(1, slName, """", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        If igExportSource = 2 Then DoEvents
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


Private Sub mCSIXMLData(slInCommand As String, slInTag As String, slInData As String)
    Dim slCommand As String
    Dim slTag As String
    Dim slData As String
    Dim ilRet As Integer
    ReDim slFields(0 To 2) As String
    
    If (imGenerating = 1) And (udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked) Then
        Exit Sub
    End If
    '7180 surrounded with bmReExportForce
    If Not bmReExportForce Then
        '6/24/14
        If Not bmAllowXMLCommands Then
            Exit Sub
        End If
    End If
    If igExportSource = 2 Then DoEvents
    slCommand = slInCommand
    slTag = slInTag
    slData = slInData
    If imExportMode = 1 Then
        Select Case slCommand
            Case "OT"   'Open Tag with data
                sgEditValue = "<Tag Data>" & "|" & slTag & "|" & slData & "|"
            Case "CA"   'open/close
                sgEditValue = "<Tag Data />" & "|" & slTag & "|" & slData & "|"
            Case "CD"   'open/close
                sgEditValue = "<Tag>Data</Tag>" & "|" & slTag & "|" & slData & "|"
            Case "CT"   'close only
                sgEditValue = "</Tag>" & "|" & "" & "|" & "" & "|" & slTag
        End Select
        frmXMLTestMode.Show vbModal
        If igAnsCMC = 1 Then
            imExportMode = 0
        End If
        ilRet = gParseItem(sgEditValue, 1, "|", slFields(0))
        ilRet = gParseItem(sgEditValue, 2, "|", slFields(1))
        ilRet = gParseItem(sgEditValue, 3, "|", slFields(2))
        Select Case slCommand
            Case "OT"   'Open Tag
                slTag = slFields(0)
                slData = slFields(1)
            Case "CA"   '
                slTag = slFields(0)
                slData = slFields(1)
            Case "CD"
                slTag = slFields(0)
                slData = slFields(1)
            Case "CT"
                slTag = slFields(2)
        End Select
        csiXMLData slCommand, slTag, slData
    Else
        csiXMLData slCommand, slTag, slData
    End If
    If igExportSource = 2 Then DoEvents
End Sub

Private Sub mExport()
    Dim sNowDate As String
    Dim ilRet As Integer
    Dim slExportType As String
    Dim slXMLFileName As String
    Dim slXMLINIInputFile As String
    Dim slOutputType As String
   ' Dim fs As New FileSystemObject
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilVehicleSelected As Integer
    Dim ilPass As Integer
    Dim ilVff As Integer
    Dim ilISCIForm As Integer
    Dim ilHBForm As Integer
    Dim ilHBPForm As Integer

    Dim sMoDate As String
    Dim sEndDate As String
    Dim slSDate As String
    Dim slEDate As String
    'Dan 2/13/13 show a message that 'not complete' if we had an issue.
    Dim blIssueStation As Boolean
    Dim blIssueAgreement As Boolean
    Dim blIssueSpot As Boolean
    Dim blIssueFile As Boolean
    '5/3/12: Allow two X-Digital exports
    Dim slSection As String
    '6082
    Dim blIssueAstFile As Boolean
    '6741
    Dim blIssueVehicle As Boolean
    '6635
    Dim slRet As String
    '7558
    Dim slSelected As String
    
    Dim slProp As String
    
    On Error GoTo ErrHand
    If imExporting Then
        Exit Sub
    End If
    imExporting = True
    'dan 3/13/13
    imTerminate = False
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    bmAstFileError = False
    '6966 message for user if did a reexport and was successful
    bmAlertAboutReExport = False
    '5896
    blIssueStation = False
    blIssueAgreement = False
    blIssueSpot = False
    blIssueFile = False
    blIssueAstFile = False
    blIssueVehicle = False
    '5896 look here, unless we got an error, then change
    bmIsError = False
    bmIsWrongServicePage = False
    '6/24/14
    bmAllowXMLCommands = True
    '11/01/10 Dan M gXmlIniPath returns xml.ini after testing in different folders. Also, if inipath doesn't exist, exit sub added
    ' NOT DONE
'    slXMLINIInputFile = sgStartupDirectory & "\xml.ini"
'    If Not fs.FileExists(slXMLINIInputFile) Then
'        MsgBox "XML.ini file is missing [" & slXMLINIInputFile & "]"
'    End If
    slXMLINIInputFile = gXmlIniPath(True)
    If LenB(slXMLINIInputFile) = 0 Then
        imExporting = False
        Beep
        gMsgBox "XML.ini file is missing [" & slXMLINIInputFile & "]", vbCritical
        Exit Sub
    End If
    If (udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked) And (udcCriteria.XExportType(1, "V") = vbUnchecked) Then
        imExporting = False
        Beep
        gMsgBox "Please Specify Export Type (Spot Insertions and/or File Delivery).", vbCritical
        Exit Sub
    End If
    If edcDate.Text = "" Then
        imExporting = False
        gMsgBox "Date must be specified.", vbOKOnly
        edcDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcDate.Text) = False Then
        imExporting = False
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        edcDate.SetFocus
        Exit Sub
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
        '12/20/14: Remove XHT records regardless of agreements that are two or more weeks old
        ''Dan M 11/7/14 moved here from load. Help with testing dates in past
        'smXHTDeleteDate = gObtainStartStd(Format(gDateValue(gObtainStartStd(smDate)) - 1, "m/d/yy"))
        '5/21/15 moved to local
        'smXHTDeleteDate = Format(gDateValue(gObtainPrevMonday(smDate)) - 14, "m/d/yy")
    End If
    '3/23/13: allow export of todays date
    sNowDate = Format$(gNow(), "m/d/yy")
    If udcCriteria.XGenType(0, slProp) Then
        If gDateValue(gAdjYear(smDate)) < gDateValue(gAdjYear(sNowDate)) Then
            imExporting = False
            Beep
            gMsgBox "Date must be today's date or later " & sNowDate, vbCritical
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        imExporting = False
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    Select Case Weekday(gAdjYear(smDate))
        Case vbMonday
            If imNumberDays > 7 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 7.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbTuesday
            If imNumberDays > 6 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 6.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbWednesday
            If imNumberDays > 5 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 5.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbThursday
            If imNumberDays > 4 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 4.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbFriday
            If imNumberDays > 3 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 3.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbSaturday
            If imNumberDays > 2 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 2.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbSunday
            If imNumberDays > 1 Then
                imExporting = False
                gMsgBox "Number of days can not exceed 1.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
    End Select
    If (udcCriteria.XSpots(ALLSPOTS) = False) And (udcCriteria.XSpots(REGIONALONLY) = False) Then
        imExporting = False
        Beep
        gMsgBox "Please Specify Export Spots Type (All or Regional Spots).", vbCritical
        Exit Sub
    End If
    If (udcCriteria.XProvider(0) = False) And (udcCriteria.XProvider(1) = False) Then
        If UBound(sgXDSSection) > 1 Then
            imExporting = False
            Beep
            gMsgBox "Please Specify Export Provider.", vbCritical
            Exit Sub
        Else
            slSection = ""
        End If
    Else
        'get name 'XDigital-CU'
        If (udcCriteria.XProvider(1) = True) Then
            slSection = Mid(sgXDSSection(1), 2, Len(sgXDSSection(1)) - 2)
        Else
            slSection = Mid(sgXDSSection(0), 2, Len(sgXDSSection(0)) - 2)
        End If
    End If
    If slSection = "" Then
        slSection = "XDigital"
    End If
    ilVehicleSelected = False
    ilISCIForm = False
    ilHBForm = False
    ilHBPForm = False
    'ilRet = gPopVff()
    '7558
    slSelected = ""
'    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
'        If igExportSource = 2 Then DoEvents
'        If lbcVehicles.Selected(ilVef) Then
    For ilVef = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilVef, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilVef, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilVef, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilVef, VEHINDEX)
                ilVehicleSelected = True
                'imVefCode = lbcVehicles.ItemData(ilVef)
                ilVff = gBinarySearchVff(imVefCode)
                ilVpf = gBinarySearchVpf(CLng(imVefCode))
                If (ilVff <> -1) And (ilVpf <> -1) Then
                    If igExportSource = 2 Then DoEvents
                    'slSelected = slSelected & Replace(lbcVehicles.List(ilVef), ",", "_") & ","
                    slSelected = slSelected & Replace(smVefName, ",", "_") & ","
                    If tgVpfOptions(ilVpf).iInterfaceID > 0 Then
                        ilISCIForm = True
                    End If
                    Select Case Trim$(tgVffInfo(ilVff).sXDXMLForm)
                        Case "A"    'HBP form
                            ilHBPForm = True
                        Case "S"        'HB form
                            ilHBForm = True
                    End Select
                End If
            End If
        End If
    Next ilVef
    If (Not ilVehicleSelected) Then
        imExporting = False
        Beep
        gMsgBox "Vehicle must be selected.", vbCritical
        Exit Sub
    End If
    '6635  myFile for reading Jeff's text file
    Set myFile = New FileSystemObject
    smExportDirectory = udcCriteria.XExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    If Not myFile.FolderExists(smExportDirectory) Then
        smExportDirectory = sgExportDirectory
        mSetResults "Chosen directory does not exist.  Export files will be written to generic export folder.", MESSAGEBLACK
       ' gLogMsg "Chosen directory " & udcCriteria.XExportToPath & " does not exist.  Export files will be written to " & smExportDirectory, smPathForgLogMsg, False
        myExport.WriteWarning "Chosen directory " & udcCriteria.XExportToPath & " does not exist.  Export files will be written to " & smExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
    '6807
    smXmlErrorFile = gGetXmlErrorFile(smXmlFileFindError)
    If Len(smXmlErrorFile) = 0 Then
       ' gLogMsg "Warning, cannot read errors from web server: " & smXmlFileFindError, smPathForgLogMsg, False
        myExport.WriteError "Cannot read errors from web server: " & smXmlFileFindError
    End If
    '6635, 6581
    imChunk = 5000
    slRet = "Not Found"
    On Error Resume Next
    gLoadFromIni slSection, "AuthorizationChunkSize", slXMLINIInputFile, slRet
    If slRet <> "Not Found" Then
        If IsNumeric(slRet) Then
            imChunk = CInt(slRet)
        End If
    End If
    '6966
    imMaxRetries = 3
    gLoadFromIni slSection, "Retries", slXMLINIInputFile, slRet
    If slRet <> "Not Found" Then
        If IsNumeric(slRet) Then
            imMaxRetries = CInt(slRet)
        End If
    End If
    On Error GoTo ErrHand:
    Screen.MousePointer = vbHourglass
    imExporting = True
    mSaveCustomValues
    'ttp 5457
    lacProcessing.Visible = True
    '6806 I need for stations and agreements to know which id to send
    'xmlReceiverIdSource...is there one in xml? set the boolean sends
    '9256 made global
    If Not gStationXmlReceiverChoices(slSection, slXMLINIInputFile, bmSendStationIds, bmSendAgreementIds) Then
        'single provider and no value means normal export
        If UBound(sgXDSSection) > 1 Then
        Else
            bmSendAgreementIds = True
            bmSendStationIds = True
        End If
    End If
    If mnuNormal.Checked Or mnuStation.Checked Then
        If Not mStationSend(slXMLINIInputFile, slSection) Then
             '5896
            blIssueStation = True
            If bmIsError Then
                 'gLogMsg "** Terminated - Error in sending Stations **", smPathForgLogMsg, False
                 myExport.WriteError "** Terminated - Error in sending Stations **", False
                 '5/7/15 Dan
                 'Exit Sub
                 GoTo Cleanup
            ElseIf bmIsWrongServicePage Then
                'gLogMsg "Attempted to send Stations to wrong service page", smPathForgLogMsg, False
                myExport.WriteError "Attempted to send Stations to wrong service page", False
             End If
    
     '       Beep
            'gMsgBox "Could not export stations. Export halted", vbCritical
     '       Call mSetResults("Station Information Failed", RGB(255, 0, 0))
    '        imExporting = False
    '        Screen.MousePointer = vbDefault
    '        cmdCancel.SetFocus
    '        Exit Sub
        End If
    End If
    If mnuNormal.Checked Or mnuProgram.Checked Then
        'ttp 6741
        If Not mVehicleSend(slXMLINIInputFile, slSection) Then
            blIssueVehicle = True
            If bmIsError Then
                ' gLogMsg "** Terminated - Error in sending Vehicles **", smPathForgLogMsg, False
                myExport.WriteError "** Terminated - Error in sending Vehicles **", False
                '5/7/15 Dan
                'Exit Sub
                GoTo Cleanup
             ElseIf bmIsWrongServicePage Then
                'gLogMsg "Attempted to send Vehicles to wrong service page", smPathForgLogMsg, False
                myExport.WriteError "Attempted to send Vehicles to wrong service page", False
             End If
        End If
    End If
    '7437
    If (mnuNormal.Checked Or mnuAuthorization.Checked) And udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
'    If mnuNormal.Checked Or mnuAuthorization.Checked Then
        'ttp 5589
        If Not mAgreementSend(slXMLINIInputFile, slSection) Then
             blIssueAgreement = True
            If bmIsError Then
                 'gLogMsg "** Terminated - Error in sending Agreements **", smPathForgLogMsg, False
                 myExport.WriteError "** Terminated - Error in sending Agreements **", False
                 '5/7/15 Dan
                 'Exit Sub
                 GoTo Cleanup
             End If
    
      '      Beep
           ' gMsgBox "Could not export Agreements. Export halted", vbCritical
    '        imExporting = False
    '        Screen.MousePointer = vbDefault
    '        cmdCancel.SetFocus
    '        Exit Sub
        End If
    End If
    If mnuNormal.Checked = False Then
        imExporting = False
        mSetResults "Guide has chosen to run only one delivery", MESSAGEBLACK
        Screen.MousePointer = vbDefault
        GoTo Cleanup
        'Exit Sub
    End If
    If Not gPopCopy(smDate, "Export X-Digital") Then
        lacProcessing.Visible = False
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        '5/7/15 Dan
        'imExporting = False
        'Exit Sub
        GoTo Cleanup
    End If
    ReDim tmXDFDInfo(0 To 0) As XDFDINFO
    '6082
    Set rsAstFiles = mPrepRecordset()
    On Error GoTo 0
    lacResult.Caption = ""
    '6632
'    If udcCriteria.XSpots(0) = True Then
'        slExportType = "!! Exporting All Spots "
'    Else
'        slExportType = "!! Exporting Regional Spots "
'    End If
    If udcCriteria.XSpots(ALLSPOTS) = True Then
        slExportType = "All Spots "
    Else
        slExportType = "Regional Spots "
    End If
    If udcCriteria.XGenType(0, slProp) Then
        slExportType = "!! Exporting " & slExportType & ","
    Else
        slExportType = "!! Writing " & slExportType & "in test mode,"
        mSetResults "Test mode, not sending to XDS.", RGB(0, 0, 0)
    End If
    'gLogMsg slExportType & " Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", smPathForgLogMsg, False
    myExport.WriteFacts slExportType & " Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
    '7558
    If chkAll.Value = vbChecked Then
        slSelected = "All"
    Else
        slSelected = mLoseLastLetterIfComma(slSelected)
        If InStr(slSelected, ",") = 0 Then
            If chkAllStation.Value = vbChecked Then
                slSelected = slSelected & " and All Stations."
            Else
                slSelected = slSelected & " and these Stations: "
                For ilVef = 0 To lbcStation.ListCount - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If lbcStation.Selected(ilVef) Then
                        slSelected = slSelected & Replace(lbcStation.List(ilVef), ",", "_") & ","
                    End If
                Next ilVef
            End If
            slSelected = mLoseLastLetterIfComma(slSelected)
        End If
    End If
    If Len(slRet) > 0 Then
        myExport.WriteFacts "Exporting:" & slSelected
    End If
    '6/24/14: Initialize previously sent record info
    Set xhtInfo_rst = mInitXHTInfo()
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .ThirdParty = Vendors.XDS_Break
        .TypeEnt = Exportunposted3rdparty
        '10099  'udcCriteriaXSpots 0 all spots  1 regional only
        If udcCriteria.XSpots(0) Then
            .XdsAllOrRegional = XdsAllOrRegionalEnum.All
        Else
            .XdsAllOrRegional = XdsAllOrRegionalEnum.Regional
        End If
        .User = igUstCode
        .ErrorLog = smPathForgLogMsg
        .CreateUniqueFilenames = True
        If Len(.ErrorMessage) > 0 Then
           myExport.WriteWarning "May not be able to create ent records. " & .ErrorMessage, True
        End If
    End With
    '8685
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    For ilPass = 0 To 2 Step 1
        If igExportSource = 2 Then DoEvents
        If ((ilPass = 0) And (ilISCIForm)) Or ((ilPass = 1) And (ilHBPForm)) Or ((ilPass = 2) And (ilHBForm)) Then
            Select Case ilPass
                Case 0
                    slXMLFileName = "XD-SI-" & Format$(smDate, "yyyymmdd") & "-ISCI_Form" & ".xml"
                    'gLogMsg "Generating Spot Insertion: ISCI Form", smPathForgLogMsg, False
                    myExport.WriteFacts "Generating Spot Insertion: ISCI Form"
                Case 1
                    slXMLFileName = "XD-SI-" & Format$(smDate, "yyyymmdd") & "-HBP_Form" & ".xml"
                   ' gLogMsg "Generating Spot Insertion: HBP Form", smPathForgLogMsg, False
                    myExport.WriteFacts "Generating Spot Insertion: HBP Form"
                Case 2
                    slXMLFileName = "XD-SI-" & Format$(smDate, "yyyymmdd") & "-HB_Form" & ".xml"
                    'gLogMsg "Generating Spot Insertion: HB Form", smPathForgLogMsg, False
                    myExport.WriteFacts "Generating Spot Insertion: HB Form"
            End Select
            'T=transmit; F=File; B=both
            'slOutputType = "B"
            'ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "XDigital", slOutputType, sgExportDirectory & slXMLFileName, "")
            'Dan M 11/01/10 use slXMLINIInputFile
            If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                If udcCriteria.XGenType(0, slProp) Then
                    slOutputType = "T"
                    ''ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "XDigital", slOutputType, sgExportDirectory & slXMLFileName, "")
                    'ilRet = csiXMLStart(slXMLINIInputFile, "XDigital", slOutputType, smExportDirectory & slXMLFileName, "")
                    ilRet = csiXMLStart(slXMLINIInputFile, slSection, slOutputType, smExportDirectory & slXMLFileName, "", smXmlErrorFile)
                Else
                    slOutputType = "F"
                    ''ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "XDigital", slOutputType, sgExportDirectory & slXMLFileName, sgCRLF)
                    'ilRet = csiXMLStart(slXMLINIInputFile, "XDigital", slOutputType, smExportDirectory & slXMLFileName, sgCRLF)
                    ilRet = csiXMLStart(slXMLINIInputFile, slSection, slOutputType, smExportDirectory & slXMLFileName, sgCRLF, smXmlErrorFile)
                End If
                If ilPass = 0 Then
                    '6979 change for deletes
                    Call csiXMLSetMethod("SetInsertions", "", "225", "")
                   ' Call csiXMLSetMethod("SetInsertions", "Sites", "225", "")
                Else
                    '6979 change for deletes
                    Call csiXMLSetMethod("SetInsertions", "", "225", "")
                    'Call csiXMLSetMethod("SetInsertions",  "Insertions", "225", "")
                End If
            End If
            ilRet = mExportSpotInsertions(ilPass, slSection)
            '7/6/14: Allow XML
            bmAllowXMLCommands = True
            gCloseRegionSQLRst
            If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
                csiXMLEnd
            Else
                ilRet = True
            End If
            '2/13/13: Dan doing differently
            '2/7/13: Ignore mExportSpotInsertion return
            'ilRet = True
            If imTerminate Then
                lacProcessing.Visible = False
                Call mSetResults("** User Terminated **", MESSAGERED)
                'gLogMsg "** User Terminated **", smPathForgLogMsg, False
                myExport.WriteWarning "** User Terminated **"
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                '5/7/15 dan
                'Exit Sub
                GoTo Cleanup
            End If
            If (ilRet = False) Then
                blIssueSpot = True
                'severe error--get out
                If bmIsError Then
                    'moved to mIsXmlError
'                    lacProcessing.Visible = False
'                    lacResult.Caption = "Check 'XDigitalExportLog.Txt' in Messages folder for issue."
'                   ' Call mSetResults("Export not completely successful", RGB(255, 0, 0))
'                    gLogMsg "** Terminated - mExportSpotInsertions returned False **", "XDigitalExportLog.Txt", False
'                    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
'                    imExporting = False
'                    Screen.MousePointer = vbDefault
'                    cmdCancel.SetFocus
                   ' gLogMsg "** Terminated - Error in sending Spot Insertions **", smPathForgLogMsg, False
                    myExport.WriteError "** Terminated - Error in sending Spot Insertions **", False
                    '5/7/15 dan
                    'Exit Sub
                    GoTo Cleanup
                End If
            End If
        End If
    Next ilPass
    '5/21/15 Dan moved here from mExportSpotInsertions
    mRemoveOldXHT
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbChecked Then
        If udcCriteria.XGenType(0, slProp) Then
            sMoDate = gObtainPrevMonday(smDate)
            sEndDate = DateAdd("d", imNumberDays - 1, smDate)
            slSDate = smDate
            slEDate = gObtainNextSunday(slSDate)
            If gDateValue(gAdjYear(sEndDate)) < gDateValue(gAdjYear(slEDate)) Then
                slEDate = sEndDate
            End If
            mClearAlerts gDateValue(slSDate), gDateValue(slEDate)
        End If
    End If
    'If ckcExportType(1).Value = vbChecked Then
    If (udcCriteria.XExportType(1, "V") = vbChecked) And (UBound(tmXDFDInfo) > LBound(tmXDFDInfo)) Then
        'gLogMsg "Generating File Delivery", smPathForgLogMsg, False
        myExport.WriteFacts "Generating File Delivery"
        imGenerating = 2
        slXMLFileName = "XD-FD-" & Format$(smDate, "yyyymmdd") & ".xml"
        'Dan M use slXMLINIInputFile
        If udcCriteria.XGenType(0, slProp) Then
            slOutputType = "T"
            ''ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "XDigital", slOutputType, sgExportDirectory & slXMLFileName, "")
            'ilRet = csiXMLStart(slXMLINIInputFile, "XDigital", slOutputType, smExportDirectory & slXMLFileName, "")
            ilRet = csiXMLStart(slXMLINIInputFile, slSection, slOutputType, smExportDirectory & slXMLFileName, "", smXmlErrorFile)
        Else
            slOutputType = "F"
            ''ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "XDigital", slOutputType, sgExportDirectory & slXMLFileName, sgCRLF)
            'ilRet = csiXMLStart(slXMLINIInputFile, "XDigital", slOutputType, smExportDirectory & slXMLFileName, sgCRLF)
            ilRet = csiXMLStart(slXMLINIInputFile, slSection, slOutputType, smExportDirectory & slXMLFileName, sgCRLF, smXmlErrorFile)
        End If
        Call csiXMLSetMethod("SetFileDeliveryPackages", "FileDeliveryPackages", "225", "")
        ilRet = mExportFileDelivery()
        csiXMLEnd
        If imTerminate Then
            lacProcessing.Visible = False
            Call mSetResults("** User Terminated **", MESSAGERED)
            'gLogMsg "** User Terminated **", smPathForgLogMsg, False
            myExport.WriteWarning "** User Terminated **"
            ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
            '5/7/15 dan
'            imExporting = False
'            Screen.MousePointer = vbDefault
'            cmdCancel.SetFocus
            'Exit Sub
            GoTo Cleanup
        End If
        'dan 3/13/13
        If (ilRet = False) Then
            blIssueFile = True
           ' lacProcessing.Visible = False
           ' Call mSetResults("Export not completely successful", RGB(255, 0, 0))
           If bmIsError Then
                'gLogMsg "** Terminated - Error in sending File Delivery **", smPathForgLogMsg, False
                myExport.WriteError "** Terminated - Error in sending File Delivery **", False
                '5/7/15 dan
                'Exit Sub
                GoTo Cleanup
            End If
          '  ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
'            If slOutputType <> "T" Then
'                lacResult.Caption = "Exports placed into: " & smExportDirectory
'            Else
'                lacResult.Caption = "Check XDigitalIgnoredErrors_mm_dd_yy.txt in Messages folder for possible Warnings like 'Invalid Site ID'"
'            End If
'            imExporting = False
'            Screen.MousePointer = vbDefault
'            cmdCancel.SetFocus
'            Exit Sub
        End If
    Else
        If (udcCriteria.XExportType(1, "V") = vbChecked) Then
           ' gLogMsg "No Files found to be Export in File Delivery", smPathForgLogMsg, False
            myExport.WriteFacts "No Files found to export in File Delivery"
        Else
            'gLogMsg "File Delivery not Selected", smPathForgLogMsg, False
            myExport.WriteFacts "File Delivery not Selected"
        End If
    End If
    '6082
    '7/9/13: Transparent file generatrion is independent of the Unit ID setting
    'If smUnitIdByAstCode = "Y" Then
    If smGenTransparency = "Y" Then
        DoEvents
        lacProcessing.Caption = "Creating Transparency file."
        blIssueAstFile = gAstFileWrite(FrmExportXDigital, rsAstFiles, smDate, smExportDirectory, bmAstFileError)
    End If
    On Error GoTo ErrHand:
    lacProcessing.Visible = False
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    '5/7/15 Dan get it lower
   ' imExporting = False
    'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
Cleanup:
    'gLogMsg "** Completed Export of X-Digital **", smPathForgLogMsg, False
    myExport.WriteFacts "** Completed Export of X-Digital **"
    If Not (blIssueAgreement Or blIssueSpot Or blIssueFile Or blIssueStation Or blIssueAstFile Or blIssueVehicle) Then
'        Call mSetResults("Export Completed Successfully", RGB(0, 155, 0))
        '6966
        If bmAlertAboutReExport Then
            mSetResults "Export completed Successfully, but reexport was needed. See log for the issue.", MESSAGEGREEN
        Else
            Call mSetResults("Export Completed Successfully", RGB(0, 155, 0))
        End If
    Else
        mSetResults "Export completed with issues", MESSAGERED
        If blIssueStation Then
            mSetResults "Station Delivery had an issue", MESSAGERED
        End If
        If blIssueAgreement Then
             mSetResults "Authorization Delivery had an issue", MESSAGERED
        End If
        If blIssueSpot Then
             mSetResults "Spot Insertions had an issue", MESSAGERED
        End If
        If blIssueFile Then
             mSetResults "File Delivery had an issue", MESSAGERED
        End If
        If blIssueAstFile Then
            mSetResults "Transparency file had an issue", MESSAGERED
        End If
        '6741
        If blIssueVehicle Then
             mSetResults "Program Delivery had an issue", MESSAGERED
        End If
    End If
    'Close #hmMsg
    If slOutputType <> "T" Then
        lacResult.Caption = "Exports placed into: " & smExportDirectory
    'dan added the test
    ElseIf blIssueAgreement Or blIssueSpot Or blIssueFile Or blIssueStation Or blIssueVehicle Then
        mSetResults "See 'IgnoredErrors' file as shown below for a list of issues", MESSAGERED
        lacResult.Caption = "Check 'XDigitalIgnoredErrors_mm_dd_yy.txt' in Messages folder for issues (ex:'Invalid Site ID')"
    End If
    '6082
    If slOutputType = "T" And Not blIssueAstFile Then
        lacResult.Caption = "Exports placed into: " & smExportDirectory
    End If
    '8685
    If bgTaskBlocked And igExportSource <> 2 Then
         mSetResults "Some spots were blocked during export.", MESSAGERED
         gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
         myExport.WriteWarning "Some spots were blocked during export.", True
         lacResult.Caption = "Please refer to the Messages folder for file TaskBlocked_" & sgTaskBlockedDate & ".txt."
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    '6082
    If Not rsAstFiles Is Nothing Then
        If (rsAstFiles.State And adStateOpen) <> 0 Then
            rsAstFiles.Close
        End If
        Set rsAstFiles = Nothing
    End If
    Set myFile = Nothing
    cmdExport.Enabled = False
    cmdExportTest.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
   ' gLogMsg "", smPathForgLogMsg, False
    myExport.WriteFacts "", True
    '7458
    Set myEnt = Nothing
    '5/7/15 Dan
    imExporting = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mExport"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    imExporting = False
    Exit Sub

End Sub
Private Function mVehicleSend(slXMLINIInputFile As String, ByVal slSection As String) As Boolean
'6741
' return false if something didn't go right.
' set bmIsError to true if error and nothing could go out.  Will stop other exports!
    Dim blRet As Boolean
    Dim blSendToWeb As Boolean
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilVff As Integer
    Dim ilVefCode As Integer
    Dim slSql As String
    Dim ilID As Integer
    Dim tlProgram() As XDIGITALVEHICLEINFO
    Dim ilUpper As Integer
    Dim c As Integer
    Dim blError As Boolean
    Dim ilVefIndex As Integer
    Dim slProp As String
    
    ilID = 0
    ilUpper = 0
    ReDim tlProgram(0)
    blRet = True
    ' only Cumulus send vehicle.  Testing webserviceUrl in xml.ini
    If mCumulusHeadEnd(slSection, slXMLINIInputFile, blError) Then
        slSql = "select siteProgToXds as Program from site"
    On Error GoTo ERRORBOX
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            If rst!Program = "Y" Then
                If Len(smXmlErrorFile) = 0 Then
                    mVehicleSend = False
                    mSetResults "Problem sending vehicles: error file doesn't exist.", MESSAGERED
                    myExport.WriteError "Problem Sending vehicles: error file doesn't exist. Make sure the xml.ini 'logfile' has a valid directory.", False, False
                    Exit Function
                End If
                If udcCriteria.XGenType(0, slProp) Then
                    blSendToWeb = True
                    myExport.WriteFacts "!! Exporting Vehicles  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
                    mSetResults "Exporting Vehicle Information", MESSAGEBLACK
                Else
                    blSendToWeb = False
                    myExport.WriteFacts "!! Writing Vehicles in test mode  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
                    mSetResults "Writing Vehicle Information to VehicleInformation.txt", MESSAGEBLACK
                End If
'                For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
'                    If igExportSource = 2 Then DoEvents
'                    If lbcVehicles.Selected(ilVef) Then
'                        ilVefCode = lbcVehicles.ItemData(ilVef)
                For ilVef = 1 To grdVeh.Rows - 1 Step 1
                    If Trim(grdVeh.TextMatrix(ilVef, VEHINDEX)) <> "" Then
                        If grdVeh.TextMatrix(ilVef, SELECTEDINDEX) = "1" Then
                            '8608 ilVefCode was incorrectly imVefCode
                            ilVefCode = grdVeh.TextMatrix(ilVef, VEHCODEINDEX)
                            smVefName = grdVeh.TextMatrix(ilVef, VEHINDEX)
                            ilVff = gBinarySearchVff(ilVefCode)
                            If ilVff <> -1 Then
                               ' If tgVffInfo(ilVff).sXDXMLForm = "P" Then
                                    'Need to send?
                                If tgVffInfo(ilVff).sSentToXDS = "M" Or tgVffInfo(ilVff).sSentToXDS = "N" Then
                                    '6836 don't send games
                                    ilVefIndex = gBinarySearchVef(CLng(ilVefCode))
                                    If ilVefIndex <> -1 Then
                                        If tgVehicleInfo(ilVefIndex).sVehType <> "G" Then
                                            ilVpf = gBinarySearchVpf(CLng(ilVefCode))
                                            If ilVpf <> -1 Then
                                                ilID = tgVpfOptions(ilVpf).iInterfaceID
                                                'ID > 0?
                                                If ilID > 0 Then
                                                    ilUpper = UBound(tlProgram)
                                                    With tlProgram(ilUpper)
                                                        .iNetworkID = ilID
                                                        .iCode = ilVefCode
                                                        '.sName = gXMLNameFilter(lbcVehicles.List(ilVef))
                                                        .sName = gXMLNameFilter(smVefName)
                                                    End With
                                                    If blSendToWeb Then
                                                        'gLogMsg "Sending vehicle " & lbcVehicles.List(ilVef), smPathForgLogMsg, False
                                                        gLogMsg "Sending vehicle " & smVefName, smPathForgLogMsg, False
                                                    Else
                                                       ' gLogMsg "Test send vehicle " & lbcVehicles.List(ilVef), smPathForgLogMsg, False
                                                        'myExport.WriteFacts "Test send vehicle " & lbcVehicles.List(ilVef)
                                                        myExport.WriteFacts "Test send vehicle " & smVefName
                                                    End If
                                                    ilUpper = ilUpper + 1
                                                    ReDim Preserve tlProgram(ilUpper)
                                                Else
                                                   ' mSetResults lbcVehicles.List(ilVef) & " missing vehicle ID", MESSAGEBLACK
                                                   '' gLogMsg lbcVehicles.List(ilVef) & " missing vehicle ID", smPathForgLogMsg, False
                                                   ' myExport.WriteWarning lbcVehicles.List(ilVef) & " missing vehicle ID"
                                                End If
                                            Else
                                                'mSetResults lbcVehicles.List(ilVef) & " missing vehicle ID", MESSAGEBLACK
                                                mSetResults smVefName & " missing vehicle ID", MESSAGEBLACK
                                               ' gLogMsg lbcVehicles.List(ilVef) & " missing vehicle ID (no vpf options)", smPathForgLogMsg, False
                                                'myExport.WriteWarning lbcVehicles.List(ilVef) & " missing vehicle ID (no vpf options)"
                                                myExport.WriteWarning smVefName & " missing vehicle ID (no vpf options)"
                                            End If ' couldn't find in options
                                        End If 'game?
                                    Else
                                         'mSetResults lbcVehicles.List(ilVef) & " missing vehicle information", MESSAGEBLACK
                                         mSetResults smVefName & " missing vehicle information", MESSAGEBLACK
                                         'myExport.WriteWarning lbcVehicles.List(ilVef) & " missing vehicle information "
                                         myExport.WriteWarning smVefName & " missing vehicle information "
                                    End If 'couldn't find vehicle
                                End If  'need to send
    '                        Else
    '                            blRet = False
    '                            mSetResults "Problem sending vehicles: info doesn't exist for " & lbcVehicles.List(ilVef), MESSAGERED
    '                            gLogMsg "Problem Sending vehicles: info doesn't exist for " & lbcVehicles.List(ilVef), "XDigitalExportLog.Txt", False
                            End If 'couldn't find in vff
                        End If
                    End If 'selected?
                Next ilVef
                If ilUpper > 0 Then
                    If mVehicleSendFacts(slXMLINIInputFile, blSendToWeb, slSection, tlProgram) Then
                    Else
                        blRet = False
                    End If
                Else
                    mSetResults "No Vehicles need to be sent", MESSAGEBLACK
                End If
            End If
        Else
            blRet = False
            mSetResults "Problem sending vehicles: Couldn't access site options.", MESSAGERED
            myExport.WriteError "Problem Sending vehicles: Couldn't access site options.", False, False
        End If
    Else
        If blError Then
            blRet = False
            mSetResults "Problem sending Vehicles: Can't read WebServiceURL in xml.ini.", MESSAGERED
            myExport.WriteError "Problem Sending Vehicles: Can't read WebServiceURL in xml.ini at " & slXMLINIInputFile, False, False
        End If
    End If
Cleanup:
    mVehicleSend = blRet
    Erase tlProgram
    Exit Function
ERRORBOX:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mVehicleSend"
    blRet = False
    GoTo Cleanup
End Function
'Private Function mStationXmlReceiverChoices(slSection As String, slXMLINIInputFile As String, blStation As Boolean, blAgreement As Boolean) As Boolean
'    'returns true if field exists, otherwise false
'    'O: blstation and blagreement
'    Dim blRet As Boolean
'    Dim slRet As String
'
'    slRet = ""
'    blRet = False
'    blStation = False
'    blAgreement = False
'    gLoadFromIni slSection, STATIONXMLRECEIVERID, slXMLINIInputFile, slRet
'    If slRet <> "Not Found" Then
'        blRet = True
'        Select Case slRet
'            Case "A"
'                blAgreement = True
'            Case "B"
'                blAgreement = True
'                blStation = True
'            Case "S"
'                blStation = True
'        End Select
'    End If
'    mStationXmlReceiverChoices = blRet
'End Function
Private Function mCumulusHeadEnd(slSection As String, slXMLINIInputFile As String, blIsError As Boolean) As Boolean
    Dim slRet As String
    Dim blRet As Boolean
    
    slRet = ""
    blRet = False
    blIsError = False
    gLoadFromIni slSection, "WebServiceURL", slXMLINIInputFile, slRet
    If slRet <> "Not Found" Then
        If InStr(1, slRet, "abcdService", vbTextCompare) > 0 Then
            blRet = True
        End If
    Else
        blIsError = True
    End If
    mCumulusHeadEnd = blRet
End Function

Private Function mVehicleSendFacts(slXMLINIInputFile As String, blSendToWeb As Boolean, slSection As String, tlVehicles() As XDIGITALVEHICLEINFO) As Boolean
    Dim llCount As Long
    Dim slDoNotUpdate As String
    Dim slNeedUpdate As String
    Dim c As Integer
    Dim blRet As Boolean
    Dim blTestUpdate As Boolean
    Dim slBadVehicles As String
    '5/7/15 need to store sldonotupdate
    Dim slDoNotMaster As String
    
    blTestUpdate = bmTestForceUpdateXHT
    blRet = True
    slDoNotMaster = ""
    slNeedUpdate = ""
    bmFailedToReadReturn = False
    If Not blSendToWeb Then
        csiXMLStart slXMLINIInputFile, slSection, "F", smExportDirectory & VEHICLELOG & smDateForLogs & ".txt", sgCRLF, smXmlErrorFile
    Else
        csiXMLStart slXMLINIInputFile, slSection, "T", "", "", smXmlErrorFile
    End If
    csiXMLSetMethod "SetAiringNetworks", "NewDataSet", "225", ""
    For c = 0 To UBound(tlVehicles) - 1
        If igExportSource = 2 Then DoEvents
        '6581 added  hit the number to send at one time?  send and start anew.
        llCount = llCount + 1
        slNeedUpdate = slNeedUpdate & tlVehicles(c).iCode & ","
        If llCount Mod imChunk = 0 Then
            '7508
           ' If Not mVehicleSendAndTest(slDoNotUpdate) Then
            If Not mSendAndTestReturn(Vehicles, slDoNotUpdate) Then
                blRet = False
                slDoNotMaster = slDoNotMaster & slDoNotUpdate
                If bmIsError Then
                    csiXMLEnd
                    GoTo Cleanup
                End If
            End If
            csiXMLSetMethod "SetAiringNetworks", "SetAiringNetworks", "225", ""
        End If
        If mVehicleWriteXml(tlVehicles(c)) Then
        Else
            blRet = False
            If bmIsError Then
                csiXMLEnd
                GoTo Cleanup
            End If
        End If
    Next c
     '7508
    ' If Not mVehicleSendAndTest(slDoNotUpdate) Then
     If Not mSendAndTestReturn(Vehicles, slDoNotUpdate) Then
        blRet = False
        slDoNotMaster = slDoNotMaster & slDoNotUpdate
        If bmIsError Then
            csiXMLEnd
            GoTo Cleanup
        End If
    End If
    csiXMLEnd
    If (blSendToWeb Or blTestUpdate) And Not bmIsWrongServicePage Then
        'handles if one of multiple didn't work.
        If bmFailedToReadReturn Then
            slDoNotMaster = ""
            bmFailedToReadReturn = False
        End If
        '5/7/15 Dan  warning, but no 'doNotUpdate'?  then none get updated!
        If blRet Or Len(slDoNotMaster) > 0 Then
            'xds returns their program code, I need the matching counterpoint code
            slBadVehicles = mVehicleAdjustDontUpdate(slDoNotMaster, tlVehicles)
            If Len(slBadVehicles) > 0 Then
               myExport.WriteWarning "XDS did not accept these vehicles: " & slBadVehicles
            End If
            slNeedUpdate = mAdjustUpdates(slNeedUpdate, slDoNotMaster)
            If Len(slNeedUpdate) > 0 Then
                mVehicleUpdate (slNeedUpdate)
            End If
        ' this means 'doNotUpdate' was blank
        ElseIf Not blRet Then
            myExport.WriteWarning "XDS did not accept some vehicles.  The return could not be parsed; No vehicles were marked as 'Sent'."
        End If
    End If
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Vehicle Information Completed. Sent: " & llCount, MESSAGEBLACK
           myExport.WriteFacts "Vehicle Information Completed. Sent: " & llCount
        End If
    Else
        mSetResults "Vehicle Information Completed. Test send: " & llCount, MESSAGEBLACK
        myExport.WriteFacts "Vehicle Information Completed. Test send: " & llCount
    End If
    mVehicleSendFacts = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "mVehicleSendFacts"
    GoTo Cleanup
End Function
Private Function mVehicleAdjustDontUpdate(slDoNotUpdate As String, tlVehicles() As XDIGITALVEHICLEINFO) As String
    'change the NetworkCode to the vefcode. Return the vehicles name for the message file.
    Dim slRet As String
    Dim slDont As String
    Dim slDonts() As String
    Dim c As Integer
    Dim j As Integer
    Dim slFindThis As String
    Dim slRightCode As String
    
    slRet = ""
    If Len(slDoNotUpdate) > 0 Then
        slRet = ""
        slDont = mLoseLastLetter(slDoNotUpdate)
        If Len(slDont) > 0 Then
            slDoNotUpdate = ""
            slDonts = Split(slDont, ",")
            mSetResults "issues with " & UBound(slDonts) + 1 & " vehicles.", MESSAGERED
            For c = 0 To UBound(slDonts)
                slFindThis = slDonts(c)
                For j = 0 To UBound(tlVehicles)
                    If slFindThis = tlVehicles(j).iNetworkID Then
                        slRightCode = tlVehicles(j).iCode
                        slDoNotUpdate = slDoNotUpdate & slRightCode & ","
                        slRet = slRet & tlVehicles(j).sName & ","
                    End If
                Next j
            Next c
        End If
    End If
    slRet = mLoseLastLetter(slRet)
    mVehicleAdjustDontUpdate = slRet
End Function
Private Function mVehicleUpdate(slNeedUpdate As String) As Boolean
    Dim ilVff As Integer
    Dim slCodes() As String
    Dim c As Integer
    Dim ilVef As Integer
    If Len(slNeedUpdate) > 0 Then
        SQLQuery = "update vff_Vehicle_Features set vffSentToXdsStatus = 'Y' where vffvefcode in ( " & slNeedUpdate & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, "mVehicleUpdate"
            mVehicleUpdate = False
            Exit Function
        End If
        slCodes = Split(slNeedUpdate, ",")
        For c = 0 To UBound(slCodes)
            ilVef = Val(slCodes(c))
            ilVff = gBinarySearchVff(ilVef)
            If ilVff <> -1 Then
                tgVffInfo(ilVff).sSentToXDS = "Y"
            End If
        Next c
    End If
    mVehicleUpdate = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "mVehicleUpdate"
    mVehicleUpdate = False
End Function
Private Function mVehicleWriteXml(tlVehicle As XDIGITALVEHICLEINFO) As Boolean
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo ErrHand
    With tlVehicle
        mCSIXMLData "OT", "AiringNetwork", ""
        mCSIXMLData "CD", "NetworkId", CStr(.iNetworkID)
        mCSIXMLData "CD", "NetworkName", .sName
        mCSIXMLData "CD", "CounterpointNetworkId", CStr(.iCode)
        mCSIXMLData "CT", "AiringNetwork", ""
    End With
    DoEvents
    mVehicleWriteXml = blRet
    Exit Function
   
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "mVehicleWriteXml"
    mVehicleWriteXml = False
End Function
'Private Function mVehicleSendAndTest(slDoNotReturn As String) As Boolean
'    'return false if warning or error
'    '6966 large rewrite
'    '6966 but still true even if errors on first try but successful on resend.
'    Dim blRet As Boolean
'    Dim c As Integer
'    Dim slRet As String
'    Dim slRoutine As String
'    Dim slStatus As String
'
'    slRoutine = "SetAiringNetworks"
'    slStatus = ""
'    blRet = True
'    If Not mSendBasic(False, False, slRoutine, slStatus) Then
'        If bmIsError Then
'            mSetResults "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted.", MESSAGERED
'            myExport.WriteWarning "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted."
'            For c = 1 To imMaxRetries - 1
'                If mSendBasic(False, True, slRoutine, slStatus) Then
'                    Exit For
'                ElseIf bmIsError = False Then
'                    blRet = False
'                    Exit For
'                End If
'            Next c
'            If bmIsError Then
'                blRet = mSendBasic(True, True, slRoutine, slStatus)
'            End If
'            'resending fixed the issue
'            If bmIsError = False Then
'                bmAlertAboutReExport = True
'                mSetResults "Error in sending " & slRoutine & " corrected. Export Ok and continuing.", MESSAGERED
'                myExport.WriteWarning "Error in sending " & slRoutine & " corrected. Export Ok and continuing."
'                If blRet = False Then
'                    slDoNotReturn = mReturnIds(slStatus)
''                    'I log message even if slDoNotReturn has nothing in it...just in case there was an error in mReturnIds
''                    myExport.WriteWarning "Some Programs not accepted by XDS: " & slStatus
'                End If
'            End If
'        Else
'            blRet = False
'            slDoNotReturn = mReturnIds(slStatus)
'            'I log message even if slDoNotReturn has nothing in it...just in case there was an error in mReturnIds
'            myExport.WriteWarning "Some Programs not accepted by XDS: " & slStatus
'        End If
'    Else
'
'    End If
'    mVehicleSendAndTest = blRet
'    Exit Function
'End Function
'ttp 5589
Private Function mAgreementSend(slXMLINIInputFile As String, ByVal slSection As String) As Boolean
    ' Cue means non-cumulus  ISCI means Cumulus
    '6581 slSection added
    Dim rstCount As ADODB.Recordset
    Dim blRet As Boolean
    Dim blSendToWeb As Boolean
    Dim blNotCumulus As Boolean
    Dim blIsError As Boolean
    Dim slRet As String
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim slYesterday As String
    '6901
    Dim slAllPorts As String
    
    Dim slProp As String
    
On Error GoTo ErrHandler
    blRet = True
    '6945 set this here instead of later
    If udcCriteria.XGenType(0, slProp) Then
        blSendToWeb = True
    Else
        blSendToWeb = False
    End If
    'note there is no message because this could be set in ini as this, or this dual provider doesn't send.
    If bmSendAgreementIds = False And bmSendStationIds = False Then
        'Dan M going to cleanup gives message that authorizations were done.  Let's skip that.
       ' GoTo cleanup
        mAgreementSend = blRet
        Exit Function
    End If
    SQLQuery = "SELECT count(*) as amount FROM Site Where siteCode = 1 AND siteAgreementToXDS = 'Y'"
    Set rstCount = gSQLSelectCall(SQLQuery)
    If rstCount!amount > 0 Then
        blNotCumulus = Not mCumulusHeadEnd(slSection, slXMLINIInputFile, blIsError)
        If blIsError Then
            blRet = False
            mSetResults "Problem sending authorizations: Can't read WebServiceURL in xml.ini.", MESSAGERED
            'gLogMsg "Problem Sending Authorizations: Can't read WebServiceURL in xml.ini at " & slXMLINIInputFile, smPathForgLogMsg, False
            myExport.WriteError "Problem Sending Authorizations: Can't read WebServiceURL in xml.ini at " & slXMLINIInputFile, False, False
            GoTo Cleanup
        End If
        '6901 need for cue and now isci games
        gLoadFromIni slSection, "AllPorts", slXMLINIInputFile, slAllPorts
        If slAllPorts = "0" Then
            bmIsAllPorts = False
        Else
            bmIsAllPorts = True
        End If
        'Cumulus needs the 2nd (or 3rd) xml.ini section ..sending to Vantive
        If blNotCumulus = False Then
            gLoadFromIni CUMULUSVANTIVESECTION, "WebServiceURL", slXMLINIInputFile, slRet
            'section doesn't exist
            If slRet = "Not Found" Then
                blRet = False
                mSetResults "Problem sending authorizations: Can't send authorizations to Cumulus head-end without '" & CUMULUSVANTIVESECTION & "' section in xml.ini.", MESSAGERED
               ' gLogMsg "Problem sending authorizations: Can't send authorizations to Cumulus head-end without '" & CUMULUSVANTIVESECTION & "' section in xml.ini.", smPathForgLogMsg, False
                myExport.WriteError "Problem sending authorizations: Can't send authorizations to Cumulus head-end without '" & CUMULUSVANTIVESECTION & "' section in xml.ini.", False, False
                GoTo Cleanup
            Else
                'going to send to 'vantive'
                slSection = CUMULUSVANTIVESECTION
            End If
        End If
        'new for 6581..better way to handle xml errors.  This was created in mExport
        If Len(smXmlErrorFile) = 0 Then
            blRet = False
            mSetResults "Problem sending authorizations: error file doesn't exist.", MESSAGERED
            'gLogMsg "Problem Sending Authorizations: error file doesn't exist. Make sure the xml.ini 'logfile' has a valid directory.", smPathForgLogMsg, False
            myExport.WriteError "Problem Sending Authorizations: error file doesn't exist. Make sure the xml.ini 'logfile' has a valid directory.", False, False
            GoTo Cleanup
        End If
        '6945 already set
        If blSendToWeb Then
       ' If udcCriteria.XGenType(0) Then
           ' blSendToWeb = True
            'gLogMsg "!! Exporting Authorizations  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", smPathForgLogMsg, False
            myExport.WriteFacts "!! Exporting Authorizations  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
            mSetResults "Exporting Authorization Information", MESSAGEBLACK
        Else
           ' blSendToWeb = False
             'gLogMsg "!! Writing Authorizations in test mode  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", smPathForgLogMsg, False
            myExport.WriteFacts "!! Writing Agreements in test mode  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
            mSetResults "Writing Agreement Information to " & AGREEMENTLOG & smDateForLogs & ".txt", MESSAGEBLACK
        End If
        If igExportSource = 2 Then DoEvents
        If blNotCumulus Then
            If Not mAgreementCueFacts(slXMLINIInputFile, blSendToWeb, slSection) Then
                blRet = False
            End If
        Else
            slEarliestDate = Format$(smDate, sgSQLDateForm)
            slLatestDate = DateAdd("d", imNumberDays - 1, smDate)
            slLatestDate = Format$(slLatestDate, sgSQLDateForm)
            slYesterday = gNow()
            slYesterday = DateAdd("d", -1, slYesterday)
            slYesterday = Format$(slYesterday, sgSQLDateForm)
            If Not mAgreementISCIFactsActive(slXMLINIInputFile, blSendToWeb, slSection, slYesterday) Then
                blRet = False
            End If
            'those marked as cancelled and already sent once
            If Not mAgreementISCIFactsCancelled(slXMLINIInputFile, blSendToWeb, slSection, slYesterday, False) Then
                blRet = False
            End If
            'those marked as cancelled but not yet sent.
            If Not mAgreementISCIFactsCancelled(slXMLINIInputFile, blSendToWeb, slSection, slYesterday, True) Then
                blRet = False
            End If
            '6835 games sent separately and as backoffice
            slSection = CUMULUSBACKOFFICESECTION
            If Not mAgreementISCIGames(slXMLINIInputFile, blSendToWeb, slSection, slYesterday) Then
                blRet = False
            End If
            '6835 don't have to cancel isci games
'           If Not mAgreementISCIGamesCancelled(slXMLINIInputFile, blSendToWeb, slSection, slYesterday) Then
'                blRet = False
'            End If
        End If
    ' agreements not selected in options.  Get out without message
    Else
        mAgreementSend = True
        Exit Function
    End If
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information Completed.", MESSAGEBLACK
            'gLogMsg "Authorization Information Completed.", smPathForgLogMsg, False
            myExport.WriteFacts "Authorization Information Completed.", True
        End If
    Else
        mSetResults "Authorization Information Completed. Test send. ", MESSAGEBLACK
        'gLogMsg "Authorization Information Completed. Test send.", smPathForgLogMsg, False
        myExport.WriteFacts "Authorization Information Completed. Test send.", True
    End If
    If Not rstCount Is Nothing Then
        If (rstCount.State And adStateOpen) <> 0 Then
            rstCount.Close
        End If
        Set rstCount = Nothing
    End If
    mAgreementSend = blRet
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementSend"
    blRet = False
    GoTo Cleanup
End Function

Private Function mAgreementCueFacts(slXMLINIInputFile As String, blSendToWeb As Boolean, slSection As String) As Boolean
    Dim blRet As Boolean
    Dim rstAgreement As ADODB.Recordset
    Dim slNeedUpdate As String
    Dim ilRet As Integer
    Dim slProgCode As String
    Dim blDontUpdate As Boolean
    Dim slAttEndDate As String
    Dim slAllowedVehicles As String
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim slVef As String
    '5991 6901 made modular-level
'    Dim slAllPorts As String
'    Dim blIsAllPorts As Boolean
    '6581
    Dim llCount As Long
    
    llCount = 0
On Error GoTo ErrHand
    blRet = True
    '6901 make module wide
    '5991
'    gLoadFromIni slSection, "AllPorts", slXMLINIInputFile, slAllPorts
'    If slAllPorts = "0" Then
'        blIsAllPorts = False
'    Else
'        blIsAllPorts = True
'    End If


'    For ilIndex = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilIndex) Then
    '8721 index of ilVef causes issue later in code.  Change back to ilIndex at top and where index is used (TextMatrix)
'    For ilVef = 1 To grdVeh.Rows - 1 Step 1
    For ilIndex = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilIndex, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilIndex, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilIndex, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilIndex, VEHINDEX)
                If igExportSource = 2 Then DoEvents
                ReDim tmAgreements(0)
                slNeedUpdate = ""
                'ilVef = lbcVehicles.ItemData(ilIndex)
                ilVef = imVefCode
                'Dan M 4/18/14 moved here from loop below
                slVef = mAgreementsGetVehicleName(ilVef)
                lacProcessing.Caption = "Collecting Authorization Information for " & slVef
                '6199 don't set to 'y' first, so agreements that have been changed go out.
                '6159 game agreements need special processing.  Always set to 'M' if in selected date range
                If Not mAgreementAdjustGamesCue(ilVef) Then
                    mAgreementCueFacts = False
                    'gLogMsg "Problem writing Authorization Information.", smPathForgLogMsg, False
                    myExport.WriteError "Problem writing Authorization Information.", True, False
                    Exit Function
                End If
                SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId, vffXDProgCodeID" & _
                " FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join shtt on attshfcode = shttcode Left outer join VAT_Vendor_Agreement on attcode = vatAttCode" & _
                " WHERE  attSentToXDsStatus in ('M', 'N') AND vffxdxmlForm in ('A','S')AND SHTTusedforXDigital = 'Y' AND vatWvtVendorId = " & Vendors.XDS_Break & _
                "  AND  attvefcode = " & ilVef
                If bmSendStationIds And bmSendAgreementIds Then
                    SQLQuery = SQLQuery & " AND (shttStationId > 0 Or attXDReceiverId > 0) "
                ElseIf bmSendStationIds Then
                    SQLQuery = SQLQuery & " AND (shttStationId > 0 ) "
                Else
                    SQLQuery = SQLQuery & " AND ( attXDReceiverId > 0) "
                End If
                Set rstAgreement = gSQLSelectCall(SQLQuery)
                Do While Not rstAgreement.EOF
    '                slVef = mAgreementsGetVehicleName(ilVef)
    '                lacProcessing.Caption = "Collecting Authorization Information for " & slVef
                    slProgCode = gXMLNameFilter(rstAgreement!vffxdprogcodeid)
                    slAttEndDate = mEarlierDate(rstAgreement!attOffAir, rstAgreement!attDropDate)
                    If (mEarlierDate(rstAgreement!attOnAir, slAttEndDate) = rstAgreement!attOnAir And Len(slProgCode) > 0) Then
                        With tmAgreementInfo
                            'this is to pass error info
                            .sStation = rstAgreement!shttCallLetters
                            .sCode = rstAgreement!attCode
                            .sStartDate = gAdjYear(Format$(rstAgreement!attOnAir, "m/d/yy"))
                            .sEndDate = gAdjYear(Format$(slAttEndDate, "m/d/yy"))
                            'Dan I already filtered above
                            '.sProgramCode = gXMLNameFilter(slProgCode)
                            .sProgramCode = slProgCode
                            .sProgramName = gXMLNameFilter(slVef)
                            If rstAgreement!attXDReceiverId > 0 Then
                                .sSiteId = rstAgreement!attXDReceiverId
                            Else
                                .sSiteId = rstAgreement!shttStationId
                            End If
                        End With
                        ' if these are games marked as 'events', then we need to send multiple ProgramCodes but NOT the slProgCode inserted earlier
                        If UCase(slProgCode) = "EVENT" Then
                           'skip updating those that writing came back with a failure
                            If Not mAgreementGameEvent(ilVef, blDontUpdate, slVef, blSendToWeb) Then
                                If mAddToArray() Then
                                    slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                                    If blSendToWeb Then
                                        myExport.WriteFacts "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
                                    Else
                                        myExport.WriteFacts "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
                                    End If
                                End If
                            ElseIf Not blDontUpdate Then
                                slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                            End If
                        ElseIf mAddToArray() Then
                             slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                            If blSendToWeb Then
                                'gLogMsg "Sending agreement " & slVef & " " & Trim(tmAgreementInfo.sStation), smPathForgLogMsg, False
                                myExport.WriteFacts "Sending agreement " & slVef & " " & Trim(tmAgreementInfo.sStation)
                            Else
                                'gLogMsg "Test send agreement " & slVef & " " & Trim(tmAgreementInfo.sStation), smPathForgLogMsg, False
                                myExport.WriteFacts "Test send agreement " & slVef & " " & Trim(tmAgreementInfo.sStation)
                            End If
                        End If
                    End If
                    rstAgreement.MoveNext
                Loop
                If Len(slNeedUpdate) > 0 Then
                    If Not mAgreementSendFacts(slNeedUpdate, blSendToWeb, True, "Y", slXMLINIInputFile, slSection, llCount) Then
                        blRet = False
                        If bmIsError Then
                            GoTo Cleanup
                        End If
                    End If
                End If
            End If
        End If
    Next
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information Cue Model. Sent: " & llCount, MESSAGEBLACK
           ' gLogMsg "Authorization Information Cue Model. Sent: " & llCount, smPathForgLogMsg, False
           myExport.WriteFacts "Authorization Information Cue Model. Sent: " & llCount, True
        End If
    Else
        mSetResults "Authorization Information Cue Model. Test send: " & llCount, MESSAGEBLACK
        'gLogMsg "Authorization Information Cue Model. Test send: " & llCount, smPathForgLogMsg, False
        myExport.WriteFacts "Authorization Information Cue Model. Test send: " & llCount, True
    End If
    mAgreementCueFacts = blRet
    Erase tmAgreements
    If Not rstAgreement Is Nothing Then
        If (rstAgreement.State And adStateOpen) <> 0 Then
            rstAgreement.Close
        End If
        Set rstAgreement = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementCueFacts"
    GoTo Cleanup
End Function
Private Function mAgreementISCIFactsActive(slXMLINIInputFile As String, blSendToWeb As Boolean, ByVal slSection As String, slYesterday As String) As Boolean
    Dim blRet As Boolean
    Dim rstAgreement As ADODB.Recordset
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim slNeedUpdate As String
    Dim ilRet As Integer
    Dim slProgCode As String
    Dim blDontUpdate As Boolean
    Dim slAttEndDate As String
    Dim slAllowedVehicles As String
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim slVef As String
    '6581
    Dim llCount As Long
    Dim ilVpf As Integer
    Dim ilID As Integer
    Dim ilVefIndex As Integer
    
    llCount = 0
On Error GoTo ErrHand
    blRet = True
'    For ilIndex = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilIndex) Then
        
    For ilIndex = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilIndex, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilIndex, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilIndex, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilIndex, VEHINDEX)
                If igExportSource = 2 Then DoEvents
                ReDim tmAgreements(0)
                slNeedUpdate = ""
                'ilVef = lbcVehicles.ItemData(ilIndex)
                ilVef = imVefCode
                ilVefIndex = gBinarySearchVef(CLng(ilVef))
                If ilVefIndex <> -1 Then
                    slVef = Trim$(tgVehicleInfo(ilVefIndex).sVehicleName)
                    ilVpf = gBinarySearchVpf(CLng(ilVef))
                    If ilVpf <> -1 Then
                        ilID = tgVpfOptions(ilVpf).iInterfaceID
                        '6835 don't send games
                       If ilID > 0 And tgVehicleInfo(ilVefIndex).sVehType <> "G" Then
                            lacProcessing.Caption = "Collecting Authorization Information for " & slVef
                            SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                            " FROM att inner join  shtt on attshfcode = shttcode Left outer join VAT_Vendor_Agreement on attcode = vatAttCode" & _
                            " WHERE  attSentToXDsStatus in ('M', 'N')  AND SHTTusedforXDigital = 'Y'  AND vatWvtVendorId = " & Vendors.XDS_ISCI & " And attvefCode = " & ilVef
    
                            If bmSendStationIds And bmSendAgreementIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 Or attXDReceiverId > 0) "
                            ElseIf bmSendStationIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 ) "
                            Else
                                SQLQuery = SQLQuery & " AND ( attXDReceiverId > 0) "
                            End If
                            Set rstAgreement = gSQLSelectCall(SQLQuery)
                            Do While Not rstAgreement.EOF
                                slAttEndDate = mEarlierDate(rstAgreement!attOffAir, rstAgreement!attDropDate)
                                'start earlier than end...not 'cancelled before aired'
                                If (mEarlierDate(rstAgreement!attOnAir, slAttEndDate) = rstAgreement!attOnAir) Then
                                    'active agreement?  end after yesterday?
                                    If mEarlierDate(slAttEndDate, slYesterday) = slYesterday Then
                                        With tmAgreementInfo
                                            'this is to pass error info
                                            .sStation = rstAgreement!shttCallLetters
                                            .sCode = rstAgreement!attCode
                                            .sStartDate = gAdjYear(Format$(rstAgreement!attOnAir, "m/d/yy"))
                                            .sEndDate = gAdjYear(Format$(slAttEndDate, "m/d/yy"))
                                            .sNetworkId = ilID
                                            .sStatus = "active"
                                            '6796
                                            .sProgramCode = ""
                                            If bmSendStationIds And bmSendAgreementIds Then
                                                If rstAgreement!attXDReceiverId > 0 Then
                                                    .sSiteId = rstAgreement!attXDReceiverId
                                                Else
                                                    .sSiteId = rstAgreement!shttStationId
                                                End If
                                            ElseIf bmSendStationIds Then
                                                 .sSiteId = rstAgreement!shttStationId
                                            Else
                                                .sSiteId = rstAgreement!attXDReceiverId
                                            End If
    '                                        If rstAgreement!attXDReceiverID > 0 Then
    '                                            .sSiteId = rstAgreement!attXDReceiverID
    '                                        Else
    '                                            .sSiteId = rstAgreement!shttStationId
    '                                        End If
                                        End With
                                         '6835 removed games
    '                                    ' if these are games, then we need to send multiple ProgramCodes but NOT the slProgCode inserted earlier
    '                                    If tgVehicleInfo(ilVefIndex).sVehType = "G" Then
    '                                        'skip updating those that writing came back with a failure
    '                                         If Not mAgreementGameEvent(ilVef, blDontUpdate, slVef, blSendToWeb) Then
    '                                             If mAddToArray() Then
    '                                                 slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
    '                                                 If blSendToWeb Then
    '                                                    myExport.WriteFacts "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
    '                                                Else
    '                                                    myExport.WriteFacts "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
    '                                                End If
    '                                            End If
    '                                         ElseIf Not blDontUpdate Then
    '                                             slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
    '                                         End If
    '                                    ElseIf mAddToArray() Then
                                        If mAddToArray() Then
                                             slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                                             If blSendToWeb Then
                                                myExport.WriteFacts "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
                                            Else
                                                myExport.WriteFacts "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef
                                            End If
                                       End If
                                    End If 'active
                                End If 'not cancelled before aired
                                rstAgreement.MoveNext
                            Loop
                        'vehicle id = 0 and not a game
                        ElseIf tgVehicleInfo(ilVefIndex).sVehType <> "G" Then
                            mSetResults slVef & " missing vehicle ID in agreements.", MESSAGEBLACK
                            myExport.WriteWarning slVef & " missing vehicle ID in agreements"
                        End If
                    Else
                        mSetResults slVef & " missing vehicle ID in agreements", MESSAGEBLACK
                        myExport.WriteWarning slVef & " missing vehicle ID in agreements(no vpf options)"
                    End If 'ilvpf <> -1
                Else
                        mSetResults ilVef & " missing in vehicle array", MESSAGERED
                        myExport.WriteError ilVef & " missing in vehicle array-mAgreeementISCIFactsActive", True, False
                End If ' ilVefIndex <> -1
            '8694 adding vehicle grid..put end if in wrong spot
           ' End If
            '6725 cumulus and cancellations: made into own function
            If Len(slNeedUpdate) > 0 Then
                If Not mAgreementSendFacts(slNeedUpdate, blSendToWeb, False, "Y", slXMLINIInputFile, slSection, llCount) Then
                    blRet = False
                    If bmIsError Then
                        GoTo Cleanup
                    End If
                End If
            End If
            '8694  moved vehicle grid end if here
            End If
        End If 'selected
    Next
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information ISCI Model. Sent: " & llCount, MESSAGEBLACK
            myExport.WriteFacts "Authorization Information ISCI Model. Sent: " & llCount, True
        End If
    Else
        mSetResults "Authorization Information ISCI Model. Test send: " & llCount, MESSAGEBLACK
        myExport.WriteFacts "Authorization Information ISCI Model. Test send: " & llCount, True
    End If
    mAgreementISCIFactsActive = blRet
    Erase tmAgreements
    If Not rstAgreement Is Nothing Then
        If (rstAgreement.State And adStateOpen) <> 0 Then
            rstAgreement.Close
        End If
        Set rstAgreement = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementISCIFactsActive"
    GoTo Cleanup
End Function
Private Function mAgreementISCIGames(slXMLINIInputFile As String, blSendToWeb As Boolean, ByVal slSection As String, slYesterday As String) As Boolean
    Dim blRet As Boolean
    Dim rstAgreement As ADODB.Recordset
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim slNeedUpdate As String
    Dim ilRet As Integer
    Dim slProgCode As String
    Dim blDontUpdate As Boolean
    Dim slAttEndDate As String
    Dim slAllowedVehicles As String
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim slVef As String
    '6581
    Dim llCount As Long
    Dim ilVpf As Integer
    Dim ilID As Integer
    Dim ilVefIndex As Integer
    '6583
    Dim slStart As String
    Dim slEnd As String
    
    llCount = 0
On Error GoTo ErrHand
    blRet = True
    slStart = "'" & Format$(smDate, sgSQLDateForm) & "'"
    slEnd = "'" & Format$(DateAdd("d", imNumberDays - 1, smDate), sgSQLDateForm) & "'"
'    For ilIndex = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilIndex) Then
'            If igExportSource = 2 Then DoEvents
    '8721 index of ilVef causes issue later in code.  Change back to ilIndex at top and where index is used (TextMatrix)
'    For ilVef = 1 To grdVeh.Rows - 1 Step 1
    For ilIndex = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilIndex, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilIndex, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilIndex, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilIndex, VEHINDEX)
                ReDim tmAgreements(0)
                slNeedUpdate = ""
                'ilVef = lbcVehicles.ItemData(ilIndex)
                ilVef = imVefCode
                ilVefIndex = gBinarySearchVef(CLng(ilVef))
                If ilVefIndex <> -1 Then
                    slVef = Trim$(tgVehicleInfo(ilVefIndex).sVehicleName)
                    ilVpf = gBinarySearchVpf(CLng(ilVef))
                    If ilVpf <> -1 Then
                        ilID = tgVpfOptions(ilVpf).iInterfaceID
                        '6835 games only
                        If ilID > 0 And tgVehicleInfo(ilVefIndex).sVehType = "G" Then
                            lacProcessing.Caption = "Collecting Event Authorization Information for " & slVef
                            '8/22/14 Dan fix don't use 'P' anymore in v70  7375 add attAudioDelivery
                           ' SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                           ' " FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join shtt on attshfcode = shttcode" & _
                           ' " WHERE vffxdxmlForm in ('P')AND SHTTusedforXDigital = 'Y'  And attvefCode = " & ilVef
    '                        SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
    '                        " FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join shtt on attshfcode = shttcode" & _
    '                        " WHERE  SHTTusedforXDigital = 'Y' AND attAudioDelivery = 'X'  And attvefCode = " & ilVef
                            '7701
                            SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                            " FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join shtt on attshfcode = shttcode Left outer join VAT_Vendor_Agreement on attcode = vatAttCode" & _
                            " WHERE  SHTTusedforXDigital = 'Y' AND vatWvtVendorId = " & Vendors.XDS_ISCI & " And attvefCode = " & ilVef
                            '6835
                            SQLQuery = SQLQuery & " AND attOnAir <= " & slEnd & " AND attOffAir >= " & slStart & " and attDropDate >=" & slStart
                            If bmSendStationIds And bmSendAgreementIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 Or attXDReceiverId > 0) "
                            ElseIf bmSendStationIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 ) "
                            Else
                                SQLQuery = SQLQuery & " AND ( attXDReceiverId > 0) "
                            End If
                            Set rstAgreement = gSQLSelectCall(SQLQuery)
                            Do While Not rstAgreement.EOF
                                slAttEndDate = mEarlierDate(rstAgreement!attOffAir, rstAgreement!attDropDate)
                                'start earlier than end...not 'cancelled before aired'
                                If (mEarlierDate(rstAgreement!attOnAir, slAttEndDate) = rstAgreement!attOnAir) Then
                                    'active agreement?  end after yesterday?
                                    If mEarlierDate(slAttEndDate, slYesterday) = slYesterday Then
                                        With tmAgreementInfo
                                            'this is to pass error info
                                            .sStation = rstAgreement!shttCallLetters
                                            .sCode = rstAgreement!attCode
    '                                        .sStartDate = gAdjYear(Format$(rstAgreement!attOnAir, "m/d/yy"))
    '                                        .sEndDate = gAdjYear(Format$(slAttEndDate, "m/d/yy"))
    '                                        '6835 this will be 'ProgramNumber
    '                                        .sProgramCode = slProgCode
                                            If bmSendStationIds And bmSendAgreementIds Then
                                                If rstAgreement!attXDReceiverId > 0 Then
                                                    .sSiteId = rstAgreement!attXDReceiverId
                                                Else
                                                    .sSiteId = rstAgreement!shttStationId
                                                End If
                                            ElseIf bmSendStationIds Then
                                                 .sSiteId = rstAgreement!shttStationId
                                            Else
                                                .sSiteId = rstAgreement!attXDReceiverId
                                            End If
                                        End With
                                        ' if these are games, then we need to send multiple ProgramNumbers(programCode) but NOT the slProgCode inserted earlier
                                        'skip updating those that writing came back with a failure
                                        If mAgreementGameEvent(ilVef, blDontUpdate, slVef, blSendToWeb) Then
                                            If Not blDontUpdate Then
                                                slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                                             End If
                                        End If
                                    End If 'active
                                End If 'not cancelled before aired
                                rstAgreement.MoveNext
                            Loop
                        ' only games with no id
                        ElseIf tgVehicleInfo(ilVefIndex).sVehType = "G" Then
                            mSetResults slVef & " missing vehicle ID.", MESSAGEBLACK
                            myExport.WriteWarning slVef & " missing vehicle ID"
                        End If
                    Else
                        mSetResults slVef & " missing vehicle ID in agreements", MESSAGEBLACK
                        myExport.WriteWarning slVef & " missing vehicle ID in agreements(no vpf options)"
                    End If 'ilvpf <> -1
                Else
                        mSetResults ilVef & " missing in vehicle array", MESSAGERED
                        myExport.WriteError ilVef & " missing in vehicle array-mAgreeementISCIFactsActive", True, False
                End If ' ilVefIndex <> -1
                '6725 cumulus and cancellations: made into own function
                If Len(slNeedUpdate) > 0 Then
                    'send as 'cumulus' and with the new section of backoffice.  That's how the function knows how to send
                    '6901 use bmIsAllPorts rather than 'True'
                    If Not mAgreementSendFacts(slNeedUpdate, blSendToWeb, False, "Y", slXMLINIInputFile, slSection, llCount) Then
                        blRet = False
                        If bmIsError Then
                            GoTo Cleanup
                        End If
                    End If
                End If
            End If 'selected
        End If
    Next
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information ISCI Events. Sent: " & llCount, MESSAGEBLACK
            myExport.WriteFacts "Authorization Information ISCI Events. Sent: " & llCount, True
        End If
    Else
        mSetResults "Authorization Information ISCI Events. Test send: " & llCount, MESSAGEBLACK
        myExport.WriteFacts "Authorization Information ISCI Events. Test send: " & llCount, True
    End If
    mAgreementISCIGames = blRet
    Erase tmAgreements
    If Not rstAgreement Is Nothing Then
        If (rstAgreement.State And adStateOpen) <> 0 Then
            rstAgreement.Close
        End If
        Set rstAgreement = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementISCIGames"
    GoTo Cleanup
End Function

Private Function mAgreementISCIFactsCancelled(slXMLINIInputFile As String, blSendToWeb As Boolean, ByVal slSection As String, slYesterday As String, blFirstSend As Boolean) As Boolean
    Dim blRet As Boolean
    Dim rstAgreement As ADODB.Recordset
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim slNeedUpdate As String
    Dim ilRet As Integer
    Dim slProgCode As String
    Dim blDontUpdate As Boolean
    Dim slAttEndDate As String
    Dim slAllowedVehicles As String
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim slVef As String
    '6581
    Dim llCount As Long
    Dim ilVpf As Integer
    Dim ilID As Integer
    Dim slEarlyCancelled As String
    Dim slvalue As String
    Dim blContinue As Boolean
    Dim slCancelled As String
    
    Const DAYSASCANCELLED As Integer = -14
    
    llCount = 0
On Error GoTo ErrHand
    blRet = True
    If blFirstSend Then
        slEarlyCancelled = DateAdd("d", DAYSASCANCELLED, slYesterday)
        slEarlyCancelled = Format$(slEarlyCancelled, sgSQLDateForm)
        slvalue = "D"
        slCancelled = "-Cancelled"
    Else
        slvalue = "I"
        slCancelled = "-Cancelled resend"
    End If
'    For ilIndex = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilIndex) Then
    For ilIndex = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilIndex, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilIndex, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilIndex, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilIndex, VEHINDEX)
                If igExportSource = 2 Then DoEvents
                ReDim tmAgreements(0)
                slNeedUpdate = ""
                'ilVef = lbcVehicles.ItemData(ilIndex)
                ilVef = imVefCode
                slVef = mAgreementsGetVehicleName(ilVef)
                ilVpf = gBinarySearchVpf(CLng(ilVef))
                If ilVpf <> -1 Then
                    ilID = tgVpfOptions(ilVpf).iInterfaceID
                    If ilID > 0 Then
                    'new for v70   AND attAudioDelivery = 'B'
                        If blFirstSend Then
                            '7701
                            SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                            " FROM att inner join shtt on attshfcode = shttcode Left outer join VAT_Vendor_Agreement on attcode = vatAttCode" & _
                            " WHERE attSentToXDsStatus in ('M','N','Y') AND SHTTusedforXDigital = 'Y'  AND vatWvtVendorId = " & Vendors.XDS_ISCI & _
                            "  AND  attvefcode = " & ilVef & _
                            " AND ((attoffAir >= '" & slEarlyCancelled & "' and attoffair <= '" & slYesterday & "') OR (ATTDropDate >= '" & slEarlyCancelled & "' and attdropdate <= '" & slYesterday & "'))"
                            '7375 add attAudioDelivery
    '                        SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
    '                        " FROM att inner join shtt on attshfcode = shttcode " & _
    '                        " WHERE attSentToXDsStatus in ('M','N','Y') AND SHTTusedforXDigital = 'Y'  AND attAudioDelivery = 'X' " & _
    '                        "  AND  attvefcode = " & ilVef & _
    '                        " AND ((attoffAir >= '" & slEarlyCancelled & "' and attoffair <= '" & slYesterday & "') OR (ATTDropDate >= '" & slEarlyCancelled & "' and attdropdate <= '" & slYesterday & "'))"
    
    '                        'get all agreements, then do date (2 weeks), then do date again later in code--make sure there isn't an earlier end date and thus don't have to send
    '                        SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
    '                        " FROM att inner join shtt on attshfcode = shttcode " & _
    '                        " WHERE attSentToXDsStatus in ('M','N','Y') AND SHTTusedforXDigital = 'Y' " & _
    '                        "  AND  attvefcode = " & ilVef & _
    '                        " AND ((attoffAir >= '" & slEarlyCancelled & "' and attoffair <= '" & slYesterday & "') OR (ATTDropDate >= '" & slEarlyCancelled & "' and attdropdate <= '" & slYesterday & "'))"
                        Else
                            ''D' status means sent once
                            SQLQuery = " SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                            " FROM att  inner join shtt on attshfcode = shttcode WHERE  attSentToXDsStatus in ('D')" & _
                            " AND SHTTusedforXDigital = 'Y' " & _
                            "  AND  attvefcode = " & ilVef
                        End If
                        If bmSendStationIds And bmSendAgreementIds Then
                            SQLQuery = SQLQuery & " AND (shttStationId > 0 Or attXDReceiverId > 0) "
                        ElseIf bmSendStationIds Then
                            SQLQuery = SQLQuery & " AND (shttStationId > 0 ) "
                        Else
                            SQLQuery = SQLQuery & " AND ( attXDReceiverId > 0) "
                        End If
                        Set rstAgreement = gSQLSelectCall(SQLQuery)
                        Do While Not rstAgreement.EOF
                            blContinue = False
                            lacProcessing.Caption = "Collecting Authorization Information for " & slVef
                            If blFirstSend Then
                                slAttEndDate = mEarlierDate(rstAgreement!attOffAir, rstAgreement!attDropDate)
                                'start earlier than end...not 'cancelled before aired'
                                If (mEarlierDate(rstAgreement!attOnAir, slAttEndDate) = rstAgreement!attOnAir) Then
                                    'cancelled agreement?  end between yesterday and 2 weeks ago?  Note that because earlyCancelled is sqlformat, on 2nd test, Att must be >.  If =, will return slAttEndDate, which will not match slEarlyCancelled
                                    If mEarlierDate(slAttEndDate, slYesterday) = slAttEndDate And mEarlierDate(slAttEndDate, slEarlyCancelled) = slEarlyCancelled Then
                                        blContinue = True
                                    End If
                                End If
                            Else
                                blContinue = True
                            End If
                            ' 2nd time doesn't need to test dates
                            If blContinue Then
                                With tmAgreementInfo
                                    'this is to pass error info
                                    .sStation = rstAgreement!shttCallLetters
                                    .sCode = rstAgreement!attCode
                                    .sStartDate = gAdjYear(Format$(rstAgreement!attOnAir, "m/d/yy"))
                                    'to make inactive, 'set enddate to today'.  cancelled is always 'immediate'
                                    .sEndDate = gAdjYear(Format$(Now(), "m/d/yy"))
                                    .sNetworkId = ilID
                                    .sStatus = "inactive"
                                    If bmSendStationIds And bmSendAgreementIds Then
                                        If rstAgreement!attXDReceiverId > 0 Then
                                            .sSiteId = rstAgreement!attXDReceiverId
                                        Else
                                            .sSiteId = rstAgreement!shttStationId
                                        End If
                                    ElseIf bmSendStationIds Then
                                         .sSiteId = rstAgreement!shttStationId
                                    Else
                                        .sSiteId = rstAgreement!attXDReceiverId
                                    End If
    '                                        If rstAgreement!attXDReceiverID > 0 Then
    '                                            .sSiteId = rstAgreement!attXDReceiverID
    '                                        Else
    '                                            .sSiteId = rstAgreement!shttStationId
    '                                        End If
                                End With
                                If mAddToArray() Then
                                     slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                                    If blSendToWeb Then
                                        myExport.WriteFacts "Sending cancelled agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef, True
                                    Else
                                        myExport.WriteFacts "Test send cancelled agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVef, True
                                    End If
                                End If
                            End If 'ok to send
                            rstAgreement.MoveNext
                        Loop
                        '6725 cumulus and cancellations: made into own function
                        If Len(slNeedUpdate) > 0 Then
                            If Not mAgreementSendFacts(slNeedUpdate, blSendToWeb, False, slvalue, slXMLINIInputFile, slSection, llCount) Then
                                blRet = False
                                If bmIsError Then
                                    GoTo Cleanup
                                End If
                            End If
                        End If
                    Else
                    'no message on cancelled
    '                    mSetResults slVef & " missing vehicle ID.", MESSAGEBLACK
    '                    myExport.WriteWarning slVef & " missing vehicle ID."
                    End If
                End If
            Else
'                mSetResults slVef & " missing vehicle ID in agreements", MESSAGEBLACK
'               myExport.WriteFacts slVef & " missing vehicle ID in agreements(no vpf options)"
            End If 'ilvpf <> -1
        End If 'selected
    Next
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information ISCI Model " & slCancelled & ". Sent: " & llCount, MESSAGEBLACK
            myExport.WriteFacts "Authorization Information ISCI Model " & slCancelled & ". Sent: " & llCount, True
        End If
    Else
        mSetResults "Authorization Information ISCI Model " & slCancelled & ". Test send: " & llCount, MESSAGEBLACK
        myExport.WriteFacts "Authorization Information ISCI Model " & slCancelled & ". Test send: " & llCount, True
    End If
    mAgreementISCIFactsCancelled = blRet
    Erase tmAgreements
    If Not rstAgreement Is Nothing Then
        If (rstAgreement.State And adStateOpen) <> 0 Then
            rstAgreement.Close
        End If
        Set rstAgreement = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementISCIFactsCancelled"
    GoTo Cleanup
End Function
Private Function mAgreementISCIGamesCancelled(slXMLINIInputFile As String, blSendToWeb As Boolean, ByVal slSection As String, slYesterday As String) As Boolean
    Dim blRet As Boolean
    Dim rstAgreement As ADODB.Recordset
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim slNeedUpdate As String
    Dim ilRet As Integer
    Dim slProgCode As String
    Dim blDontUpdate As Boolean
    Dim slAttEndDate As String
    Dim slAllowedVehicles As String
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim slVef As String
    '6581
    Dim llCount As Long
    Dim ilVpf As Integer
    Dim ilID As Integer
    Dim slEarlyCancelled As String
    Dim slvalue As String
    Dim slCancelled As String
    Dim ilVefIndex As Integer
    
    Const DAYSASCANCELLED As Integer = -14
    
    llCount = 0
On Error GoTo ErrHand
    blRet = True
    slEarlyCancelled = DateAdd("d", DAYSASCANCELLED, slYesterday)
    slEarlyCancelled = Format$(slEarlyCancelled, sgSQLDateForm)
    slvalue = "D"
    slCancelled = "-CancelledGames"
    'For ilIndex = 0 To lbcVehicles.ListCount - 1
    '    If lbcVehicles.Selected(ilIndex) Then
        
    For ilIndex = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilIndex, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilIndex, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilIndex, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilIndex, VEHINDEX)
                If igExportSource = 2 Then DoEvents
                ReDim tmAgreements(0)
                slNeedUpdate = ""
                'ilVef = lbcVehicles.ItemData(ilIndex)
                ilVef = imVefCode
                ilVpf = gBinarySearchVpf(CLng(ilVef))
                ilVefIndex = gBinarySearchVef(CLng(ilVef))
                If ilVpf <> -1 And ilVefIndex <> -1 Then
                    slVef = Trim(tgVehicleInfo(ilVefIndex).sVehicle)
                    ilID = tgVpfOptions(ilVpf).iInterfaceID
                    If ilID > 0 And tgVehicleInfo(ilVefIndex).sVehType = "G" Then
                        'cancelled games that have to go?
                        SQLQuery = "SELECT count(*) as amount FROM gsf_Game_Schd WHERE gsfVefCode = " & ilVef & " AND gsfAirDate >= '" & slEarlyCancelled & "' AND  gsfAirDate <= '" & slYesterday & "' AND Ltrim(gsfxdsprogcodeid) <> '' "
                        Set rstAgreement = gSQLSelectCall(SQLQuery)
                        If rstAgreement!amount > 0 Then
                            'get agreements marked as sent
                            'removed this: attSentToXDsStatus in ('Y') AND  for 1-didn't show in 'test' mode.
                            'different in v70
                            SQLQuery = "SELECT attCode,  attOnAir, attOffAir, attDropDate,attXDReceiverId, shttCallLetters, shttStationId" & _
                            " FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join shtt on attshfcode = shttcode" & _
                            " WHERE   SHTTusedforXDigital = 'Y' " & _
                            "  AND  attvefcode = " & ilVef
                            If bmSendStationIds And bmSendAgreementIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 Or attXDReceiverId > 0) "
                            ElseIf bmSendStationIds Then
                                SQLQuery = SQLQuery & " AND (shttStationId > 0 ) "
                            Else
                                SQLQuery = SQLQuery & " AND ( attXDReceiverId > 0) "
                            End If
                            Set rstAgreement = gSQLSelectCall(SQLQuery)
                            'for each agreement, cancel the game.
                            Do While Not rstAgreement.EOF
                                lacProcessing.Caption = "Collecting Authorization Information for " & slVef
                                slAttEndDate = mEarlierDate(rstAgreement!attOffAir, rstAgreement!attDropDate)
                                'start earlier than end...not 'cancelled before aired'
                                If (mEarlierDate(rstAgreement!attOnAir, slAttEndDate) = rstAgreement!attOnAir) Then
                                    With tmAgreementInfo
                                        'this is to pass error info
                                        .sStation = rstAgreement!shttCallLetters
                                        .sCode = rstAgreement!attCode
                                        .sStartDate = gAdjYear(Format$(rstAgreement!attOnAir, "m/d/yy"))
                                        'to make inactive, 'set enddate to today'.  cancelled is always 'immediate'
                                        .sEndDate = gAdjYear(Format$(Now(), "m/d/yy"))
                                        .sNetworkId = ilID
                                        .sStatus = "inactive"
                                        If bmSendStationIds And bmSendAgreementIds Then
                                            If rstAgreement!attXDReceiverId > 0 Then
                                                .sSiteId = rstAgreement!attXDReceiverId
                                            Else
                                                .sSiteId = rstAgreement!shttStationId
                                            End If
                                        ElseIf bmSendStationIds Then
                                             .sSiteId = rstAgreement!shttStationId
                                        Else
                                            .sSiteId = rstAgreement!attXDReceiverId
                                        End If
                                    End With
                                    'we are only doing games
                                    If mAgreementGameEvent(ilVef, blDontUpdate, slVef, blSendToWeb, slEarlyCancelled, slYesterday) And Not blDontUpdate Then
                                        slNeedUpdate = slNeedUpdate & rstAgreement!attCode & ","
                                    End If
                                End If
                                rstAgreement.MoveNext
                            Loop
                        End If 'there are cancelled games for vehicle
                        '6725 cumulus and cancellations: made into own function
                        If Len(slNeedUpdate) > 0 Then
                            If Not mAgreementSendFacts(slNeedUpdate, blSendToWeb, False, slvalue, slXMLINIInputFile, slSection, llCount) Then
                                blRet = False
                                If bmIsError Then
                                    GoTo Cleanup
                                End If
                            End If
                        End If
                    End If
                Else
                'no message on cancelled
    '                mSetResults slVef & " missing vehicle ID in agreements", MESSAGEBLACK
    '               myExport.WriteFacts slVef & " missing vehicle ID in agreements(no vpf options)"
                End If 'ilvpf <> -1
            End If 'selected
        End If
    Next
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Authorization Information ISCI Model " & slCancelled & ". Sent: " & llCount, MESSAGEBLACK
            myExport.WriteFacts "Authorization Information ISCI Model " & slCancelled & ". Sent: " & llCount, True
        End If
    Else
        mSetResults "Authorization Information ISCI Model " & slCancelled & ". Test send: " & llCount, MESSAGEBLACK
        myExport.WriteFacts "Authorization Information ISCI Model " & slCancelled & ". Test send: " & llCount, True
    End If
    mAgreementISCIGamesCancelled = blRet
    Erase tmAgreements
    If Not rstAgreement Is Nothing Then
        If (rstAgreement.State And adStateOpen) <> 0 Then
            rstAgreement.Close
        End If
        Set rstAgreement = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementISCIGamesCancelled"
    GoTo Cleanup
End Function

Private Function mAgreementSendFacts(slNeedUpdate As String, blSendToWeb As Boolean, blNotCumulus As Boolean, slUpdateValue As String, slXMLINIInputFile As String, slSection As String, llCount As Long) As Boolean
    Dim blRet As Boolean
    Dim slDoNotUpdate As String
    Dim c As Integer
    Dim blTestUpdate As Boolean
    '6835
    Dim blIsCumulusGame As Boolean
    '5/7/15 Dan need to store sldonotupdate
    Dim slDoNotMaster As String
    
    blTestUpdate = bmTestForceUpdateXHT
    blRet = True
    slDoNotMaster = ""
    bmFailedToReadReturn = False
    '6835 these are cumulus games!
    blIsCumulusGame = False
    If blNotCumulus = False And slSection = CUMULUSBACKOFFICESECTION Then
        blNotCumulus = True
        blIsCumulusGame = True
    End If
On Error GoTo ErrHand
    If Not blSendToWeb Then
        'csiXMLStart slXMLINIInputFile, slSection, "F", smExportDirectory & "AuthorizationInformation" & ".txt", sgCRLF '-" & slVef & "
        csiXMLStart slXMLINIInputFile, slSection, "F", smExportDirectory & AGREEMENTLOG & smDateForLogs & ".txt", sgCRLF, smXmlErrorFile
    Else
        csiXMLStart slXMLINIInputFile, slSection, "T", "", "", smXmlErrorFile
    End If
    If blNotCumulus Then
        csiXMLSetMethod "SetAuthorizations", "Authorizations", "225", ""
    Else
        csiXMLSetMethod "SetClearances", "NewDataSet", "225", ""
    End If
    If UBound(tmAgreements) > 0 Then
        For c = 0 To UBound(tmAgreements) - 1
            If igExportSource = 2 Then DoEvents
            llCount = llCount + 1
            If llCount Mod imChunk = 0 Then
                '7508
               ' If Not mAgreementSendAndTest(slDoNotUpdate) Then
                If Not mSendAndTestReturn(Authorizations, slDoNotUpdate) Then
                    blRet = False
                    slDoNotMaster = slDoNotMaster & slDoNotUpdate
                    If bmIsError Then
                        csiXMLEnd
                        GoTo Cleanup
                    End If
                End If
                If blNotCumulus Then
                    csiXMLSetMethod "SetAuthorizations", "Authorizations", "225", ""
                Else
                    csiXMLSetMethod "SetClearances", "NewDataSet", "225", ""
                End If
            End If
            If blNotCumulus Then
                '6835
                If blIsCumulusGame Then
                    '6901 added allports
                    If mAgreementWriteCumulusGame(tmAgreements(c)) Then
                    Else
                        blRet = False
                        If bmIsError Then
                            csiXMLEnd
                            GoTo Cleanup
                        End If
                    End If
                    
                Else
                    If mAgreementWriteXml(tmAgreements(c)) Then
                    Else
                        blRet = False
                        If bmIsError Then
                            csiXMLEnd
                            GoTo Cleanup
                        End If
                    End If
                End If
            Else
                If Not mAgreementWriteCumulus(tmAgreements(c)) Then
                    blRet = False
                    If bmIsError Then
                        csiXMLEnd
                        GoTo Cleanup
                    End If
                End If
            End If  'cumulus?
        Next c
    Else
        blRet = False
        'gLogMsg "Problem writing Authorization Information.", smPathForgLogMsg, False
        myExport.WriteError "Problem writing Authorization Information.", True, False
    End If
     '7508
    ' If Not mAgreementSendAndTest(slDoNotUpdate) Then
     If Not mSendAndTestReturn(Authorizations, slDoNotUpdate) Then
        blRet = False
        slDoNotMaster = slDoNotMaster & slDoNotUpdate
        If bmIsError Then
            csiXMLEnd
            GoTo Cleanup
        End If
    End If
    csiXMLEnd
    If (blSendToWeb Or blTestUpdate) And Not bmIsWrongServicePage Then
        'handles if one of multiple didn't work.
        If bmFailedToReadReturn Then
            slDoNotMaster = ""
            bmFailedToReadReturn = False
        End If
        '5/7/15 Dan  warning, but no 'doNotUpdate'?  then none get updated!
        If blRet Or Len(slDoNotMaster) > 0 Then
            slNeedUpdate = mAdjustUpdates(slNeedUpdate, slDoNotMaster)
            If Len(slNeedUpdate) > 0 Then
                mAgreementUpdate slNeedUpdate, slUpdateValue
            End If
        ElseIf Not blRet Then
            myExport.WriteWarning "XDS did not accept some Agreements.  The return could not be parsed; No Agreements were marked as 'Sent'."
        End If
    End If
Cleanup:
    mAgreementSendFacts = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    mAgreementSendFacts = blRet
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementSendFacts"
    GoTo Cleanup
End Function
Private Function mAgreementsGetVehicleName(ilVefCode As Integer) As String
    Dim ilVefIndex As Integer
    Dim slRet As String
    
    slRet = ""
    ilVefIndex = gBinarySearchVef(CLng(ilVefCode))
    If ilVefIndex <> -1 Then
        slRet = Trim$(tgVehicleInfo(ilVefIndex).sVehicleName)
    End If
    mAgreementsGetVehicleName = slRet
End Function
'Private Function mAgreementsWriteToXml(blSendToWeb As Boolean, slXMLINIInputFile As String, slVef As String) As Boolean
'    Dim c As Integer
'    Dim blRet As Boolean
'    'Dim tlXmlStatus As CSIRspGetXMLStatus
'    Dim ilRet As Integer
'
'  On Error GoTo ErrHand
'    blRet = True
'    If UBound(tmAgreements) > 0 Then
''        If blSendToWeb Then
''            csiXMLStart slXMLINIInputFile, Section, "T", "", ""
''            Call mSetResults("Sending Authorization Information", RGB(0, 0, 0))
''        Else
''            csiXMLStart slXMLINIInputFile, Section, "F", smExportDirectory & "AuthorizationInformation.txt", sgCRLF '-" & slVef & "
'''            Call mSetResults("Writing Authorization Information to ", RGB(0, 155, 0))
'''            Call mSetResults(smExportDirectory & "AuthorizationInformation.txt", RGB(0, 155, 0))
''        End If
'        'one file? uncomment out
'        'csiXMLSetMethod "SetAuthorizations", "Authorizations", "225", ""
'        For c = 0 To UBound(tmAgreements) - 1
'            If Not mAgreementWriteXml(slXMLINIInputFile, blSendToWeb, tmAgreements(c)) Then
'                blRet = False
'            End If
'        Next c
'        'one file? uncomment out
''        ilRet = csiXMLWrite(1)
''        If ilRet <> True Then
''            blRet = False
''            csiXMLStatus tlXmlStatus
''            gLogMsg "ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''            mSetResults "Problem with sending an authorization; please see XDigitalExportLog.Txt", vbRed
''        End If
''        csiXMLEnd
'    End If
'    mAgreementsWriteToXml = blRet
'    Exit Function
'ErrHand:
'    mAgreementsWriteToXml = False
'End Function
Private Function mAddToArray() As Boolean
    Dim ilUpper As Integer
    Dim blRet As Boolean
On Error GoTo ErrHand
    blRet = True
    ilUpper = UBound(tmAgreements)
    tmAgreements(ilUpper) = tmAgreementInfo
    ReDim Preserve tmAgreements(ilUpper + 1)
    mAddToArray = blRet
    Exit Function
ErrHand:
    mAddToArray = False
End Function
Private Function mEarlierDate(slDate1 As String, slDate2 As String) As String
    Dim slRet As String
    
    slRet = ""
    If IsDate(slDate1) And IsDate(slDate2) Then
        If DateDiff("d", slDate1, slDate2) > -1 Then
            slRet = slDate1
        Else
            slRet = slDate2
        End If
    End If
    mEarlierDate = slRet
End Function

Private Function mAgreementGameEvent(ilVefCode As Integer, blError As Boolean, slVehicle As String, blSendToWeb As Boolean, Optional slStart As String, Optional slEnd As String) As Boolean
' return true if games/events sent and false if not.
' blError will stop updating
    Dim rst As ADODB.Recordset
    Dim blRet As Boolean
    Dim llHash As Long
    Dim blHash() As Byte
    '6809
    Dim slGameEndDate As String
    
    blRet = False
    blError = False
On Error GoTo ErrHand
    '6155 using different dates.  Don't care about the agreement's dates at this point, now concerned about the games.
'    slStart = Format$(slStart, sgSQLDateForm)
'    slEnd = Format$(slEnd, sgSQLDateForm)
    'try different dates
    If Len(slStart) = 0 Then
        slStart = Format$(smDate, sgSQLDateForm)
        slEnd = Format$(DateAdd("d", imNumberDays - 1, smDate), sgSQLDateForm)
    End If
    '6155
   ' SQLQuery = " SELECT distinct gsfXDSProgCodeId as programName FROM gsf_Game_Schd WHERE gsfVefCode = " & ilVefCode & " AND gsfAirDate >= '" & slStart & "' AND  gsfAirDate <= '" & slEnd & "' AND Ltrim(gsfxdsprogcodeid) <> '' ORDER BY gsfXDSProgCodeId"
    SQLQuery = " SELECT gsfXDSProgCodeId as programCode, gsfAirDate FROM gsf_Game_Schd WHERE gsfVefCode = " & ilVefCode & " AND gsfAirDate >= '" & slStart & "' AND  gsfAirDate <= '" & slEnd & "' AND Ltrim(gsfxdsprogcodeid) <> '' "
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If UCase$(Left$(rst!ProgramCode, 2)) <> "XX" Then
            blRet = True
            '6796
            tmAgreementInfo.sProgramCode = gXMLNameFilter(rst!ProgramCode)
            '6155 added
            blHash() = Format(rst!gsfAirDate, "mmddyy") & rst!ProgramCode & tmAgreementInfo.sProgramName & tmAgreementInfo.sStation
            llHash = mCalcCRC32(blHash)
            '6238
            'tmAgreementInfo.sCode = llHash
            tmAgreementInfo.sCode = llHash And &H7FFFFFFF
            '6809
             tmAgreementInfo.sStartDate = gAdjYear(Format$(rst!gsfAirDate, "m/d/yy"))
             slGameEndDate = gAdjYear(Format$(DateAdd("d", 1, rst!gsfAirDate), "m/d/yy"))
             slGameEndDate = slGameEndDate & " 06:00:00"
             tmAgreementInfo.sEndDate = slGameEndDate
            mAddToArray
            If blSendToWeb Then
                'gLogMsg "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle, smPathForgLogMsg, False
               ' myExport.WriteFacts "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle, False
                myExport.WriteFacts "Sending active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle & "/" & Trim$(tmAgreementInfo.sProgramCode), False
            Else
               ' gLogMsg "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle, smPathForgLogMsg, False
              ' myExport.WriteFacts "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle, False
               myExport.WriteFacts "Test send active agreement " & Trim(tmAgreementInfo.sStation) & " - " & slVehicle & "/" & Trim$(tmAgreementInfo.sProgramCode), False
            End If
            'mAgreementWriteXml
        End If
        rst.MoveNext
    Loop
Cleanup:
    If Not rst Is Nothing Then
        If (rst.State And adStateOpen) <> 0 Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    mAgreementGameEvent = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementGameEvent"
    'return true: this stops writing the basic agreement info, and thus stops updating.
    blRet = True
    blError = True
    GoTo Cleanup
End Function
Private Function mAgreementUpdate(slNeedUpdate As String, slStatus As String) As Boolean
    '8694
    Dim ilStoreMessagebox As Integer
    If Len(slNeedUpdate) > 0 Then
        SQLQuery = "UPDATE att set attSentToXDSStatus = '" & slStatus & "' WHERE attCode in ( " & slNeedUpdate & ")"
        '8694
        slNeedUpdate = ""
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            '8694
'            Screen.MousePointer = vbDefault
'            gHandleError smPathForgLogMsg, "ExportXDigital-mAgreementUpdate"
            ilStoreMessagebox = igShowMsgBox
            igShowMsgBox = 0
            gHandleError smPathForgLogMsg, "ExportXDigital-mAgreementUpdate"
            mSetResults "Could not update agreements sent. Does not affect export.", MESSAGERED
            igShowMsgBox = ilStoreMessagebox
            mAgreementUpdate = False
            Exit Function
        End If
    End If
    mAgreementUpdate = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "mAgreementUpdate"
    mAgreementUpdate = False

End Function
Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function
Private Function mAgreementWriteCumulus(tlAgreementInfo As XDIGITALAGREEMENTINFO) As Boolean
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo ErrHand
    With tlAgreementInfo
        mCSIXMLData "OT", "SIS_CLEARANCE_VW", ""
        mCSIXMLData "CD", "ClearanceId", .sCode
        mCSIXMLData "CD", "ClearanceStartDate", .sStartDate
        mCSIXMLData "CD", "ClearanceEndDate", .sEndDate
        mCSIXMLData "CD", "SiteId", .sSiteId
        mCSIXMLData "CD", "NetworkId", .sNetworkId
        mCSIXMLData "CD", "Status", .sStatus
        '6796
        If Len(.sProgramCode) > 0 Then
            mCSIXMLData "CD", "ProgramCode", .sProgramCode
        End If
        mCSIXMLData "CT", "SIS_CLEARANCE_VW", ""
    End With
    DoEvents
    mAgreementWriteCumulus = blRet
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementWriteCumulus"
    mAgreementWriteCumulus = False

End Function
Private Function mAgreementWriteCumulusGame(slAgreementInfo As XDIGITALAGREEMENTINFO) As Boolean
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo ErrHand
    With slAgreementInfo
        mCSIXMLData "OT", "Authorization", ""
        mCSIXMLData "CD", "AuthorizationId", .sCode
        mCSIXMLData "CD", "AuthorizationTestDate", ""
        mCSIXMLData "CD", "AuthorizationStartDate", .sStartDate
        mCSIXMLData "CD", "AuthorizationEndDate", .sEndDate
        mCSIXMLData "CD", "SiteId", .sSiteId
        mCSIXMLData "CD", "ProgramNumber", .sProgramCode
'        '5732
'        If Len(.sProgramName) > 0 Then
'            mCSIXMLData "CD", "ProgramName", .sProgramName
'        End If
        '6901
        'mCSIXMLData "CD", "Rights", "All"
        If bmIsAllPorts Then
            mCSIXMLData "CD", "Rights", "All"
        Else
            mCSIXMLData "OT", "Rights", ""
            mCSIXMLData "CD", "Live", ""
            mCSIXMLData "CD", "Record", "00:00:00"
            mCSIXMLData "CD", "Download", ""
            mCSIXMLData "CT", "Rights", ""
        End If

        mCSIXMLData "CT", "Authorization", ""
    End With
    DoEvents
    DoEvents
    mAgreementWriteCumulusGame = blRet
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementWriteCumulusGame"
    mAgreementWriteCumulusGame = False
End Function
Private Function mAgreementWriteXml(slAgreementInfo As XDIGITALAGREEMENTINFO) As Boolean
    'remove tlxmlstatus and ilret and comment out csixmlsetmethod and csixmlwrite(and following if statement and add to function that calls this to write to one file
    Dim blRet As Boolean
   ' Dim ilRet As Integer
    'Dim tlXmlStatus As CSIRspGetXMLStatus
   ' Dim slError As String
    
    blRet = True
On Error GoTo ErrHand
    '6581 remove
    'csiXMLSetMethod "SetAuthorizations", "Authorizations", "225", ""
    With slAgreementInfo
        mCSIXMLData "OT", "Authorization", ""
        mCSIXMLData "CD", "AuthorizationId", .sCode
        mCSIXMLData "CD", "AuthorizationTestDate", ""
        mCSIXMLData "CD", "AuthorizationStartDate", .sStartDate 'Format$(.sStartDate, sgShowDateForm)
        mCSIXMLData "CD", "AuthorizationEndDate", .sEndDate 'Format$(.sEndDate, sgShowDateForm)
        mCSIXMLData "CD", "SiteId", .sSiteId
        mCSIXMLData "CD", "ProgramCode", .sProgramCode
        '5732
        If Len(.sProgramName) > 0 Then
            mCSIXMLData "CD", "ProgramName", .sProgramName
        End If
        '5991
        If bmIsAllPorts Then
            mCSIXMLData "CD", "Rights", "All"
        Else
            mCSIXMLData "OT", "Rights", ""
            mCSIXMLData "CD", "Live", ""
            mCSIXMLData "CD", "Record", "00:00:00"
            mCSIXMLData "CD", "Download", ""
            mCSIXMLData "CT", "Rights", ""
        End If
        mCSIXMLData "CT", "Authorization", ""
    End With
    DoEvents
    '6581 remove
    'ilRet = csiXMLWrite(1)
    DoEvents
    '6581 remove
'    If ilRet <> True Then
'        blRet = False
'        '5896 now cancel on error, still continue on warning
'        csiXMLStatus tlXmlStatus
'        mIsXmlError tlXmlStatus.sStatus, "SetAuthorizations"
''        slError = "Problem with sending an authorization for " & slAgreementInfo.sProgramName & ", Program Code " & slAgreementInfo.sProgramCode & " for station " & slAgreementInfo.sStation
''        '5896
''        If gIsNull(tlXmlStatus.sStatus) Then
''            gLogMsg "ERROR: " & slError, "XDigitalExportLog.Txt", False
''        Else
''            gLogMsg "ERROR: " & slError & " " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''        End If
''        mSetResults slError, vbRed
'    End If
    mAgreementWriteXml = blRet
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAgreementWriteXml"
    mAgreementWriteXml = False
End Function
Private Function mStationSend(slXMLINIInputFile As String, ByVal slSection As String) As Boolean
    'return false only if meant to send but couldn't. Returns true if doesn't need to send
    Dim blRet As Boolean
    Dim slXUrl As String
    Dim blSendToWeb As Boolean
    Dim blNotCumulus As Boolean
    Dim blIsError As Boolean
    Dim slRet As String
    Dim slProp As String
    ' chose to send, or there are records to send
    blSendToWeb = True
    blRet = True
    If gIsSiteXDStation() Then
        'file
        If udcCriteria.XGenType(1, slProp) Then
            blSendToWeb = False
        End If
        'Dan M 04/01/14 added this 'block'
        If Len(smXmlErrorFile) = 0 Then
            blRet = False
            mSetResults "Problem sending Stations: error file doesn't exist.", MESSAGERED
            'gLogMsg "Problem Sending Stations: error file doesn't exist. Make sure the xml.ini 'logfile' has a valid directory.", smPathForgLogMsg, False
            myExport.WriteError "Problem Sending Stations: error file doesn't exist. Make sure the xml.ini 'logfile' has a valid directory.", True, False
            GoTo Cleanup
        End If
        'note there is no message because this could be set in ini as this, or this dual provider doesn't send.
        If bmSendAgreementIds = False And bmSendStationIds = False Then
            GoTo Cleanup
        End If
        blNotCumulus = Not mCumulusHeadEnd(slSection, slXMLINIInputFile, blIsError)
        If blIsError Then
            blRet = False
            mSetResults "Problem sending stations: Can't read WebServiceURL in xml.ini.", MESSAGERED
            myExport.WriteError "Problem Sending stations: Can't read WebServiceURL in xml.ini at " & slXMLINIInputFile, False, False
            GoTo Cleanup
        End If
        If igExportSource = 2 Then DoEvents
        'Cumulus needs the 2nd (or 3rd) xml.ini section ..sending to Vantive
        If blNotCumulus = False Then
            gLoadFromIni CUMULUSBACKOFFICESECTION, "WebServiceURL", slXMLINIInputFile, slRet
            'section doesn't exist
            If slRet = "Not Found" Then
                blRet = False
                mSetResults "Problem sending stations: Can't send to Cumulus head-end without '" & CUMULUSBACKOFFICESECTION & "' section in xml.ini.", MESSAGERED
                myExport.WriteError "Problem sending stations: Can't send to Cumulus head-end without '" & CUMULUSBACKOFFICESECTION & "' section in xml.ini.", False, False
                GoTo Cleanup
            Else
                'going to send to 'vantive'
                slSection = CUMULUSBACKOFFICESECTION
            End If
        End If
        'if slSection = CUMULUSBACKOFFICESECTION, its Cumulus!
        If Not mStationCreateFacts(slXMLINIInputFile, blSendToWeb, slSection) Then
            blRet = False
        End If
    End If
Cleanup:
    mStationSend = blRet
End Function
Private Function mStationUpdate(ByRef ilSentStation As Integer) As Boolean
    SQLQuery = "UPDATE shtt set shttSentToXDSStatus = 'Y' WHERE shttCode = " & ilSentStation
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "ExportXDigital-mStationUpdate"
        mStationUpdate = False
        Exit Function
    End If
    mStationUpdate = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "mStationUpdate"
    mStationUpdate = False
End Function
Private Function mStationCreateFacts(ByRef slXMLINIInputFile As String, blSendToWeb As Boolean, slSection As String) As Boolean
    'Rule: only unique stationIds in each 'set' of values.  So send one at a time!
    Dim blRet As Boolean
    Dim rstStation As ADODB.Recordset
    Dim slAddress As String
    Dim slAddress2 As String
    Dim slCity As String
    Dim slState As String
    Dim slZip As String
    Dim slStation As String
    Dim slBand As String
    Dim llStationID As Long
    Dim ilShttCode As Integer
    Dim blDontUpdate As Boolean
    Dim blError As Boolean
    Dim llCount As Long
    
    On Error GoTo ErrHand
    blRet = True
    llCount = 0
    'stations marked as 'M' or 'N' (not blank or 'Y'), and UsedForXDigital = 'Y'
    '6834
   ' If slSection <> CUMULUSVANTIVESECTION Then
        SQLQuery = "SELECT shttcode, shttCallLetters,shttONAddress1 as physicalAddress1,shttOnAddress2 as physicalAddress2,shttONCity as physicalCity," & _
        " shttONState as physicalState,shttONZip as physicalZip,shttAddress1 as mailAddress1, shttAddress2 as mailAddress2," & _
        " shttCity as mailCity, shttState as mailState, shttZip as mailZip, shttFrequency, shttStationId, arttFirstName, arttLastName" & _
        " FROM shtt left outer join artt  on shttownerarttcode = arttcode"
        SQLQuery = SQLQuery & " WHERE (shttsenttoxdsstatus = 'N' or shttsenttoxdsstatus = 'M') and shttUsedForXDigital = 'Y'" ' and shttStationId > 0"
'    Else
'        SQLQuery = "SELECT shttcode, shttCallLetters, shttFrequency, shttStationId" & _
'        " FROM shtt WHERE (shttsenttoxdsstatus = 'N' or shttsenttoxdsstatus = 'M') and shttUsedForXDigital = 'Y'" ' and shttStationId > 0"
'    End If
    ' StationId is not tested because I may get from agreement later.
    Set rstStation = gSQLSelectCall(SQLQuery)
    If Not rstStation.EOF Then
        If blSendToWeb Then
            csiXMLStart slXMLINIInputFile, slSection, "T", "", "", smXmlErrorFile
            '5896 '6632 added and slightly modified.
            'gLogMsg "!! Exporting Stations  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", smPathForgLogMsg, False
            myExport.WriteFacts "!! Exporting Stations  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
            Call mSetResults("Exporting Station Information", MESSAGEBLACK)
        Else
           ' csiXMLStart slXMLINIInputFile, slSection, "F", smExportDirectory & "StationInformation.txt", sgCRLF
            csiXMLStart slXMLINIInputFile, slSection, "F", smExportDirectory & STATIONLOG & smDateForLogs & ".txt", sgCRLF, smXmlErrorFile
            'gLogMsg "!! Writing Stations in test mode  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", smPathForgLogMsg, False
            myExport.WriteFacts "!! Writing Stations in test mode  Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", True
            Call mSetResults("Writing Station Information to StationInformation.txt", MESSAGEBLACK)
           ' Call mSetResults(smExportDirectory & "StationInformation.txt", RGB(0, 0, 0))
        End If
        Do While Not rstStation.EOF
            If igExportSource = 2 Then DoEvents
            'grab generic station info and place into tmStationInfo to be sent later.
            With rstStation
                ilShttCode = .Fields("shttCode")
                llStationID = .Fields("shttStationId").Value
                slStation = .Fields("shttCallLetters").Value
                slBand = ""
                mStationSplitBand slStation, slBand
                'only gather for nonCumulus
                'changed 6834
               ' If slSection <> CUMULUSVANTIVESECTION Then
                    If .Fields("shttCode").Value > 0 Then
                        gXDStationContact .Fields("shttCode").Value, tmStationInfo
                    End If
                    If Len(.Fields("physicalAddress1").Value) > 0 Then
                        slAddress = .Fields("physicaladdress1").Value
                        slAddress2 = .Fields("physicaladdress2").Value
                        slCity = .Fields("physicalCity").Value
                        slState = .Fields("physicalState").Value
                        slZip = .Fields("physicalzip").Value
                    Else
                        slAddress = .Fields("mailaddress1").Value
                        slAddress2 = .Fields("mailaddress2").Value
                        slCity = .Fields("mailCity").Value
                        slState = .Fields("mailState").Value
                        slZip = .Fields("mailzip").Value
                    End If
              '  End If
            End With
            With tmStationInfo
                .sBand = slBand
                .sCallLetters = slStation
                .sFrequency = rstStation.Fields("shttFrequency").Value
                '6834
               '  If slSection <> CUMULUSVANTIVESECTION Then
                    .sAddress = slAddress
                    .sAddress2 = slAddress2
                    .sCity = slCity
                    .sState = slState
                    .sZip = slZip
                    If Not IsNull(rstStation.Fields("ArttLastName").Value) Then
                        .sOwnership = rstStation.Fields("ArttLastName").Value
                    Else
                        .sOwnership = ""
                    End If
              '  End If
            End With
            'false means we need to send station id, either because an affiliate was blank, or there were no affiliate ids to send
             'blDontUpdate means there was an error  llCount is set here
           If Not mStationSendAffiliateSiteId(ilShttCode, blSendToWeb, llCount, slSection, blDontUpdate) Then
                'dan M 4/01/14 if got an error, stop
                'If llStationID > 0 Then
                If llStationID > 0 And Not blDontUpdate And bmSendStationIds Then
                    tmStationInfo.sSiteId = llStationID
                    If blSendToWeb Then
                        'gLogMsg "Sending station " & slStation & "-" & slBand & "-" & llStationID, smPathForgLogMsg, False
                        myExport.WriteFacts "Sending station " & slStation & "-" & slBand & "-" & llStationID
                    Else
                       ' gLogMsg "Test send station " & slStation & "-" & slBand & "-" & llStationID, smPathForgLogMsg, False
                       myExport.WriteFacts "Test send station " & slStation & "-" & slBand & "-" & llStationID
                    End If
                    If Not mStationWriteXml(slSection) Then
                        blDontUpdate = True
                        blRet = False
                        '5896
'                        If bmIsError Then
'                            Exit Do
'                        End If
                        'Dan M 4/1/14 added WrongServicePage.  This error shouldn't stop rest of exports, unlike bmIsError
                        If bmIsError Or bmIsWrongServicePage Then
                            Exit Do
                        End If
                    Else
                        llCount = llCount + 1
                   End If 'station write failed
                End If 'don't update
            End If 'send affiliate ids
            If blSendToWeb And Not blDontUpdate Then
                mStationUpdate ilShttCode
            'bldont update is true, and we had an error. stop trying to send stations
            ElseIf bmIsError Or bmIsWrongServicePage Then
                blRet = False
                Exit Do
            End If
            'dan 5/7/15 one warning is stopping all others
            blDontUpdate = False
            rstStation.MoveNext
        Loop
        csiXMLEnd
    End If 'recordset has records
Cleanup:
    If blSendToWeb Then
        If Not bmIsError And Not bmIsWrongServicePage Then
            mSetResults "Station Information Completed. Sent: " & llCount, MESSAGEBLACK
            'gLogMsg "Station Information Completed. Sent: " & llCount, smPathForgLogMsg, False
            myExport.WriteFacts "Station Information Completed. Sent: " & llCount, True
        End If
    Else
        mSetResults "Station Information Completed. Test send: " & llCount, MESSAGEBLACK
       ' gLogMsg "Station Information Completed. Test send: " & llCount, smPathForgLogMsg, False
       myExport.WriteFacts "Station Information Completed. Test send: " & llCount, True
    End If
    mStationCreateFacts = blRet
    If Not rstStation Is Nothing Then
        If (rstStation.State And adStateOpen) <> 0 Then
            rstStation.Close
        End If
        Set rstStation = Nothing
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    blRet = False
    gHandleError smPathForgLogMsg, "frmExportXDigital-mStationCreateFacts"
    GoTo Cleanup
End Function

Private Function mStationSendAffiliateSiteId(ilCode As Integer, blSendToWeb As Boolean, llCount As Long, slSection As String, blError As Boolean) As Boolean
    Dim rstATT As ADODB.Recordset
    Dim slSql As String
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim blRet As Boolean
    Dim llPreviousId As Long
    
    ' return false if some valid agreements were blank...thus we need to send station site id
On Error GoTo ErrHand
    If bmSendAgreementIds = False Then
        mStationSendAffiliateSiteId = False
        Exit Function
    End If
    blError = False
    blRet = True
    llPreviousId = 0
    If ilCode > 0 Then
        ' testing valid agreement (dates) AND vehicle set to send to xdigital(vffXDXMLForm)
        'how to test a set of dates?
        ' onAir is <= to the latest date, offAir,DropDate are both >= earliest date.  This seems to work as long as onair <= off/drop...so I test that after sql call
        slEarliestDate = Format$(smDate, sgSQLDateForm)
        slLatestDate = DateAdd("d", imNumberDays - 1, smDate)
        slLatestDate = Format$(slLatestDate, sgSQLDateForm)
        'Dan 8/26/14 this still has 'v60' way of finding isci...
'        slSql = "select attXDreceiverId as attID, attDropDate, attOnAir, attOffAir FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode " & _
'        " WHERE attXDReceiverId > 0 and attshfcode = " & ilCode & "  and  '" & slEarliestDate & "' < attOffAir and '" & slEarliestDate & "' < attdropDate AND '" & slLatestDate & "' >= attOnAir " & _
'        " AND vffxdxmlForm in ('A','P','S') ORDER BY attID"
        slSql = "select attXDreceiverId as attID, attDropDate, attOnAir, attOffAir FROM att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join vpf_Vehicle_Options on attvefCode = vpfvefKCode" & _
        " WHERE attXDReceiverId > 0 and attshfcode = " & ilCode & "  and  '" & slEarliestDate & "' < attOffAir and '" & slEarliestDate & "' < attdropDate AND '" & slLatestDate & "' >= attOnAir " & _
        " AND (vffxdxmlForm in ('A','S') OR vpfinterfaceId > 0)  ORDER BY attID"
        Set rstATT = gSQLSelectCall(slSql)
        Do While Not rstATT.EOF
            If igExportSource = 2 Then DoEvents
            If rstATT!attOnAir <= rstATT!attOffAir And rstATT!attOnAir <= rstATT!attDropDate Then
                If llPreviousId <> rstATT!attID Then
                    If blSendToWeb Then
                       ' gLogMsg "Sending station" & Trim(tmStationInfo.sCallLetters) & "-" & Trim(tmStationInfo.sBand) & "-" & rstATT!attid, smPathForgLogMsg, False
                        myExport.WriteFacts "Sending station" & Trim(tmStationInfo.sCallLetters) & "-" & Trim(tmStationInfo.sBand) & "-" & rstATT!attID
                    Else
                        'gLogMsg "Test send station " & Trim(tmStationInfo.sCallLetters) & "-" & Trim(tmStationInfo.sBand) & "-" & rstATT!attid, smPathForgLogMsg, False
                        myExport.WriteFacts "Test send station " & Trim(tmStationInfo.sCallLetters) & "-" & Trim(tmStationInfo.sBand) & "-" & rstATT!attID
                    End If
                    tmStationInfo.sSiteId = rstATT!attID
                    If mStationWriteXml(slSection) Then
                        llPreviousId = rstATT!attID
                        llCount = llCount + 1
                    Else
                        blError = True
                    End If
                End If
            End If
            rstATT.MoveNext
        Loop
        'Do we have to send station id?
        'Dan 8/26/14  have to change here too!
'        slSql = "select count(*) as amount from att inner join VFF_Vehicle_Features on attVefCode = vffVefCode" & _
'        " where attXDReceiverId = 0 and attshfcode = " & ilCode & "  and  '" & slEarliestDate & "' < attOffAir and '" & slEarliestDate & "' < attdropDate AND '" & slLatestDate & "' > attOnAir " & _
'        " AND vffxdxmlForm in ('A','P','S')"

        slSql = "select count(*) as amount from att inner join VFF_Vehicle_Features on attVefCode = vffVefCode inner join vpf_Vehicle_Options on attvefCode = vpfvefKCode " & _
        " where attXDReceiverId = 0 and attshfcode = " & ilCode & "  and  '" & slEarliestDate & "' < attOffAir and '" & slEarliestDate & "' < attdropDate AND '" & slLatestDate & "' > attOnAir " & _
        " AND (vffxdxmlForm in ('A','S') OR vpfinterfaceId > 0) "
        Set rstATT = gSQLSelectCall(slSql)
        If rstATT!amount > 0 Then
            blRet = False
        End If
    End If
Cleanup:
    If Not rstATT Is Nothing Then
        If (rstATT.State And adStateOpen) <> 0 Then
            rstATT.Close
        End If
        Set rstATT = Nothing
    End If
    mStationSendAffiliateSiteId = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mAffOrStationSiteId"
    blRet = False
    blError = True
    GoTo Cleanup
        
End Function
Private Function mStationWriteXml(slSection As String) As Boolean
   ' Dim ilRet As Integer
    Dim blRet As Boolean
   ' Dim tlXmlStatus As CSIRspGetXMLStatus
    Dim slStatus As String
    Dim slSetName As String
  '  Dim blCumulus As Boolean
    
    blRet = True
    
On Error GoTo ErrHand
'6834
'    If slSection = CUMULUSVANTIVESECTION Then
'        blCumulus = True
'        slSetName = "NewDataSet"
'    Else
      '  blCumulus = False
        slSetName = "Stations"
'    End If
    csiXMLSetMethod "SetStations", slSetName, "225", ""
'    If blCumulus Then
'        With tmStationInfo
'            mCSIXMLData "OT", "Table", ""
'            mCSIXMLData "CD", "SiteId", .sSiteId
'            mCSIXMLData "CD", "CallLetters", gXMLNameFilter(.sCallLetters)
'            mCSIXMLData "CD", "BandCode", gXMLNameFilter(.sBand)
'            mCSIXMLData "CD", "Frequency", gXMLNameFilter(.sFrequency)
'            mCSIXMLData "CD", "Slogan", ""
'            mCSIXMLData "CT", "Table", ""
'        End With
'    Else
 '       csiXMLSetMethod "SetStations", "Stations", "225", ""
        With tmStationInfo
            mCSIXMLData "OT", "Station", ""
            mCSIXMLData "CD", "SiteId", .sSiteId
            mCSIXMLData "CD", "CallLetters", gXMLNameFilter(.sCallLetters)
            mCSIXMLData "CD", "BandCode", gXMLNameFilter(.sBand)
            mCSIXMLData "CD", "Frequency", gXMLNameFilter(.sFrequency)
            mCSIXMLData "CD", "Ownership", gXMLNameFilter(.sOwnership)
            mCSIXMLData "CD", "Address", gXMLNameFilter(.sAddress)
            mCSIXMLData "CD", "Addr2", gXMLNameFilter(.sAddress2)
            mCSIXMLData "CD", "City", gXMLNameFilter(.sCity)
            mCSIXMLData "CD", "State", gXMLNameFilter(.sState)
            mCSIXMLData "CD", "Zip", gXMLNameFilter(.sZip)
            mCSIXMLData "CD", "Email", gXMLNameFilter(.sEmail)
            'already filtered
            mCSIXMLData "CD", "Contact1", .sContactName
            mCSIXMLData "CD", "Phone1", gXMLNameFilter(.sPhone)
            mCSIXMLData "CT", "Station", ""
            'Dan M 4/01/14 can't handle faultcodes.  Change to new method of reading
    '        ilRet = csiXMLWrite(1)
    '        If ilRet <> True Then
    '            blRet = False
    '            csiXMLStatus tlXmlStatus
    '            'Dan M 4/1/14 strip out ""
    '            slStatus = Replace(tlXmlStatus.sStatus, """", "")
    '            '5896 cancel on error, ignore warnng
    '            'mIsXmlError tlXmlStatus.sStatus, "SetStations"
    '            mIsXmlError slStatus, "SetStations"
    ''            '5896
    ''            If gIsNull(tlXmlStatus.sStatus) Then
    ''                gLogMsg "ERROR: Could not write station information for " & Trim$(tmStationInfo.sCallLetters) & Trim$(tmStationInfo.sBand), "XDigitalExportLog.Txt", False
    ''            Else
    ''                gLogMsg "ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
    ''            End If
    ''            'gLogMsg "ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
    ''            mSetResults "Problem with sending a station; please see XDigitalExportLog.Txt", vbRed
    '        End If
        End With
    'End If
    '7508
    'If Not mSendAndWriteReturn("SetStations") Then
    If Not mSendAndTestReturn(Stations) Then
        blRet = False
    End If
    mStationWriteXml = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mStationWriteXml"
    mStationWriteXml = False
End Function

Private Function mStationSplitBand(slStation As String, slBand As String) As Integer
    Dim ilPos As Integer
    Dim slWorkingStation As String
    
    slWorkingStation = Trim$(slStation)
    ilPos = InStr(1, slWorkingStation, "-", vbBinaryCompare)
    If ilPos > 0 And Len(slWorkingStation) = (ilPos + 2) Then
        slBand = Mid(slWorkingStation, ilPos + 1)
        slStation = Mid(slWorkingStation, 1, ilPos - 1)
    Else
        ilPos = 0
    End If
    mStationSplitBand = ilPos
End Function

Private Sub mSaveFD(slInVefCode As String, slInStationID As String, slInISCI As String, slCreativeTitle As String, llRotStartD As Long, llRotEndD As Long, slShortTitle As String, slXDXMLForm As String, slEventProgCodeID As String)
    'Build array of ISCI if File delivery requested
    Dim llUpper As Long
    Dim llLoop As Long
    Dim llIndex As Long
    Dim slKey As String
    Dim slVefCode As String
    Dim slStationID As String
    Dim slISCI As String
    Dim slProgCodeID As String
   ' Dim ilVpf As Integer
    '7169
    Dim slCreativeSafe As String
    
    '1/15/10:  Disallow File Delivery for ISCI until later
    'If (ckcExportType(1).Value = vbChecked) Then
    If igExportSource = 2 Then DoEvents
    '7162 don't need this anymore!
'    ilVpf = gBinarySearchVpf(CLng(slInVefCode))
'    If ilVpf <> -1 Then
        '7162 this works because the isci pass passes "" for slEventProgCodeID
        'If (udcCriteria.XExportType(1, "V") = vbChecked) And (tgVpfOptions(ilVpf).iInterfaceID <= 0) Then
        If (udcCriteria.XExportType(1, "V") = vbChecked) And (Len(slEventProgCodeID) > 0) Then
            slVefCode = slInVefCode
            Do While Len(slVefCode) < 5
                slVefCode = "0" & slVefCode
            Loop
            slStationID = slInStationID
            Do While Len(slStationID) < 10
                slStationID = "0" & slStationID
            Loop
            slISCI = slInISCI
            Do While Len(slISCI) < 80
                slISCI = " " & slISCI
            Loop
            slProgCodeID = slEventProgCodeID
            Do While Len(slProgCodeID) < 8
                slProgCodeID = " " & slProgCodeID
            Loop
            llUpper = UBound(tmXDFDInfo)
            slKey = slVefCode & "|" & slProgCodeID & "|" & slISCI & "|" & slStationID
            llIndex = mBinarySearchXDFD(slKey)
            If llIndex = -1 Then
                '7169
                slCreativeSafe = mSafeForTrim(slCreativeTitle, Len(tmXDFDInfo(llUpper).sCreativeTitle))
                tmXDFDInfo(llUpper).sKey = slKey
                tmXDFDInfo(llUpper).iVefCode = Val(slInVefCode)
                tmXDFDInfo(llUpper).lStationID = Val(slInStationID)
                tmXDFDInfo(llUpper).sISCI = slISCI
                'tmXDFDInfo(llUpper).sCreativeTitle = slCreativeTitle
                tmXDFDInfo(llUpper).sCreativeTitle = slCreativeSafe
                tmXDFDInfo(llUpper).lRotStartDate = llRotStartD
                tmXDFDInfo(llUpper).lRotEndDate = llRotEndD
                tmXDFDInfo(llUpper).sShortTitle = slShortTitle
                tmXDFDInfo(llUpper).sProgCodeID = slEventProgCodeID
                ReDim Preserve tmXDFDInfo(0 To llUpper + 1) As XDFDINFO
                If llUpper >= 1 Then
                    ArraySortTyp fnAV(tmXDFDInfo(), 0), UBound(tmXDFDInfo), 0, LenB(tmXDFDInfo(0)), 0, LenB(tmXDFDInfo(0).sKey), 0
                End If
            Else
                If llRotStartD < tmXDFDInfo(llIndex).lRotStartDate Then
                    tmXDFDInfo(llIndex).lRotStartDate = llRotStartD
                End If
                If llRotEndD > tmXDFDInfo(llIndex).lRotEndDate Then
                    tmXDFDInfo(llIndex).lRotEndDate = llRotEndD
                End If
            End If
        End If
    'End If
End Sub

Public Function mBinarySearchXDFD(slKey As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llResult As Long
    
    llMin = LBound(tmXDFDInfo)
    llMax = UBound(tmXDFDInfo) - 1
    Do While llMin <= llMax
        If igExportSource = 2 Then DoEvents
        llMiddle = (llMin + llMax) \ 2
        llResult = StrComp(Trim(tmXDFDInfo(llMiddle).sKey), Trim$(slKey), vbTextCompare)
        Select Case llResult
            Case 0:
                mBinarySearchXDFD = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    mBinarySearchXDFD = -1
    Exit Function
    
End Function


Private Function mExportFileDelivery() As Integer
    Dim llLoop As Long
    Dim llIndex As Long
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim slSave As String
    Dim ilVff As Integer
    Dim ilRet As Integer
    Dim llStation As Long
    Dim slTrackingCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slPackageCodeDate As String
    'removed for 6635
  '  Dim tlXmlStatus As CSIRspGetXMLStatus
    Dim slVefCode As String
    Dim slStationID As String
    Dim slRotStartDate As String
    Dim slRotEndDate As String
    Dim blRetStatus As Boolean
    '6635
    Dim llCount As Long
    
    'Dan 3/15/13
    mExportFileDelivery = False
    llCount = 0
    blRetStatus = True
    slStartDate = smDate
    slEndDate = DateAdd("d", imNumberDays - 1, smDate)
    If imNumberDays = 1 Then
        slPackageCodeDate = Format$(slStartDate, "mmddyy")
    Else
        slPackageCodeDate = Format$(slStartDate, "mmddyy") & "-" & Format$(slEndDate, "mmddyy")
    End If
    '2/3/12: Group Matching Rotation dates
    For llLoop = LBound(tmXDFDInfo) To UBound(tmXDFDInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        slVefCode = tmXDFDInfo(llLoop).iVefCode
        Do While Len(slVefCode) < 5
            slVefCode = "0" & slVefCode
        Loop
        slStationID = tmXDFDInfo(llLoop).lStationID
        Do While Len(slStationID) < 10
            slStationID = "0" & slStationID
        Loop
        slRotStartDate = tmXDFDInfo(llLoop).lRotStartDate
        Do While Len(slRotStartDate) < 6
            slRotStartDate = "0" & slRotStartDate
        Loop
        slRotEndDate = tmXDFDInfo(llLoop).lRotEndDate
        Do While Len(slRotEndDate) < 6
            slRotEndDate = "0" & slRotEndDate
        Loop
        tmXDFDInfo(llLoop).sKey = slVefCode & "|" & tmXDFDInfo(llLoop).sProgCodeID & "|" & tmXDFDInfo(llLoop).sISCI & "|" & slRotStartDate & "|" & slRotEndDate & "|" & slStationID
    Next llLoop
    If igExportSource = 2 Then DoEvents
    If UBound(tmXDFDInfo) > 1 Then
        ArraySortTyp fnAV(tmXDFDInfo(), 0), UBound(tmXDFDInfo), 0, LenB(tmXDFDInfo(0)), 0, LenB(tmXDFDInfo(0).sKey), 0
    End If
    llLoop = LBound(tmXDFDInfo)
    Do While llLoop < UBound(tmXDFDInfo)
        If igExportSource = 2 Then DoEvents
        ilVef = gBinarySearchVef(CLng(tmXDFDInfo(llLoop).iVefCode))
        If ilVef <> -1 Then
            ilVff = gBinarySearchVff(tmXDFDInfo(llLoop).iVefCode)
            ilVpf = gBinarySearchVpf(CLng(tmXDFDInfo(llLoop).iVefCode))
            If (ilVff <> -1) And (ilVpf <> -1) Then
                If igExportSource = 2 Then DoEvents
                slSave = ""
                If tgVffInfo(ilVff).sXDSaveCF <> "N" Then
                    slSave = "CF"
                End If
                If tgVffInfo(ilVff).sXDSaveHDD = "Y" Then
                    If slSave = "" Then
                        slSave = "HDD"
                    Else
                        slSave = slSave & "+HDD"
                    End If
                End If
                If tgVffInfo(ilVff).sXDSaveNAS = "Y" Then
                    If slSave = "" Then
                        slSave = "NAS"
                    Else
                        slSave = slSave & "+NAS"
                    End If
                End If
                'If Trim$(tgVffInfo(ilVff).sXDXMLForm) = "P" Then
                If tgVpfOptions(ilVpf).iInterfaceID > 0 Then
                    slSave = ""
                    If tgVffInfo(ilVff).sXDSSaveCF <> "N" Then
                        slSave = "CF"
                    End If
                    If tgVffInfo(ilVff).sXDSSaveHDD = "Y" Then
                        If slSave = "" Then
                            slSave = "HDD"
                        Else
                            slSave = slSave & "+HDD"
                        End If
                    End If
                    If tgVffInfo(ilVff).sXDSSaveNAS = "Y" Then
                        If slSave = "" Then
                            slSave = "NAS"
                        Else
                            slSave = slSave & "+NAS"
                        End If
                    End If
                    slTrackingCode = Trim$(Str$(tgVpfOptions(ilVpf).iInterfaceID))
                Else
                    If Trim$(UCase(tgVffInfo(ilVff).sXDProgCodeID)) <> "EVENT" Then
                        slTrackingCode = Trim$(tgVffInfo(ilVff).sXDProgCodeID)
                    Else
                        slTrackingCode = Trim$(tmXDFDInfo(llLoop).sProgCodeID)
                    End If
                End If
                llIndex = llLoop
                Do While tmXDFDInfo(llLoop).iVefCode = tmXDFDInfo(llIndex).iVefCode
                    If igExportSource = 2 Then DoEvents
                    If llLoop = llIndex Then
                        mCSIXMLData "OT", "FileDeliveryPackage", "Name=" & """" & gXMLNameFilter(Trim$(tgVehicleInfo(ilVef).sVehicleName)) & """"
                        'mCSIXMLData "CD", "PackageCode", slTrackingCode & "_" & Trim$(tmXDFDInfo(llIndex).sISCI)
                        mCSIXMLData "CD", "PackageCode", slPackageCodeDate & "_" & slTrackingCode
                        mCSIXMLData "OT", "Files", ""
                        'mCSIXMLData "OT", "File", "FileName=" & """" & Trim$(tmXDFDInfo(llIndex).sISCI) & ".MP2" & """"
                        If Trim$(tmXDFDInfo(llIndex).sShortTitle) = "" Then
                            '7496
                            'mCSIXMLData "OT", "File", "FileName=" & """" & Trim$(tmXDFDInfo(llIndex).sISCI) & ".MP2" & """"
                            mCSIXMLData "OT", "File", "FileName=" & """" & Trim$(tmXDFDInfo(llIndex).sISCI) & UCase(sgAudioExtension) & """"
                        Else
                            '7496
                            'mCSIXMLData "OT", "File", "FileName=" & """" & Trim$(tmXDFDInfo(llIndex).sShortTitle) & "(" & Trim$(tmXDFDInfo(llIndex).sISCI) & ")" & ".MP2" & """"
                            mCSIXMLData "OT", "File", "FileName=" & """" & Trim$(tmXDFDInfo(llIndex).sShortTitle) & "(" & Trim$(tmXDFDInfo(llIndex).sISCI) & ")" & UCase(sgAudioExtension) & """"
                        End If
                        mCSIXMLData "CD", "Title", Trim$(tmXDFDInfo(llIndex).sCreativeTitle)
                        mCSIXMLData "CD", "ISCI", Trim$(tmXDFDInfo(llIndex).sISCI)
                        mCSIXMLData "CD", "FileType", "Audio"
                        'mCSIXMLData "CD", "FileTrackingCode", slTrackingCode & "_" & Trim$(tmXDFDInfo(llIndex).sISCI)
                        '7496
                        'mCSIXMLData "CD", "FileTrackingCode", Trim$(tmXDFDInfo(llIndex).sISCI) & ".MP2"
                        mCSIXMLData "CD", "FileTrackingCode", Trim$(tmXDFDInfo(llIndex).sISCI) & UCase(sgAudioExtension)
                        mCSIXMLData "CD", "StartDate", Format(tmXDFDInfo(llIndex).lRotStartDate, "mm/dd/yyyy")
                        '2/9/12: add 14 days to end date to make sure that thye copy is not deleted to soon
                        'mCSIXMLData "CD", "EndDate", Format(tmXDFDInfo(llIndex).lRotEndDate, "mm/dd/yyyy")
                        mCSIXMLData "CD", "EndDate", Format(tmXDFDInfo(llIndex).lRotEndDate + 14, "mm/dd/yyyy")
                        mCSIXMLData "CD", "save", slSave
                        mCSIXMLData "CT", "File", ""
                        mCSIXMLData "CT", "Files", ""
                    End If
                    '2/3/12: Add date test
                    'If tmXDFDInfo(llLoop).sISCI <> tmXDFDInfo(llIndex).sISCI Then
                    If (tmXDFDInfo(llLoop).sISCI <> tmXDFDInfo(llIndex).sISCI) Or (tmXDFDInfo(llLoop).lRotStartDate <> tmXDFDInfo(llIndex).lRotStartDate) Or (tmXDFDInfo(llLoop).lRotEndDate <> tmXDFDInfo(llIndex).lRotEndDate) Then
                        If igExportSource = 2 Then DoEvents
                        mCSIXMLData "OT", "Stations", ""
                        For llStation = llLoop To llIndex - 1 Step 1
                            mCSIXMLData "OT", "Station", "SiteId = " & """" & tmXDFDInfo(llStation).lStationID & """"
                            mCSIXMLData "CT", "Station", ""
                        Next llStation
                        mCSIXMLData "CT", "Stations", ""
                        mCSIXMLData "CT", "FileDeliveryPackage", ""
                        '6635 replace below (which writes out for each file) with chunk size 5000
                        llCount = llCount + 1
                        If llCount Mod imChunk = 0 Then
                            '7508
                           ' If Not mSendAndWriteReturn("File Delivery") Then
                           If Not mSendAndTestReturn(FILEDELIVERY) Then
                                blRetStatus = False
                                If bmIsError Then
                                    Exit Function
                                End If
                            End If
                        End If
'                        ilRet = csiXMLWrite(1)
'                        If ilRet <> True Then
'                             blRetStatus = False
'                            ilRet = csiXMLStatus(tlXmlStatus)
'                            '5896  log error, change name of log to show user in display
'                             If mIsXmlError(tlXmlStatus.sStatus, "FileDelivery") Then
'                                 Exit Function
'                             End If
'                        End If
''                        If ilRet <> True Then
''                            '11/26/12: Continue
''                            'imTerminate = True
''                            'imExporting = False
''                            Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt for error", RGB(155, 0, 0))
''                            ilRet = csiXMLStatus(tlXmlStatus)
''                            '5896
''                            If gIsNull(tlXmlStatus.sStatus) Then
''                                gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR", "XDigitalExportLog.Txt", False
''                            Else
''                                gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                            End If
''                            'gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                            For llStation = llLoop To llIndex - 1 Step 1
''                                gLogMsg "      Station ID: " & tmXDFDInfo(llStation).lStationID, "XDigitalExportLog.Txt", False
''                            Next llStation
''                            blRetStatus = False
''                            '11/26/12: Continue
''                            'Exit Function
''                        End If
                        llLoop = llIndex
                        llIndex = llIndex - 1
                    End If
                    llIndex = llIndex + 1
                    If llIndex >= UBound(tmXDFDInfo) Then
                        Exit Do
                    End If
                Loop
                If igExportSource = 2 Then DoEvents
                'Dan what's happening?  Previous loop didn't write out everything needed.  Pick up stragglers and write here
                mCSIXMLData "OT", "Stations", ""
                'mCSIXMLData "OT", "Station", "SiteId = " & """" & tmXDFDInfo(llLoop).lStationID & """"
                'mCSIXMLData "CT", "Station", ""
                For llStation = llLoop To llIndex - 1 Step 1
                    mCSIXMLData "OT", "Station", "SiteId = " & """" & tmXDFDInfo(llStation).lStationID & """"
                    mCSIXMLData "CT", "Station", ""
                Next llStation
                mCSIXMLData "CT", "Stations", ""
                mCSIXMLData "CT", "FileDeliveryPackage", ""
                '6635
                '7508
                'If Not mSendAndWriteReturn("File Delivery") Then
                If Not mSendAndTestReturn(FILEDELIVERY) Then
                    blRetStatus = False
                    If bmIsError Then
                        Exit Function
                    End If
                End If
'                ilRet = csiXMLWrite(1)
'                If ilRet <> True Then
'                     blRetStatus = False
'                    ilRet = csiXMLStatus(tlXmlStatus)
'                    '5896  log error, change name of log to show user in display
'                     If mIsXmlError(tlXmlStatus.sStatus, "FileDelivery") Then
'                         Exit Function
'                     End If
'                    'Dan 3/13/13 continue to next vehicle
'                   ' Exit Do
'                End If
'
''                If ilRet <> True Then
''                    'imTerminate = True
''                    'imExporting = False
''                    Call mSetResults("Export not completely successful. see XDigitalExportLog.Txt for error", RGB(155, 0, 0))
''                    ilRet = csiXMLStatus(tlXmlStatus)
''                    'gLogMsg "Station ID: " & tmXFDInfo(llStation).lStationID & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                    '5896
''                    If gIsNull(tlXmlStatus.sStatus) Then
''                        gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR", "XDigitalExportLog.Txt", False
''                    Else
''                        gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                    End If
''                   ' gLogMsg "Vehicle " & Trim$(tgVehicleInfo(ilVef).sVehicle) & " ERROR: " & tlXmlStatus.sStatus, "XDigitalExportLog.Txt", False
''                    For llStation = llLoop To llIndex - 1 Step 1
''                        gLogMsg "      Station ID: " & tmXDFDInfo(llStation).lStationID, "XDigitalExportLog.Txt", False
''                    Next llStation
''                    blRetStatus = False
''                    '11/26/12: Continue
''                    'Exit Function
''                End If
                llLoop = llIndex - 1
            End If
        End If
        llLoop = llLoop + 1
    Loop
    mExportFileDelivery = blRetStatus
End Function

Private Sub mClearAlerts(llSDate As Long, llEDate As Long)
    Dim ilVef As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim llStartDate As Long
    Dim ilRet As Integer
    
    slDate = gObtainPrevMonday(Format(llSDate, "m/d/yy"))
    llStartDate = gDateValue(slDate)
'    For ilVef = 0 To lbcVehicles.ListCount - 1
'        If igExportSource = 2 Then DoEvents
'        If lbcVehicles.Selected(ilVef) Then
    For ilVef = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilVef, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilVef, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilVef, VEHCODEINDEX)
                'imVefCode = lbcVehicles.ItemData(ilVef)
                For llDate = llStartDate To llEDate Step 7
                    If igExportSource = 2 Then DoEvents
                    slDate = Format$(llDate, "m/d/yy")
                    ilRet = gAlertClear("A", "F", "S", imVefCode, slDate)
                    ilRet = gAlertClear("A", "R", "S", imVefCode, slDate)
                Next llDate
            End If
        End If
    Next ilVef
    ilRet = gAlertForceCheck()
End Sub

Private Sub menuErrorFile_Click()
    Dim slError As String
    Dim slRet As String
    
    slRet = gGetXmlErrorFile(slError)
    MsgBox slRet
End Sub

Private Sub mnuAuthorization_Click()
    mnuAuthorization.Checked = True
    mnuProgram.Checked = False
    mnuNormal.Checked = False
    mnuStation.Checked = False
End Sub

Private Sub mnuNormal_Click()
    mnuAuthorization.Checked = False
    mnuProgram.Checked = False
    mnuStation.Checked = False
    mnuNormal.Checked = True
End Sub

Private Sub mnuProgram_Click()
    mnuAuthorization.Checked = False
    mnuProgram.Checked = True
    mnuStation.Checked = False
    mnuNormal.Checked = False
End Sub

Private Sub mnuStation_Click()
    mnuAuthorization.Checked = False
    mnuProgram.Checked = False
    mnuStation.Checked = True
    mnuNormal.Checked = False
End Sub

Private Sub mnuTestChanges_Click()
    If mnuTestChanges.Checked Then
        mnuTestChanges.Checked = False
        bmTestForceUpdateXHT = False
    Else
        mnuTestChanges.Checked = True
        bmTestForceUpdateXHT = True
    End If
End Sub
Private Sub mnuTestErrors_Click()
    If mnuTestErrors.Checked Then
        bmTestError = False
        mnuTestErrors.Checked = False
    Else
        bmTestError = True
        mnuTestErrors.Checked = True
    End If
End Sub

Private Sub tmcDelay_Timer()
    '8163
    tmcDelay.Enabled = False
    If IsDate(edcDate.Text) Then
        If lbcStation.Visible Then
            'lbcVehicles_Click
            grdVeh_Click
        End If
    'Else
    '    tmcDelay.Enabled = True
    End If
    mSetLogPgmSplitColumns
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload FrmExportXDigital
End Sub
Private Sub mSaveCustomValues()
    Dim ilLoop As Integer
    ReDim ilVefCode(0 To 0) As Integer
    ReDim ilShttCode(0 To 0) As Integer
    If igExportSource <> 2 Then
        ReDim tgEhtInfo(0 To 1) As EHTINFO
        ReDim tgEvtInfo(0 To 0) As EVTINFO
        ReDim tgEctInfo(0 To 0) As ECTINFO
        lgExportEhtInfoIndex = 0
        tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1
'        For ilLoop = 0 To lbcVehicles.ListCount - 1
'            If lbcVehicles.Selected(ilLoop) Then
'                ilVefCode(UBound(ilVefCode)) = lbcVehicles.ItemData(ilLoop)
        For ilLoop = 1 To grdVeh.Rows - 1 Step 1
            If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
                If grdVeh.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                    ilVefCode(UBound(ilVefCode)) = grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
            End If
        Next ilLoop
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("X", "X-Digital", "X", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub


Private Function mGetProgCode(slProgCodeID As String, llGsfCode As Long) As String
    On Error GoTo ErrHandler
    mGetProgCode = slProgCodeID
    If UCase(slProgCodeID) = "EVENT" Then
        If lmEventGsfCode <> llGsfCode Then
            SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfCode = " & llGsfCode & ")"
            Set rst_Gsf = gSQLSelectCall(SQLQuery)
            If Not rst_Gsf.EOF Then
                lmEventGsfCode = llGsfCode
                smEventProgCodeID = Trim$(rst_Gsf!gsfXDSProgCodeID)
                mGetProgCode = smEventProgCodeID
            End If
        Else
            mGetProgCode = smEventProgCodeID
        End If
    End If
    Exit Function
ErrHandler:
    gHandleError smPathForgLogMsg, "frmExportXDigital-mGetProgCode"
    Exit Function
End Function

'Private Function mAdjustToEstZone(slInZone As String, slDate As String, slTime As String) As Integer
'    Dim slZone As String
'    Dim ilZone As Integer
'    Dim ilLocalAdj As Integer
'    Dim ilZoneFound As Integer
'    Dim ilNumberAsterisk As Integer
'    Dim llSpotTime As Long
'    Dim llSpotDate As Long
'    Dim ilVef As Integer
'
'    mAdjustToEstZone = 0
'    If smMidnightBasedHours <> "Y" Then
'        Exit Function
'    End If
'    llSpotDate = gDateValue(slDate)
'    llSpotTime = gTimeToLong(slTime, False)
'    slZone = UCase$(Trim$(slInZone))
'    ilLocalAdj = 0
'    ilZoneFound = False
'    ilNumberAsterisk = 0
'    ' Adjust time zone properly.
'    ilVef = gBinarySearchVef(CLng(imVefCode))
'    If (Len(slZone) <> 0) And (ilVef <> -1) Then
'        'Get zone
'        DoEvents
'        For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
'            If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = slZone Then
'                If tgVehicleInfo(ilVef).sFed(ilZone) <> "*" Then
'                    slZone = tgVehicleInfo(ilVef).sZone(tgVehicleInfo(ilVef).iBaseZone(ilZone))
'                    ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
'                    ilZoneFound = True
'                End If
'                Exit For
'            End If
'        Next ilZone
'        For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
'            If tgVehicleInfo(ilVef).sFed(ilZone) = "*" Then
'                ilNumberAsterisk = ilNumberAsterisk + 1
'            End If
'        Next ilZone
'    End If
'    If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
'        slZone = ""
'    End If
'    If ilLocalAdj = 0 Then
'        Exit Function
'    End If
'    ilLocalAdj = -1 * ilLocalAdj
'    llSpotTime = llSpotTime + 3600 * ilLocalAdj
'    If llSpotTime < 0 Then
'        llSpotTime = llSpotTime + 86400
'        llSpotDate = llSpotDate - 1
'    ElseIf llSpotTime > 86400 Then
'        llSpotTime = llSpotTime - 86400
'        llSpotDate = llSpotDate + 1
'    End If
'    slDate = Format(llSpotDate, "m/d/yy")
'    slTime = gLongToTime(llSpotTime)
'    mAdjustToEstZone = ilLocalAdj
'
'End Function
Private Sub mAdjustToHeadendZone(ilLocalAdj As Integer, slDate As String, slTime As String)
    Dim llSpotTime As Long
    Dim llSpotDate As Long
   '9629 removed this block
'    If smMidnightBasedHours <> "Y" Then
'        Exit Sub
'    End If
    llSpotDate = gDateValue(slDate)
    llSpotTime = gTimeToLong(slTime, False)
    llSpotTime = llSpotTime + 3600 * ilLocalAdj
    If llSpotTime < 0 Then
        llSpotTime = llSpotTime + 86400
        llSpotDate = llSpotDate - 1
    ElseIf llSpotTime > 86400 Then
        llSpotTime = llSpotTime - 86400
        llSpotDate = llSpotDate + 1
    End If
    slDate = Format(llSpotDate, "m/d/yy")
    slTime = gLongToTime(llSpotTime)
End Sub
'
'Private Sub mXMLSiteTags(ilPassForm As Integer, slXDReceiverID As String, slTransmissionID As String, slUnitHB As String, slUnitHBP As String, slVefCode5 As String, llAstCode As Long)
'    Dim slUnitIDAstCode As String
'    '8357
'    Dim slUnitIdForDeleteComparison As String
'
'    slUnitIdForDeleteComparison = ""
'    If ilPassForm = 0 Then
'        mCSIXMLData "CT", "Insert", ""
'        mCSIXMLData "CT", "Site", ""
'    ElseIf (ilPassForm = 1) Or (ilPassForm = 2) Then
'        mCSIXMLData "OT", "Sites", ""
'        If smUnitIdByAstCodeForBreak = "Y" Then
'            slUnitIDAstCode = Trim$(Str$(llAstCode))
'            Do While Len(slUnitIDAstCode) < 9
'                slUnitIDAstCode = "0" & slUnitIDAstCode
'            Loop
'            '9818
'            If imSharedHeadEndCue > 0 Then
'                slUnitIDAstCode = imSharedHeadEndCue & slUnitIDAstCode
'            End If
'            If ilPassForm = 1 Then
'                '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
'                'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHBP & """"
'                mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slUnitIDAstCode & """"
'            Else
'                '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
'                'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHB & """"
'                mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slUnitIDAstCode & """"
'                '9113
'                slUnitIdForDeleteComparison = slUnitIDAstCode
'            End If
'        Else
'            '9818
'            If imSharedHeadEndCue > 0 Then
'                If ilPassForm = 1 Then
'                    mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & imSharedHeadEndCue & Mid(slTransmissionID, 4) & slUnitHBP & slVefCode5 & """"
'                Else
'                    mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & imSharedHeadEndCue & Mid(slTransmissionID, 4) & slUnitHB & slVefCode5 & """"
'                    slUnitIdForDeleteComparison = imSharedHeadEndCue & Mid(slTransmissionID, 4) & slUnitHB & slVefCode5
'                End If
'            Else
'                'original:
'                If ilPassForm = 1 Then
'                    '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
'                    'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHBP & """"
'                    mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5 & """"
'                Else
'                    '9/6/11:  Add vehicle code to Unit ID so that it is unique across vehicles
'                    'mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHB & """"
'                    mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slUnitHB & slVefCode5 & """"
'                    slUnitIdForDeleteComparison = slTransmissionID & slUnitHB & slVefCode5
'                End If
'            End If
'        End If
'        mCSIXMLData "CT", "Site", ""
'        mCSIXMLData "CT", "Sites", ""
'        mCSIXMLData "CT", "Insert", ""
'        '8357
'        mRetainSiteIdAndUnitId slXDReceiverID, slUnitIdForDeleteComparison
'    End If
'
'End Sub
'10021
Private Sub mXMLSiteTags(ilPassForm As Integer, slXDReceiverID As String, slTransmissionID As String, slUnitHB As String, slUnitHBP As String, slVefCode5 As String, llAstCode As Long)
    Dim slUnitID As String

    If ilPassForm = ISCIFORM Then
        mCSIXMLData "CT", "Insert", ""
        mCSIXMLData "CT", "Site", ""
    Else  '  can I lose this?If (ilPassForm = 1) Or (ilPassForm = 2) Then
        If ilPassForm = HBPFORM Then
            slUnitID = slUnitHBP
        Else
            slUnitID = slUnitHB
        End If
        slUnitID = mCreateUnitIDForCue(ilPassForm, slUnitID & slVefCode5, llAstCode, slTransmissionID)
        mCSIXMLData "OT", "Sites", ""
        mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slUnitID & """"
        mCSIXMLData "CT", "Site", ""
        mCSIXMLData "CT", "Sites", ""
        mCSIXMLData "CT", "Insert", ""
        '8357
        mRetainSiteIdAndUnitId slXDReceiverID, slUnitID
    End If

End Sub
Private Function mAstFileSave(slXDSiteId As String, slCall As String, slISCI As String, slCreative As String, llAstIndex As Long, slTransmission As String, slProgramCode As String, blIsRegional As Boolean, slVehicle As String) As Boolean
    Dim llContract As Long
    Dim slAdv As String
    Dim slAgency As String
    Dim slBuyer As String
    '7161 use feed date/time instead
'    Dim slPledgeDate As String
'    Dim slPledgeTime As String
    Dim slFeedDate As String
    Dim slFeedTime As String
    Dim slUnitID As String
    Dim llIsciId As Long
    Dim ilAdvId As Integer
    Dim ilAgencyId As Long
    Dim llContractId As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llXDSite As Long
    Dim ilVef As Integer

On Error GoTo ERRORBOX
'    slPledgeDate = ""
'    slPledgeTime = ""
    slFeedDate = ""
    slFeedTime = ""
    ilAdvId = 0
    llContractId = 0
    'not filled at this time
    ilAgencyId = 0
    slAgency = ""
    slBuyer = ""
    slStartDate = ""
    slEndDate = ""
    slBuyer = ""
    llXDSite = CLng(slXDSiteId)
    With tmAstInfo(llAstIndex)
        slUnitID = Trim$(Str$(.lCode))
        Do While Len(slUnitID) < 9
            slUnitID = "0" & slUnitID
        Loop
        ilVef = .iVefCode
'        slPledgeDate = .sPledgeDate
'        slPledgeTime = Format(.sPledgeStartTime, "HH:NN:SS")
        slFeedDate = .sFeedDate
        slFeedTime = Format(.sFeedTime, "HH:NN:SS")
        ilAdvId = .iAdfCode
        If blIsRegional Then
            llIsciId = .lRCpfCode
        Else
            llIsciId = .lCpfCode
        End If
        '7161 changed pledge to feed
        rsAstFiles.AddNew Array("UnitId", "Agency", "Client", "Buyer", "ISCI", "Contract", "FeedDate", "FeedTime", "CallLetters", "isciID", "StationID", "AdvID", "AgencyID", "ContractID", "CntrStart", "CntrEnd", "XDSiteId", "ProgramID", "ProgramName") _
        , Array(slUnitID, slAgency, slAdv, slBuyer, slISCI, .lCntrNo, slFeedDate, slFeedTime, slCall, llIsciId, .iShttCode, ilAdvId, ilAgencyId, llContractId, slStartDate, slEndDate, llXDSite, ilVef, slVehicle)
    End With
    mAstFileSave = True
Exit Function
ERRORBOX:
    'gLogMsg "Error in mAstFileSave.  Error: " & Err.Description, smPathForgLogMsg, False
    myExport.WriteError "mAstFileSave.  Error: " & Err.Description, False, False
    mAstFileSave = False
End Function
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
 With myRs.Fields
        .Append "UnitId", adChar, 9
        .Append "Agency", adChar, 40
        .Append "Client", adChar, 30
        .Append "Buyer", adChar, 20
        .Append "ISCI", adChar, 40
        .Append "Contract", adInteger
        '7161 from pledge to feed
'        .Append "PledgeDate", adChar, 10
'        .Append "PledgeTime", adChar, 8
        .Append "FeedDate", adChar, 10
        .Append "FeedTime", adChar, 8
        .Append "CallLetters", adChar, 40
        .Append "AgencyID", adInteger
        .Append "ContractID", adInteger
        .Append "StationID", adInteger
        .Append "ISCIID", adInteger
        .Append "AdvID", adInteger
        .Append "CntrStart", adChar, 10
        .Append "CntrEnd", adChar, 10
        .Append "ProgramID", adInteger
        .Append "ProgramName", adChar, 40
        .Append "XDSiteId", adInteger
    End With
    myRs.Open
    '7161 changed
    myRs!FeedDate.Properties("optimize") = True
    myRs.Sort = "FeedDate desc"
    Set mPrepRecordset = myRs
End Function


Private Function mBuildMergeAstInfo(slInMoDate As String, slEndDate As String, ilPassForm As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilOkStation As Integer
    Dim ilVefCode As Integer
    Dim slMoDate As String
    Dim slSDate As String
    Dim slEDate As String
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim llAst As Long
    Dim slStr As String
    Dim ilLocalAdj As Integer
    
    mBuildMergeAstInfo = True
    ReDim tmMergeAstInfo(0 To 0) As ASTINFO
    If UBound(imMergeVefCode) <= LBound(imMergeVefCode) Then
        Exit Function
    End If
    slMoDate = slInMoDate
    slStr = ""
    For ilLoop = 0 To UBound(imMergeVefCode) - 1 Step 1
        If slStr = "" Then
            slStr = Trim$(Str$(imMergeVefCode(ilLoop)))
        Else
            slStr = slStr & ", " & Trim$(Str$(imMergeVefCode(ilLoop)))
        End If
    Next ilLoop
    
    Do
        If igExportSource = 2 Then DoEvents
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attVoiceTracked, attXDReceiverID, vefName"
        SQLQuery = SQLQuery & " FROM shtt, cptt, att, vef_Vehicles"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery & " AND vefCode = cpttVefCode"
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND vefCode IN (" & slStr & ")" & ")"
        SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            
            lacProcessing.Caption = "Checking " & Trim$(cprst!vefName)
            If igExportSource = 2 Then DoEvents
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If lbcStation.Selected(ilLoop) Then
                        If lbcStation.ItemData(ilLoop) = cprst!shttCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next ilLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation Then
                On Error GoTo ErrHand
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cprst!cpttCode
                tgCPPosting(0).iStatus = cprst!cpttStatus
                tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cprst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cprst!attTimeType
                tgCPPosting(0).iVefCode = cprst!cpttvefcode
                tgCPPosting(0).iShttCode = cprst!shttCode
                tgCPPosting(0).sZone = cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(smDate, sgShowDateForm) 'Format$(sMoDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                tgCPPosting(0).iNumberDays = imNumberDays
                'Create AST records
                igTimes = 3 'By Date
                imAdfCode = -1
                If igExportSource = 2 Then DoEvents
                If ilPassForm = 0 Then
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , , , , True)
                    gFilterAstExtendedTypes tmAstInfo
                    mRemoveExtraAirplays
                Else
                    'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                    If (smMidnightBasedHours = "Y") And (ilPassForm <> 0) Then
                        '6082 change first 'false' to 'true to get rid of 0 in astcode
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , True, , , True)
                        gFilterAstExtendedTypes tmAstInfo
                        mRemoveExtraAirplays
                        '10938 replaced how to get ilLocalAdj
'                        myZoneAndDSTHelper.StationZone = tgCPPosting(0).sZone
'                        myZoneAndDSTHelper.StationHonorDaylight ilAcknowledgeDaylight
                        'ilLocalAdj = myZoneAndDSTHelper.FindZoneDifference
                        '3/17/16: Handle any head end zone
                        ilLocalAdj = mStationAdj(tgCPPosting(0).sZone)
                        'If Left(tgCPPosting(0).sZone, 1) <> "E" Then
                        If ilLocalAdj <> 0 Then
                            ReDim tmAstAdj1(LBound(tmAstInfo) To UBound(tmAstInfo)) As ASTINFO
                            For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                                tmAstAdj1(llAst) = tmAstInfo(llAst)
                                '3/17/16: Handle any head end zone
                                'ilLocalAdj = mAdjustToEstZone(tgCPPosting(0).sZone, tmAstAdj1(llAst).sFeedDate, tmAstAdj1(llAst).sFeedTime)
                                mAdjustToHeadendZone ilLocalAdj, tmAstAdj1(llAst).sFeedDate, tmAstAdj1(llAst).sFeedTime
                            Next llAst
                            ReDim tmAstInfo(LBound(tmAstAdj1) To UBound(tmAstAdj1)) As ASTINFO
                            For llAst = LBound(tmAstAdj1) To UBound(tmAstAdj1) - 1 Step 1
                                tmAstInfo(llAst) = tmAstAdj1(llAst)
                            Next llAst
                        End If
                    Else
                        '6082 change first 'false' to 'true to get rid of 0 in astcode
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True, , , , , , True)
                        gFilterAstExtendedTypes tmAstInfo
                        mRemoveExtraAirplays
                    End If
                End If
                'Merge
                llUpper = UBound(tmMergeAstInfo)
                ReDim Preserve tmMergeAstInfo(0 To llUpper + UBound(tmAstInfo)) As ASTINFO
                For llAst = 0 To UBound(tmAstInfo) - 1 Step 1
                    tmMergeAstInfo(llUpper) = tmAstInfo(llAst)
                    llUpper = llUpper + 1
                Next llAst
            End If
            cprst.MoveNext
        Wend
        slMoDate = DateAdd("d", 7, slMoDate)
        slSDate = slMoDate
        slEDate = gObtainNextSunday(slSDate)
        If gDateValue(gAdjYear(slEndDate)) < gDateValue(gAdjYear(slEDate)) Then
            slEDate = slEndDate
        End If
    Loop While gDateValue(gAdjYear(slMoDate)) < gDateValue(gAdjYear(slEndDate))
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    'Dan M 11/7/14 changed name of function
    gHandleError smPathForgLogMsg, "frmExportXDigital-mBuildMergeAstInfo"
    Resume Next
    mBuildMergeAstInfo = False
    Exit Function
End Function

Private Sub mMergeAsts(ilVpf As Integer)
    Dim llAst As Long
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llTime As Long
    Dim llMerge As Long
    Dim ilTestDay As Integer
    Dim llTestTime1 As Long
    Dim llTestTime2 As Long
    Dim llMove As Long
    Dim llVpf As Long
    Dim llDayLoop As Long
    Dim ilFirstSpotOfDay As Integer
    Dim ilLastSpotOfDay As Integer
    Dim llLimit As Long
    
    If ilVpf <> -1 Then
        If (Asc(tgVpfOptions(ilVpf).sUsingFeatures2) And XDSAPPLYMERGE) <> XDSAPPLYMERGE Then
            Exit Sub
        End If
    End If
    For llAst = 0 To UBound(tmMergeAstInfo) - 1 Step 1
        llDate = gDateValue(tmMergeAstInfo(llAst).sFeedDate)
        ilDay = Weekday(tmMergeAstInfo(llAst).sFeedDate, vbMonday) - 1
        llTime = gTimeToLong(tmMergeAstInfo(llAst).sFeedTime, False)
        For llDayLoop = 0 To UBound(tmAstTimeRange) - 1 Step 1
            If (llDate = tmAstTimeRange(llDayLoop).lDate) And (llTime >= tmAstTimeRange(llDayLoop).lStartTime) And (llTime <= tmAstTimeRange(llDayLoop).lEndTime) Then
                'Merge
                ilFirstSpotOfDay = -1
                ilLastSpotOfDay = -1
                llLimit = 2
                For llMerge = 0 To UBound(tmAstInfo) - 1 Step 1
                    If tmMergeAstInfo(llAst).iShttCode = tmAstInfo(llMerge).iShttCode Then
                        ilTestDay = Weekday(tmAstInfo(llMerge).sFeedDate, vbMonday) - 1
                        If (ilDay = ilTestDay) And (ilFirstSpotOfDay = -1) Then
                            ilFirstSpotOfDay = llMerge
                        End If
                        If (ilDay = ilTestDay) Then
                            ilLastSpotOfDay = llMerge
                        End If
                    End If
                Next llMerge
                'If UBound(tmAstInfo) = 1 Then
                    llLimit = 1
                'End If
                For llMerge = 0 To UBound(tmAstInfo) - llLimit Step 1
                    If tmMergeAstInfo(llAst).iShttCode = tmAstInfo(llMerge).iShttCode Then
                        ilTestDay = Weekday(tmAstInfo(llMerge).sFeedDate, vbMonday) - 1
                        If ilDay = ilTestDay Then
                            llTestTime1 = gTimeToLong(tmAstInfo(llMerge).sFeedTime, False)
                            If llMerge < UBound(tmAstInfo) - 1 Then
                                ilTestDay = Weekday(tmAstInfo(llMerge + 1).sFeedDate, vbMonday) - 1
                                If ilDay = ilTestDay Then
                                    llTestTime2 = gTimeToLong(tmAstInfo(llMerge + 1).sFeedTime, False)
                                Else
                                    llTestTime2 = 86400
                                End If
                            Else
                                llTestTime2 = 86400
                            End If
                            If (llTime < llTestTime1) And (llMerge = ilFirstSpotOfDay) Then
                                ReDim tmAstAdj1(0 To UBound(tmAstInfo) + 1) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    tmAstAdj1(llMove + 1) = tmAstInfo(llMove)
                                Next llMove
                                tmMergeAstInfo(llAst).lCode = Abs(tmMergeAstInfo(llAst).lCode)
                                tmAstAdj1(0) = tmMergeAstInfo(llAst)
                                tmMergeAstInfo(llAst).lCode = -tmMergeAstInfo(llAst).lCode
                                ReDim tmAstInfo(0 To UBound(tmAstAdj1)) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    tmAstInfo(llMove) = tmAstAdj1(llMove)
                                Next llMove
                                Exit For
                            ElseIf (llTime >= llTestTime1) And (llTime < llTestTime2) Then
                                ReDim tmAstAdj1(0 To UBound(tmAstInfo) + 1) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    If llMove <= llMerge Then
                                        tmAstAdj1(llMove) = tmAstInfo(llMove)
                                        If llMove = llMerge Then
                                            tmMergeAstInfo(llAst).lCode = Abs(tmMergeAstInfo(llAst).lCode)
                                            tmAstAdj1(llMove + 1) = tmMergeAstInfo(llAst)
                                            tmMergeAstInfo(llAst).lCode = -tmMergeAstInfo(llAst).lCode
                                        End If
                                    Else
                                        tmAstAdj1(llMove + 1) = tmAstInfo(llMove)
                                    End If
                                Next llMove
                                ReDim tmAstInfo(0 To UBound(tmAstAdj1)) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    tmAstInfo(llMove) = tmAstAdj1(llMove)
                                Next llMove
                                Exit For
                            ElseIf (llTime > llTestTime2) And (((llMerge + 1 = ilLastSpotOfDay) And (llLimit = 2)) Or ((llMerge = ilLastSpotOfDay) And (llLimit = 1))) Then
                                ReDim tmAstAdj1(0 To UBound(tmAstInfo) + 1) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    tmAstAdj1(llMove) = tmAstInfo(llMove)
                                Next llMove
                                tmMergeAstInfo(llAst).lCode = Abs(tmMergeAstInfo(llAst).lCode)
                                tmAstAdj1(UBound(tmAstInfo)) = tmMergeAstInfo(llAst)
                                tmMergeAstInfo(llAst).lCode = -tmMergeAstInfo(llAst).lCode
                                ReDim tmAstInfo(0 To UBound(tmAstAdj1)) As ASTINFO
                                For llMove = 0 To UBound(tmAstInfo) Step 1
                                    tmAstInfo(llMove) = tmAstAdj1(llMove)
                                Next llMove
                                Exit For
                            End If
                        End If
                    End If
                Next llMerge
            End If
        Next llDayLoop
    Next llAst
End Sub
Private Function mAgreementAdjustGamesCue(ilVef As Integer) As Boolean
    Dim myRs As ADODB.Recordset
    Dim slSql As String
    Dim slStart As String
    Dim slEnd As String
    Dim blRet As Boolean
    
On Error GoTo ERRORBOX:
    blRet = True
    If ilVef > 0 Then
        ' is this vehicle for games...marked with 'event'?
        slSql = "select vffCode from VFF_Vehicle_Features where vffxdxmlForm in ('A','S') AND UCASE(vffXDProgCodeId) = 'EVENT' and vffvefcode = " & ilVef
        Set myRs = gSQLSelectCall(slSql)
        If Not myRs.EOF Then
            '6199. No Don't do this part
            'Yes?  Set to not send any authorizations
            slStart = "'" & Format$(smDate, sgSQLDateForm) & "'"
            slEnd = "'" & Format$(DateAdd("d", imNumberDays - 1, smDate), sgSQLDateForm) & "'"
'            slSql = "UPDATE att set attSentToXDSStatus = 'Y' WHERE attVefCode = " & ilVef
'            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
'                GoSub ERRORBOX:
'            End If
            ' Now set to send authorizations for this time period.
            slSql = "update att set attsenttoxdsstatus = 'M' where attVefCode = " & ilVef & " and attOnAir <= "
            slSql = slSql & slEnd & " AND attOffAir >= " & slStart & " and attDropDate >=" & slStart
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ERRORBOX:
                myExport.WriteError "mAgreementAdjustGamesCue.  Error: " & Err.Description, True, False
                blRet = False
                GoTo Cleanup
            End If
        End If
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    mAgreementAdjustGamesCue = blRet
    Exit Function
ERRORBOX:
    'gLogMsg "Error in mAgreementAdjustGamesCue.  Error: " & Err.Description, smPathForgLogMsg, False
    myExport.WriteError "mAgreementAdjustGamesCue.  Error: " & Err.Description, True, False
    blRet = False
    GoTo Cleanup
End Function
'Private Function mAgreementAdjustGamesISCI(ilVef As Integer) As Boolean
'    Dim myRs As ADODB.Recordset
'    Dim slSql As String
'    Dim slStart As String
'    Dim slEnd As String
'    Dim blRet As Boolean
'
'On Error GoTo ERRORBOX:
'    blRet = True
'    slStart = "'" & Format$(smDate, sgSQLDateForm) & "'"
'    slEnd = "'" & Format$(DateAdd("d", imNumberDays - 1, smDate), sgSQLDateForm) & "'"
'    ' Now set to send authorizations for this time period.
'    'I think it's ok not to have 'P' - is isci, because we tested the interface id
'    slSql = "update att set attsenttoxdsstatus = 'M' where attVefCode = " & ilVef & " and attOnAir <= "
'    slSql = slSql & slEnd & " AND attOffAir >= " & slStart & " and attDropDate >=" & slStart
'    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
'        GoSub ERRORBOX:
'    End If
'cleanup:
'    If Not myRs Is Nothing Then
'        If (myRs.State And adStateOpen) <> 0 Then
'            myRs.Close
'        End If
'        Set myRs = Nothing
'    End If
'    mAgreementAdjustGamesISCI = blRet
'    Exit Function
'ERRORBOX:
'    myExport.WriteError "mAgreementAdjustGamesISCI.  Error: " & Err.Description, True, False
'    blRet = False
'    GoTo cleanup
'End Function
Private Function mCalcCRC32(ByteArray() As Byte) As Long
Dim i As Long
Dim j As Long
Dim Limit As Long
Dim CRC As Long
Dim Temp1 As Long
Dim Temp2 As Long
Dim CRCTable(0 To 255) As Long
  
  Limit = &HEDB88320
  For i = 0 To 255
    CRC = i
    For j = 8 To 1 Step -1
      If CRC < 0 Then
        Temp1 = CRC And &H7FFFFFFF
        Temp1 = Temp1 \ 2
        Temp1 = Temp1 Or &H40000000
      Else
        Temp1 = CRC \ 2
      End If
      If CRC And 1 Then
        CRC = Temp1 Xor Limit
      Else
        CRC = Temp1
      End If
    Next j
    CRCTable(i) = CRC
  Next i
  Limit = UBound(ByteArray)
  CRC = -1
  For i = 0 To Limit
    If CRC < 0 Then
      Temp1 = CRC And &H7FFFFFFF
      Temp1 = Temp1 \ 256
      Temp1 = (Temp1 Or &H800000) And &HFFFFFF
    Else
      Temp1 = (CRC \ 256) And &HFFFFFF
    End If
    Temp2 = ByteArray(i)   ' get the byte
    Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
    CRC = Temp1 Xor Temp2
  Next i
  CRC = CRC Xor &HFFFFFFFF
  mCalcCRC32 = CRC
End Function


Private Sub mGetRotDTForGames(slProgCodeID As String, llGsfCode As Long, slRotStartDT As String, slRotEndDT As String)
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim llTime As Long
    Dim llDate As Long
    Dim ilVpf As Integer
    Dim slEndDate As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    If UCase(slProgCodeID) <> "EVENT" Then
        Exit Sub
    End If
    llVef = gBinarySearchVef(CLng(imVefCode))
    If llVef <> -1 Then
        If tgVehicleInfo(llVef).sVehType <> "G" Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfCode = " & llGsfCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        Exit Sub
    End If
    If UBound(tmGameTimeRange) <= LBound(tmGameTimeRange) Then
        ilVpf = gBinarySearchVpf(CLng(imVefCode))
        If ilVpf <> -1 Then
            If (Asc(tgVpfOptions(ilVpf).sUsingFeatures2) And XDSAPPLYMERGE) = XDSAPPLYMERGE Then
                Exit Sub
            End If
        End If
        slEndDate = DateAdd("d", imNumberDays - 1, smDate)
        ilRet = gGetProgramTimes(imVefCode, smDate, slEndDate, tmGameTimeRange(), False)
    End If
    For ilLoop = 0 To UBound(tmGameTimeRange) - 1 Step 1
        If tmGameTimeRange(ilLoop).iGameNo = rst!gsfGameNo Then
            llDate = tmGameTimeRange(ilLoop).lDate
            '7/12/13: Adjust start time by one hour
            llTime = tmGameTimeRange(ilLoop).lStartTime - 3600
            If llTime < 0 Then
                llTime = llTime + 86400
                llDate = llDate - 1
            End If
            slRotStartDT = Format(llDate, "m/d/yy") & " " & Format(gLongToTime(llTime), "hh:mm:ss")
            llDate = tmGameTimeRange(ilLoop).lDate
            '7/12/13:  Adjust end time by five hours
            llTime = tmGameTimeRange(ilLoop).lEndTime + 5 * 3600
            If llTime > 86400 Then
                llTime = llTime - 86400
                llDate = llDate + 1
            End If
            If llTime = 86400 Then
                llTime = llTime - 1
            End If
            slRotEndDT = Format(llDate, "m/d/yy") & " " & Format(gLongToTime(llTime), "hh:mm:ss")
            Exit Sub
        End If
    Next ilLoop
    Exit Sub
ErrHandler:
    gHandleError smPathForgLogMsg, "frmExportXDigital-mGetRotDTForGames"
    Exit Sub
End Sub
Private Function mSendBasic(blHaltOnError As Boolean, blIsResend As Boolean, slRoutine As String, slStatus As String) As Boolean
    'return slStatus for parsing of vehicles and agreements
    Dim ilRet As Integer
    Dim blRet As Boolean
   ' Dim slStatus As String
    Dim myErrorText As TextStream
    
    blRet = True
    bmIsError = False
    'Dan M 6/27/14 this isn't needed
    'Set myFile = New FileSystemObject
    ' delete Jeff's error file. Write, then test for file.  If it exists, we got 'error/warning'
    'bmTestError helps me test when there is an error.  from menu/tools
    If myFile.FILEEXISTS(smXmlErrorFile) And bmTestError = False Then
On Error GoTo ERRORNODELETE
        myFile.DeleteFile smXmlErrorFile, True
    End If
    If blIsResend Then
        ilRet = csiXMLResend(1)
    Else
        ilRet = csiXMLWrite(1)
    End If
    DoEvents
     ' 6581 reading errors.  Same as ilret <> true
     If myFile.FILEEXISTS(smXmlErrorFile) Then
        blRet = False
On Error GoTo ERRORNOOPEN
        Set myErrorText = myFile.OpenTextFile(smXmlErrorFile, ForReading, False)
        slStatus = myErrorText.ReadAll
        myErrorText.Close
        'Jeff now returns "" around attributes  <msg code = "-2" strip this
        slStatus = Replace(slStatus, """", "")
        'warning... if it's an error, bmIsError is set to true and error info written out.  Just need to stop export.
        If Not mIsXmlError(slStatus, slRoutine, blHaltOnError) Then
           ' gLogMsg "Warnings - a call in " & slRoutine & " not accepted by XDS: " & slStatus, smPathForgLogMsg, False
            myExport.WriteWarning "A call in " & slRoutine & " not accepted by XDS: " & slStatus
'            '7236 ignore mapping issue
            If mIgnoreWarning(slStatus) Then
                blRet = True
                myExport.WriteFacts "Warning above will be ignored"
            End If
        End If
        myFile.DeleteFile smXmlErrorFile
    End If
Cleanup:
    mSendBasic = blRet
    Set myErrorText = Nothing
    Exit Function
ERRORNOOPEN:
    blRet = False
    myExport.WriteError "Warning- could not read xml error file " & Err.Description, True, False
    GoTo Cleanup
    Exit Function
ERRORNODELETE:
    myExport.WriteError "Warning- could not delete xml error file " & Err.Description, True, False
    Resume Next
    
End Function
'Private Function mSendAndWriteReturn(slRoutine As String) As Boolean
''6635
''6966 large rewrite to handle retries
'    'return false if warning or error
'    '6966 but still true even if errors on first try but successful on resend.
'    Dim ilRet As Integer
'    Dim blRet As Boolean
'    Dim c As Integer
'    Dim slRet As String
'    Dim slStatus As String
'    Dim slRoutine As String
'    Dim blReturnIds As Boolean
'    Dim blReturnSiteIds As Boolean
'
'    bmAllowXMLCommands = True
'    slStatus = ""
'    blRet = True
'    If Not mSendBasic(False, False, slRoutine, slStatus) Then
'        If bmIsError Then
'            mSetResults "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted.", MESSAGERED
'            myExport.WriteWarning "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted."
'            For c = 1 To imMaxRetries - 1
'                If mSendBasic(False, True, slRoutine, slStatus) Then
'                    Exit For
'                ElseIf bmIsError = False Then
'                    blRet = False
'                    Exit For
'                End If
'            Next c
'            If bmIsError Then
'                blRet = mSendBasic(True, True, slRoutine, slStatus)
'            End If
'            'resending fixed the issue
'            If bmIsError = False Then
'                bmAlertAboutReExport = True
'                mSetResults "Error in sending " & slRoutine & " corrected. Export Ok and continuing.", MESSAGERED
'                myExport.WriteWarning "Error in sending " & slRoutine & " corrected. Export Ok and continuing."
'            End If
'        Else
'            blRet = False
'        End If
'    Else
'
'    End If
'    mSendAndWriteReturn = blRet
'    Exit Function
'End Function
'Private Function mSendAndWriteReturn(slRoutine As String) As Boolean
''6635
'    'return false if warning or error
'    Dim ilRet As Integer
'    Dim blRet As Boolean
'    Dim slStatus As String
'    Dim myErrorText As TextStream
'
'    If (slRoutine = "Spot Insertions") Then
'        '6/24/14: Remove records not found in current export
'        '         Placed here because HB might send the same spot unchanged
'        '         as another spot in the break changed.
'        ilRet = mSendDeleteCommands()
'    End If
'
'    Set myFile = New FileSystemObject
'    blRet = True
'    ' delete Jeff's error file. Write, then test for file.  If it exists, we got 'error/warning'
'    'bmTestError helps me test when there is an error.  from menu/tools
'    If myFile.FileExists(smXmlErrorFile) And bmTestError = False Then
'    'If myFile.FileExists(smXmlErrorFile) Then
'On Error GoTo ERRORNODELETE
'        myFile.DeleteFile smXmlErrorFile, True
'    End If
'    ilRet = csiXMLWrite(1)
'    DoEvents
'     ' 6581 reading errors.  Same as ilret <> true
'     If myFile.FileExists(smXmlErrorFile) Then
'        blRet = False
'        '5896 now cancel on error, still continue on warning
'On Error GoTo ERRORNOOPEN
'        Set myErrorText = myFile.OpenTextFile(smXmlErrorFile, ForReading, False)
'        slStatus = myErrorText.ReadAll
'        myErrorText.Close
'        'Jeff now returns "" around attributes  <msg code = "-2" strip this
'        slStatus = Replace(slStatus, """", "")
'        'warning... if it's an error, bmIsError is set to true and error info written out.  Just need to stop export.
'        If Not mIsXmlError(slStatus, slRoutine) Then
'           ' gLogMsg "Warnings - a call in " & slRoutine & " not accepted by XDS: " & slStatus, smPathForgLogMsg, False
'            myExport.WriteWarning "A call in " & slRoutine & " not accepted by XDS: " & slStatus
'        End If
'        myFile.DeleteFile smXmlErrorFile
'    End If
'    If (blRet) And (slRoutine = "Spot Insertions") Then
'        '6/24/14: Add current export records
'        ilRet = mUpdateXHT()
'    End If
'cleanup:
'    mSendAndWriteReturn = blRet
'    Set myErrorText = Nothing
'    Exit Function
'ERRORNOOPEN:
'    blRet = False
'    'gLogMsg "Warning- could not read xml error file " & Err.Description, smPathForgLogMsg, False
'    myExport.WriteError "Warning- could not read xml error file " & Err.Description, True, False
'    GoTo cleanup
'    Exit Function
'ERRORNODELETE:
'    'gLogMsg "Warning- could not delete xml error file " & Err.Description, smPathForgLogMsg, False
'    myExport.WriteError "Warning- could not delete xml error file " & Err.Description, True, False
'    Resume Next
'End Function
'Private Function mAgreementSendAndTest(slDoNotReturn As String) As Boolean
'    'return false if warning or error
'    '6966 large rewrite
'    '6966 but still true even if errors on first try but successful on resend.
'    Dim blRet As Boolean
'    Dim c As Integer
'    Dim slRet As String
'    Dim slRoutine As String
'    Dim slStatus As String
'
'    slRoutine = "SetAuthorizations"
'    slStatus = ""
'    blRet = True
'    If Not mSendBasic(False, False, slRoutine, slStatus) Then
'        If bmIsError Then
'            mSetResults "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted.", MESSAGERED
'            myExport.WriteWarning "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted."
'            For c = 1 To imMaxRetries - 1
'                If mSendBasic(False, True, slRoutine, slStatus) Then
'                    Exit For
'                ElseIf bmIsError = False Then
'                    blRet = False
'                    Exit For
'                End If
'            Next c
'            If bmIsError Then
'                blRet = mSendBasic(True, True, slRoutine, slStatus)
'            End If
'            'resending fixed the issue
'            If bmIsError = False Then
'                bmAlertAboutReExport = True
'                mSetResults "Error in sending " & slRoutine & " corrected. Export Ok and continuing.", MESSAGERED
'                myExport.WriteWarning "Error in sending " & slRoutine & " corrected. Export Ok and continuing."
'                If blRet = False Then
'                    slDoNotReturn = mReturnIds(slStatus)
''                    'I log message even if slDoNotReturn has nothing in it...just in case there was an error in mReturnIds
''                    myExport.WriteWarning "Some Programs not accepted by XDS: " & slStatus
'                End If
'            End If
'        Else
'            blRet = False
'            slDoNotReturn = mReturnIds(slStatus)
'            'I log message even if slDoNotReturn has nothing in it...just in case there was an error in mReturnIds
'            myExport.WriteWarning "Some Programs not accepted by XDS: " & slStatus
'        End If
'    Else
'
'    End If
'    mAgreementSendAndTest = blRet
'    Exit Function
'End Function
'Private Function mAgreementSendAndTest(slDoNotReturn As String) As Boolean
''6581
'    'return false if warning or error
'    Dim ilRet As Integer
'    Dim blRet As Boolean
'    Dim slStatus As String
'    Dim myErrorText As TextStream
'
'    Set myFile = New FileSystemObject
'    blRet = True
'    ' delete Jeff's error file. Write, then test for file.  If it exists, we got 'error/warning'
'    'bmTestError helps me test when there is an error.  from menu/tools
'    If myFile.FileExists(smXmlErrorFile) And bmTestError = False Then
'   ' If myFile.FileExists(smXmlErrorFile) Then
'On Error GoTo ERRORNODELETE
'        myFile.DeleteFile smXmlErrorFile, True
'    End If
'    ilRet = csiXMLWrite(1)
'    DoEvents
'     ' 6581 reading errors.  Same as ilret <> true
'     If myFile.FileExists(smXmlErrorFile) Then
'        blRet = False
'        '5896 now cancel on error, still continue on warning
'On Error GoTo ERRORNOOPEN
'        Set myErrorText = myFile.OpenTextFile(smXmlErrorFile, ForReading, False)
'        slStatus = myErrorText.ReadAll
'        myErrorText.Close
'        'Jeff now returns "" around attributes  <msg code = "-2" strip this
'        slStatus = Replace(slStatus, """", "")
'        'warning
'        If Not mIsXmlError(slStatus, "SetAuthorizations") Then
'            slDoNotReturn = mReturnIds(slStatus)
'            'I log message even if slDoNotReturn has nothing in it...just in case there was an error in mReturnIds
'            'gLogMsg "Warnings - some authorizations not accepted by XDS: " & slStatus, smPathForgLogMsg, False
'            myExport.WriteWarning "Some authorizations not accepted by XDS: " & slStatus
'        End If
'        myFile.DeleteFile smXmlErrorFile
'    End If
'cleanup:
'    mAgreementSendAndTest = blRet
'    Set myErrorText = Nothing
'    Exit Function
'ERRORNOOPEN:
'    blRet = False
'    'gLogMsg "Warning- could not read xml error file " & Err.Description, smPathForgLogMsg, False
'    myExport.WriteError "Could not read xml error file " & Err.Description, True, False
'    GoTo cleanup
'    Exit Function
'ERRORNODELETE:
'    'gLogMsg "Warning- could not delete xml error file " & Err.Description, smPathForgLogMsg, False
'    myExport.WriteError "Could not delete xml error file " & Err.Description, True, False
'    Resume Next
'End Function
Private Function mAdjustUpdates(ByVal slNeedUpdate As String, slDoNotUpdate As String) As String
    'both come in with extra comma at end.
    Dim slSafe As String
    Dim slDont As String
    Dim slDonts() As String
    Dim c As Integer
    Dim slRemoveThis As String
    
    slSafe = slNeedUpdate
    slDont = mLoseLastLetter(slDoNotUpdate)
    If Len(slDont) > 0 Then
        slDonts = Split(slDont, ",")
        slSafe = "," & slSafe
        For c = 0 To UBound(slDonts)
            slRemoveThis = "," & slDonts(c) & ","
            slSafe = Replace(slSafe, slRemoveThis, ",")
        Next c
        'must start with "," or is empty
        If Len(slSafe) > 1 Then
            slSafe = mLoseLastLetter(slSafe)
            slSafe = Mid(slSafe, 2)
        Else
            slSafe = ""
        End If
    '6717
    Else
        slSafe = mLoseLastLetter(slSafe)
    End If
    mAdjustUpdates = slSafe
End Function
Private Function mReturnIds(slMessage As String) As String
    'return 7,11,41,
    'strip #s from <SetAuthorizationsResult TransmissionID=225 Count=5><msgs><msg code=-2 ID=7 ...etc.
    Dim ilPos As Long
    Dim ilEnd As Long
    Dim slRet As String
    Dim slTemp As String
    
    slRet = ""
    ilPos = 1
On Error GoTo errbox
    Do While ilPos > 0
        '-2 too restrictive.  Already tested if 'error', so ignore the code #
        ilPos = InStr(ilPos, slMessage, "<msg code=-")
       ' ilPos = InStr(ilPos, slMessage, "<msg code=-2")
        If ilPos > 0 Then
            ilPos = InStr(ilPos, slMessage, " ID=")
            If ilPos > 0 Then
                ilEnd = InStr(ilPos + 1, slMessage, " ")
                If ilEnd > ilPos Then
                    slRet = slRet & Mid(slMessage, ilPos + 4, ilEnd - ilPos - 4) & ","
                End If
            End If
        End If
    Loop
    '7508
    If Len(slRet) = 0 Then
        bmFailedToReadReturn = True
    End If
    mReturnIds = slRet
    Exit Function
errbox:
   ' gLogMsg "Error in mReturnIds, some agreements may have been marked as updated when they should not have been.", smPathForgLogMsg, False
    myExport.WriteError "Error in mReturnIds, some agreements may have been marked as updated when they should not have been.", True, False
    bmFailedToReadReturn = True
    mReturnIds = ""
End Function

Function mPathOfFile(slFile As String) As String
    Dim ilPos As Integer
    ilPos = InStrRev(slFile, "\")
    If ilPos > 0 Then
        'mPathOfFile = Left$(slFile, ilPos)
        mPathOfFile = Mid(slFile, 1, ilPos)
    Else
        mPathOfFile = ""
    End If
End Function

Private Function mInitPreFeedInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ToDate", adInteger
        .Append "TimeOffset", adInteger
        .Append "FromDate", adInteger
        .Append "FromStartTime", adInteger
        .Append "FromEndTime", adInteger
    End With
    rst.Open
    'rst!ToDate.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitPreFeedInfo = rst
End Function

Private Sub mClosePreFeedInfo()
    On Error Resume Next
    If Not PreFeedInfo_rst Is Nothing Then
        If (PreFeedInfo_rst.State And adStateOpen) <> 0 Then
            PreFeedInfo_rst.Close
        End If
        Set PreFeedInfo_rst = Nothing
    End If

End Sub


Private Function mBuildPreFeedInfo(ilVefCode As Integer) As Integer
    Dim ilPass As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim blAddDate As Boolean
    Dim blFound As Boolean
    Dim llTimeOffset As Long
    Dim llFromDate As Long
    Dim llFromStartTime As Long
    Dim llFromEndTime As Long
    Dim llToStartTime As Long
    Dim slSuDate As String
    
    On Error GoTo ErrHand
    slSuDate = gObtainNextSunday(smDate)
    mBuildPreFeedInfo = False
    mClosePreFeedInfo
    Set PreFeedInfo_rst = mInitPreFeedInfo()
    For ilPass = 0 To 1 Step 1
        ilDay = -1
        Do
            If ilDay = -1 Then
                ilDay = 0
            ElseIf ilDay = 0 Then
                ilDay = 6
            ElseIf ilDay = 6 Then
                ilDay = 7
            Else
                Exit Do
            End If
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM " & """PFF_Pre-Feed"""
            If ilPass = 0 Then
                SQLQuery = SQLQuery + " WHERE (pffType = " & "'D'"
            Else
                SQLQuery = SQLQuery + " WHERE (pffType = " & "'E'"
            End If
            SQLQuery = SQLQuery + " AND pffVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND pffAirDay = " & ilDay
            'SQLQuery = SQLQuery + " AND pffStartDate <= '" & Format(smDate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " AND pffStartDate <= '" & Format(slSuDate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " Order by pffStartDate Desc"
            Set Pff_rst = gSQLSelectCall(SQLQuery)
            Do While Not Pff_rst.EOF
                If imTerminate Then
                    Exit Function
                End If
                llSDate = gDateValue(gObtainPrevMonday(smDate)) + Pff_rst!pffToDay
                If llSDate < gDateValue(smDate) Then
                    llSDate = gDateValue(smDate)
                End If
                'llEDate = gDateValue(gObtainNextSunday(smDate))
                llEDate = gDateValue(DateAdd("d", imNumberDays - 1, smDate))
                For llDate = llSDate To llEDate Step 1
                    blAddDate = False
                    If (Pff_rst!pffFromDay <> gWeekDayLong(llDate)) Then
                        llFromDate = gDateValue(gObtainPrevMonday(smDate)) + Pff_rst!pffFromDay
                        llFromStartTime = gTimeToLong(Format$(Pff_rst!pffFromStartTime, "h:mm:ssam/pm"), False)
                        llFromEndTime = gTimeToLong(Format$(Pff_rst!pffFromEndTime, "h:mm:ssam/pm"), True)
                        llToStartTime = gTimeToLong(Format$(Pff_rst!pffToStartTime, "h:mm:ssam/pm"), False)
                        llTimeOffset = llToStartTime - llFromStartTime
                    Else
                        llTimeOffset = 0
                        llFromDate = llDate
                        llFromStartTime = 0
                        llFromEndTime = 86400
                    End If
                    'Look for duplicates
                    blFound = False
                    PreFeedInfo_rst.Filter = adFilterNone
                    Do While Not PreFeedInfo_rst.EOF
                        If llFromDate = PreFeedInfo_rst!FromDate Then
                            If (llFromStartTime = PreFeedInfo_rst!FromStartTime) And (llFromEndTime = PreFeedInfo_rst!FromEndTime) Then
                                If llDate = PreFeedInfo_rst!ToDate Then
                                    blFound = True
                                    Exit Do
                                End If
                            End If
                        End If
                        PreFeedInfo_rst.MoveNext
                    Loop
                    If Not blFound Then
                        blAddDate = True
                    End If
                    If blAddDate Then
                        PreFeedInfo_rst.AddNew Array("ToDate", "TimeOffset", "FromDate", "FromStartTime", "FromEndTime"), Array(llDate, llTimeOffset, llFromDate, llFromStartTime, llFromEndTime)
                    End If
                Next llDate
                Pff_rst.MoveNext
            Loop
        Loop
    Next ilPass
    mBuildPreFeedInfo = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export XDigital-mBuildPreFeedInfo"
    Resume Next
End Function

Private Sub mCreateAstPreFeedSpots()
    Dim llAst As Long
    Dim llUpper As Long
    Dim llAirTime As Long
    
    On Error GoTo ErrorHandle:
    
    If PreFeedInfo_rst.RecordCount <= 0 Then
        Exit Sub
    End If
    ReDim tmAstAdj2(LBound(tmAstInfo) To UBound(tmAstInfo)) As ASTINFO
    For llAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        tmAstAdj2(llAst) = tmAstInfo(llAst)
    Next llAst
    ReDim tmAstInfo(0 To 0) As ASTINFO
    llUpper = 0
    PreFeedInfo_rst.Sort = "ToDate,TimeOffset"
    PreFeedInfo_rst.MoveFirst
    Do While Not PreFeedInfo_rst.EOF
        For llAst = LBound(tmAstAdj2) To UBound(tmAstAdj2) - 1 Step 1
            If (gDateValue(tmAstAdj2(llAst).sFeedDate) = PreFeedInfo_rst!FromDate) Then
                llAirTime = gTimeToLong(tmAstAdj2(llAst).sFeedTime, False)
                If (llAirTime >= PreFeedInfo_rst!FromStartTime) And (llAirTime < PreFeedInfo_rst!FromEndTime) Then
                    tmAstInfo(llUpper) = tmAstAdj2(llAst)
                    tmAstInfo(llUpper).sFeedDate = Format(PreFeedInfo_rst!ToDate, "m/d/yy")
                    tmAstInfo(llUpper).sFeedTime = gFormatTimeLong(llAirTime + PreFeedInfo_rst!TimeOffset, "A", "1")
                    llUpper = llUpper + 1
                    ReDim Preserve tmAstInfo(0 To llUpper) As ASTINFO
                End If
            End If
        Next llAst
        PreFeedInfo_rst.MoveNext
    Loop
    Exit Sub
ErrorHandle:
    Resume Next
End Sub

Private Sub mRemoveExtraAirplays()
    Dim ilAst As Integer
    Dim ilCount As Integer
    ReDim tlAstInfo(0 To UBound(tmAstInfo)) As ASTINFO
    
    ilCount = 0
    For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
        If tmAstInfo(ilAst).iAirPlay <= 1 Then
            tlAstInfo(ilCount) = tmAstInfo(ilAst)
            ilCount = ilCount + 1
        End If
    Next ilAst
    If ilCount < UBound(tmAstInfo) Then
        ReDim tmAstInfo(0 To ilCount) As ASTINFO
        For ilAst = 0 To ilCount Step 1
            tmAstInfo(ilAst) = tlAstInfo(ilAst)
        Next ilAst
    End If
End Sub
 Private Function mSafeChunkSize() As Integer
 '6882
    'reduce chunk by 20% for spot insertions
    'Const MAYSEND As Integer = 100
    Const PERCENTOF As Integer = 8
    Dim ilRet As Integer
    Dim llTemp As Long
    ' if I don't make a long, get an overflow error
    llTemp = CLng(imChunk) * PERCENTOF
    ilRet = llTemp \ 10
    If ilRet < 1 Then
        ilRet = imChunk
    End If
    mSafeChunkSize = ilRet
 End Function

Private Function mInitXHTInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    mCloseXHTInfo
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "xhtCode", adInteger
        .Append "attCode", adInteger
        .Append "FeedDate", adInteger
        .Append "SiteID", adInteger
        .Append "TransmissionID", adChar, 20
        .Append "UnitID", adChar, 20
        .Append "ISCI", adChar, 20
        .Append "ProgCodeID", adChar, 8
        .Append "Type", adChar, 1 'Pass 0 (ISCI): G=Generic; R=Regional. Pass 1(HBP) or Pass 2(HB): G=General only.  R is not required as either Generic or Region is send not both
        .Append "Status", adChar, 1 'U=Unchanged; D=Delete; R=Replaced; N=New
        '7509
        .Append "GsfCode", adInteger
    End With
    rst.Open
    'rst!ToDate.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitXHTInfo = rst
End Function
Private Function mBuildXHTInfo(llAttCode As Long, slFeedStartDate As String, slFeedEndDate As String, ilPass As Integer) As Integer
    '7256 return is true if xht was built..that is, is a re-export
    Dim llDate As Long
    Dim slType As String
    Dim slPrevTransmissionID As String
    Dim slPrevSiteId As String
    Dim slPrevUnitId As String
    Dim slPrevProgCodeId As String
    Dim blFound As Boolean
    Dim ilProgCode As Integer
    Dim slProp As String
    
    On Error GoTo ErrHand
    
    mBuildXHTInfo = False
    'dan 4/13/15  This looks like it blocks games from merging.  But ok if not games ( or not merging)
    ReDim tmProgCodeMatch(0 To 0) As PROGCODEMATCH
    
    xhtInfo_rst.Filter = "attCode = " & llAttCode
    If Not xhtInfo_rst.EOF Then
        mBuildXHTInfo = True
        Exit Function
    End If
    slPrevTransmissionID = ""
    '7675 find N-unconfirmed.  National model may have to delete these if there is a regional and new doesn't have regional
    If ilPass = 0 Then
    'now delete!
         If Not mSendISCIDeletes(llAttCode, slFeedStartDate, slFeedEndDate) Then
                Call mSetResults("Issue with ISCI deletes. see XDigitalExportLog.Txt", MESSAGERED)
         End If
    End If
    If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
        '7236 delete unconfirmed before getting.
        SQLQuery = "Delete FROM xht"
        SQLQuery = SQLQuery & " WHERE xhtAttCode = " & llAttCode
        SQLQuery = SQLQuery & " AND xhtFeedDate >= '" & Format(slFeedStartDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND xhtFeedDate <= '" & Format(slFeedEndDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND (xhtStatus = 'N')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError smPathForgLogMsg, "Export XDigital-mBuildXHTInfo"
            mBuildXHTInfo = False
            Exit Function
        End If
        'update "D" to ""
        SQLQuery = "Update xht set xhtStatus = '' "
        SQLQuery = SQLQuery & " WHERE xhtAttCode = " & llAttCode
        SQLQuery = SQLQuery & " AND xhtFeedDate >= '" & Format(slFeedStartDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND xhtFeedDate <= '" & Format(slFeedEndDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND (xhtStatus = 'D')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError smPathForgLogMsg, "Export XDigital-mBuildXHTInfo"
            mBuildXHTInfo = False
            Exit Function
        End If
    End If
    SQLQuery = "SELECT * FROM xht"
    SQLQuery = SQLQuery & " WHERE xhtAttCode = " & llAttCode
    SQLQuery = SQLQuery & " AND xhtFeedDate >= '" & Format(slFeedStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND xhtFeedDate <= '" & Format(slFeedEndDate, sgSQLDateForm) & "'"
    'Dan M this wouldn't be needed if I deleted if generating file.
    SQLQuery = SQLQuery & " AND (xhtStatus <> 'D' AND xhtStatus <> 'N')"
    SQLQuery = SQLQuery & " Order By xhtCode"
    Set xht_rst = gSQLSelectCall(SQLQuery)
    Do While Not xht_rst.EOF
        blFound = False
        For ilProgCode = 0 To UBound(tmProgCodeMatch) - 1 Step 1
            If Trim(tmProgCodeMatch(ilProgCode).sProgCodeID) = Trim$(xht_rst!xhtProgCodeID) Then
                blFound = True
            End If
        Next ilProgCode
        If Not blFound Then
            ilProgCode = UBound(tmProgCodeMatch)
            tmProgCodeMatch(ilProgCode).sProgCodeID = Trim$(xht_rst!xhtProgCodeID)
            tmProgCodeMatch(ilProgCode).bMatch = False
            ReDim Preserve tmProgCodeMatch(0 To ilProgCode + 1) As PROGCODEMATCH
        End If
        llDate = gDateValue(xht_rst!xhtFeedDate)
        If (slPrevTransmissionID = Trim$(xht_rst!xhtTransmissionId)) And (slPrevSiteId = Trim$(xht_rst!xhtSiteId)) And (slPrevUnitId = Trim$(xht_rst!xhtunitid)) And (slPrevProgCodeId = Trim$(xht_rst!xhtProgCodeID)) Then
            slType = "R"
        Else
            slType = "G"
            If ilPass = 0 Then
                slPrevTransmissionID = Trim$(xht_rst!xhtTransmissionId)
                slPrevSiteId = Trim$(xht_rst!xhtSiteId)
                slPrevUnitId = Trim$(xht_rst!xhtunitid)
                slPrevProgCodeId = Trim$(xht_rst!xhtProgCodeID)
            End If
        End If
        '7509
        'xhtInfo_rst.AddNew Array("xhtCode", "attCode", "FeedDate", "SiteID", "TransmissionID", "UnitID", "ISCI", "ProgCodeID", "Type", "Status"), Array(xht_rst!xhtcode, llAttCode, llDate, xht_rst!xhtSiteID, xht_rst!xhtTransmissionId, xht_rst!xhtUnitID, xht_rst!xhtISCI, xht_rst!xhtProgCodeID, slType, "D")
       ' xhtInfo_rst.AddNew Array("xhtCode", "attCode", "FeedDate", "SiteID", "TransmissionID", "UnitID", "ISCI", "ProgCodeID", "Type", "Status", "GsfCode"), Array(xht_rst!xhtcode, llAttCode, llDate, xht_rst!xhtSiteId, xht_rst!xhtTransmissionId, xht_rst!xhtunitid, xht_rst!xhtISCI, xht_rst!xhtProgCodeID, slType, "D", xht_rst!xhtGsfCode)
        '7675 get the generics
        If ilPass <> 0 Or slType = "G" Then
            xhtInfo_rst.AddNew Array("xhtCode", "attCode", "FeedDate", "SiteID", "TransmissionID", "UnitID", "ISCI", "ProgCodeID", "Type", "Status"), Array(xht_rst!xhtcode, llAttCode, llDate, xht_rst!xhtSiteId, xht_rst!xhtTransmissionId, xht_rst!xhtunitid, xht_rst!xhtISCI, xht_rst!xhtProgCodeID, slType, "D")
        End If
       '7256
        mBuildXHTInfo = True
        xht_rst.MoveNext
    Loop
    '7256
    'mBuildXHTInfo = True
    Exit Function
'ErrHand1:
'    Screen.MousePointer = vbDefault
'    gHandleError smPathForgLogMsg, "Export XDigital-mBuildXHTInfo"
'    mBuildXHTInfo = False
'    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mBuildXHTInfo"
    Resume Next
End Function
Private Function mFindXHT(llSiteID As Long, slTransmissionID As String, slUnitID As String, slISCI As String, slProgCodeID As String, slType As String) As Integer
    Dim llDate As Long
    
    On Error GoTo ErrHand
    mFindXHT = False
    xhtInfo_rst.Filter = "SiteID = " & llSiteID & " And TransmissionID = '" & slTransmissionID & "' And UnitID = '" & slUnitID & "' And Type = '" & slType & "'"
    Do While Not xhtInfo_rst.EOF
        If (Trim$(slISCI) = Trim$(xhtInfo_rst!ISCI)) And (Trim$(slProgCodeID) = Trim$(xhtInfo_rst!ProgCodeId)) Then
            mFindXHT = True
            Exit Do
        End If
        xhtInfo_rst.MoveNext
    Loop
    xhtInfo_rst.Filter = adFilterNone
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mFindXHT"
    Resume Next
End Function
Private Function mSetXHT(llSiteID As Long, slTransmissionID As String, slUnitID As String, slISCI As String, slProgCodeID As String, slType As String, slSetValue As String) As Integer
    Dim llDate As Long
    Dim ilMatch As Integer
    Dim blFound As Boolean
    
    On Error GoTo ErrHand
    mSetXHT = False
    xhtInfo_rst.Filter = "SiteID = " & llSiteID & " And TransmissionID = '" & slTransmissionID & "' And UnitID = '" & slUnitID & "'" & " And Type = '" & slType & "'" & " And Status = 'D'"
    Do While Not xhtInfo_rst.EOF
        If ((Trim$(slISCI) = Trim$(xhtInfo_rst!ISCI) And (Trim$(slProgCodeID) = Trim$(xhtInfo_rst!ProgCodeId)))) Or (slSetValue = "R") Then
            'xhtInfo_rst!Status = slSetValue
            'mSetXHT = True
            'If slSetValue <> "R" Then
            '    Exit Do
            'End If
            If slSetValue <> "R" Then
                xhtInfo_rst!Status = slSetValue
                mSetXHT = True
                Exit Do
            End If
            If slProgCodeID = "" Then
                xhtInfo_rst!Status = slSetValue
                mSetXHT = True
            Else
                If Trim$(slProgCodeID) = Trim$(xhtInfo_rst!ProgCodeId) Then
                    xhtInfo_rst!Status = slSetValue
                    mSetXHT = True
                Else
                    blFound = False
                    For ilMatch = 0 To UBound(tmProgCodeMatch) - 1 Step 1
                        If (Trim$(tmProgCodeMatch(ilMatch).sProgCodeID) = Trim$(xhtInfo_rst!ProgCodeId)) And (tmProgCodeMatch(ilMatch).bMatch = True) Then
                            blFound = True
                            Exit For
                        End If
                    Next ilMatch
                    If Not blFound Then
                        For ilMatch = 0 To UBound(tmProgCodeMatch) - 1 Step 1
                            If (slProgCodeID <> Trim$(tmProgCodeMatch(ilMatch).sProgCodeID)) And (tmProgCodeMatch(ilMatch).bMatch = False) Then
                                xhtInfo_rst!Status = slSetValue
                                mSetXHT = True
                                Exit For
                            End If
                        Next ilMatch
                    End If
                End If
            End If
        End If
        xhtInfo_rst.MoveNext
    Loop
    xhtInfo_rst.Filter = adFilterNone
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mSetXHT"
    Resume Next
End Function
Private Function mAddXHT(llAttCode As Long, slFeedDate As String, llSiteID As Long, slTransmissionID As String, slUnitID As String, slISCI As String, slProgCodeID As String, slType As String, llGsfCode As Long) As Integer
    '7509 added gsfCode
    Dim llDate As Long
    
    On Error GoTo ErrHand
    llDate = gDateValue(slFeedDate)
    xhtInfo_rst.AddNew Array("xhtCode", "attCode", "FeedDate", "SiteID", "TransmissionID", "UnitID", "ISCI", "ProgCodeID", "Type", "Status", "GsfCode"), Array(0, llAttCode, llDate, llSiteID, slTransmissionID, slUnitID, slISCI, slProgCodeID, slType, "N", llGsfCode)
    mAddXHT = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mAddXHT"
    Resume Next
End Function
Private Sub mCloseXHTInfo()
    On Error Resume Next
    If Not xhtInfo_rst Is Nothing Then
        If (xhtInfo_rst.State And adStateOpen) <> 0 Then
            xhtInfo_rst.Close
        End If
        Set xhtInfo_rst = Nothing
    End If
End Sub

Private Function mUpdateXHT() As String
    '7236 return attcode if a change was made
    Dim llXhtCode As Long
    Dim ilXhttRecLen As Integer
    Dim ilRet As Integer
    Dim slRet As String
    Dim slAtts As String
    Dim blUpdate As Boolean
    Dim slProp As String
    
    slAtts = ","
    slRet = ""
    On Error GoTo ErrHand
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked Then
        mUpdateXHT = slRet
        Exit Function
    End If
    blUpdate = True
    '7508 move below
'    'Dan don't update if generating file and not overridden in menu  (1) is send to file
    If udcCriteria.XGenType(1, slProp) And bmTestForceUpdateXHT = False Then
        blUpdate = False
        'mUpdateXHT = slRet
        'Exit Function
    End If
    xhtInfo_rst.Filter = "Status <> " & "'U'"
    Do While Not xhtInfo_rst.EOF
        If InStr(1, slAtts, "," & xhtInfo_rst!attCode & ",") = 0 Then
            slAtts = slAtts & xhtInfo_rst!attCode & ","
        End If
        '7508 Dan I want to write out the atts for testing
        If blUpdate Then
            If (xhtInfo_rst!Status = "D") Or (xhtInfo_rst!Status = "R") Then
                '7236 don't remove: mark to remove
                SQLQuery = "UPDATE xht set xhtstatus = 'D' where xhtCode = " & xhtInfo_rst!xhtcode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand1:
                    Screen.MousePointer = vbDefault
                    gHandleError smPathForgLogMsg, "Export XDigital-mUpdateXHT"
                End If
    '            'Remove XHT record
    '            SQLQuery = "DELETE FROM xht where xhtCode = " & xhtInfo_rst!xhtcode
    '            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '                GoSub ErrHand1:
    '            End If
            ElseIf xhtInfo_rst!Status = "N" Then
                'Add new XHT record
    '            SQLQuery = "Insert Into xht ( "
    '            SQLQuery = SQLQuery & "xhtCode, "
    '            SQLQuery = SQLQuery & "xhtAttCode, "
    '            SQLQuery = SQLQuery & "xhtFeedDate, "
    '            SQLQuery = SQLQuery & "xhtSiteID, "
    '            SQLQuery = SQLQuery & "xhtTransmissionID, "
    '            SQLQuery = SQLQuery & "xhtUnitID, "
    '            SQLQuery = SQLQuery & "xhtISCI, "
    '            SQLQuery = SQLQuery & "xhtProgCodeID, "
    '            SQLQuery = SQLQuery & "xhtUnused "
    '            SQLQuery = SQLQuery & ") "
    '            SQLQuery = SQLQuery & "Values ( "
    '            SQLQuery = SQLQuery & "Replace" & ", "
    '            SQLQuery = SQLQuery & xhtInfo_rst!attCode & ", "
    '            SQLQuery = SQLQuery & "'" & Format$(xhtInfo_rst!FeedDate, sgSQLDateForm) & "', "
    '            SQLQuery = SQLQuery & xhtInfo_rst!SiteId & ", "
    '            SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(xhtInfo_rst!TransmissionId)) & "', "
    '            SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(xhtInfo_rst!UnitId)) & "', "
    '            SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(xhtInfo_rst!ISCI)) & "', "
    '            SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(xhtInfo_rst!ProgCodeID)) & "', "
    '            SQLQuery = SQLQuery & "'" & "" & "' "
    '            SQLQuery = SQLQuery & ") "
    '            llXhtCode = gInsertAndReturnCode(SQLQuery, "xht", "xhtCode", "Replace")
    '            If llXhtCode <= 0 Then
    '                GoSub ErrHand1:
    '            End If
                tmXht.lCode = 0
                tmXht.lAttCode = xhtInfo_rst!attCode
                gPackDate Format$(xhtInfo_rst!FeedDate, sgShowDateForm), tmXht.iFeedDate(0), tmXht.iFeedDate(1)
                tmXht.lSiteID = xhtInfo_rst!SiteID
                tmXht.sTransmissionID = Trim$(gFixQuote(xhtInfo_rst!TransmissionID))
                tmXht.sUnitID = Trim$(gFixQuote(xhtInfo_rst!UnitID))
                tmXht.sISCI = Trim$(gFixQuote(xhtInfo_rst!ISCI))
                tmXht.sProgCodeID = Trim$(gFixQuote(xhtInfo_rst!ProgCodeId))
                '7236
                tmXht.sStatus = "N"
                '7509
                tmXht.lgsfCode = xhtInfo_rst!gsfCode
                tmXht.sUnused = ""
                ilXhttRecLen = Len(tmXht)
                ilRet = btrInsert(hmXht, tmXht, ilXhttRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand1:
                    Screen.MousePointer = vbDefault
                    gHandleError smPathForgLogMsg, "Export XDigital-mUpdateXHT"
                End If
            End If
        End If
        xhtInfo_rst.MoveNext
    Loop
    xhtInfo_rst.Filter = adFilterNone
'    '6/24/14: Re-Initialize previously sent record info
    Set xhtInfo_rst = mInitXHTInfo()
    'don't accept ","
    If Len(slAtts) > 1 Then
        'change ",14,13,12," to "14,13,12,"
        slRet = Mid(slAtts, 2)
    End If
    mUpdateXHT = slRet
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mUpdateXHT"
    Resume Next
ErrHand1:
    gHandleError smPathForgLogMsg, "Export XDigital-mUpdateXHT"
    Return
End Function

Private Function mSendDeleteCommands() As Integer
    Dim slUnitID As String
    Dim slSiteId As String
    Dim slTransmissionID As String
    Dim ilPos As Integer
    Dim blIsDelete As Boolean
    Dim llLoop As Long
    
    bmAllowXMLCommands = True
    blIsDelete = False
    On Error GoTo ErrHand
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked Then
        mSendDeleteCommands = True
        Exit Function
    End If
    'xhtInfo_rst.Filter = "Status = " & "'D'"
    'If Not xhtInfo_rst.EOF Then
    If UBound(tmRetainDeletions) > LBound(tmRetainDeletions) Then
        blIsDelete = True
        mCSIXMLData "OT", "Deletes", ""
    End If
    'Do While Not xhtInfo_rst.EOF
    For llLoop = LBound(tmRetainDeletions) To UBound(tmRetainDeletions) - 1 Step 1
        lmReExportDelete = lmReExportDelete + 1
        'slSiteId = Trim$(xhtInfo_rst!SiteId)
        'slTransmissionID = Trim$(xhtInfo_rst!TransmissionID)
        'slUnitID = Trim$(xhtInfo_rst!UnitID)
        slSiteId = Trim$(tmRetainDeletions(llLoop).sSiteId)
        slTransmissionID = Trim$(tmRetainDeletions(llLoop).sTransmissionID)
        slUnitID = Trim$(tmRetainDeletions(llLoop).sUnitID)
        ilPos = InStr(1, slUnitID, "-", vbTextCompare)
        If ilPos > 0 Then
            slUnitID = Left(slUnitID, ilPos - 1)
        End If
        '7188 SiteID becomes SiteId
       ' mCSIXMLData "CA", "Delete", "UnitID=" & gAddQuotes(slUnitID) & " SiteID=" & gAddQuotes(slSiteId) & " TransmissionID=" & gAddQuotes(slTransmissionID)
        mCSIXMLData "CA", "Delete", "UnitID=" & gAddQuotes(slUnitID) & " SiteId=" & gAddQuotes(slSiteId) & " TransmissionID=" & gAddQuotes(slTransmissionID)
        'xhtInfo_rst.MoveNext
    'Loop
    Next llLoop
    If blIsDelete Then
        mCSIXMLData "CT", "Deletes", ""
    End If
    'xhtInfo_rst.Filter = adFilterNone
    ReDim tmRetainDeletions(0 To 0) As REATINDELETIONS
    mSendDeleteCommands = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mSendDeleteCommands"
    Resume Next
End Function
'9114 pass UnitId as its own, rather than trans,seq,vefcode
'Private Function mSetXMLCommandPass0(llIndexLoop As Long, slStationID As String, slTransmissionID As String, slSeqNo As String, slVefCode5 As String, slProgCodeID As String) As Boolean
'    Dim blExportSpot As Boolean
'    Dim ilRegionExist As Integer
'    Dim ilRet As Integer
'
'    blExportSpot = False
'    If tmAstInfo(llIndexLoop).iRegionType > 0 Then
'        ilRegionExist = True
'    Else
'        ilRegionExist = False
'    End If
'    If mFindXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G") Then
'        If ilRegionExist Then
'            If Not mFindXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R") Then
'                blExportSpot = True
'            End If
'        End If
'    Else
'        blExportSpot = True
'    End If
'    If Not blExportSpot Then
'        ilRet = mSetXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", "U")
'        If ilRegionExist Then
'            ilRet = mSetXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R", "U")
'        End If
'    Else
'        ilRet = mSetXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, "", Trim$(slProgCodeID), "G", "R")
'        '7509 added gsfcode
'        ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
'        If ilRegionExist Then
'            ilRet = mSetXHT(Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, "", Trim$(slProgCodeID), "R", "R")
'            '7509 added gsfcode
'            ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slStationID), slTransmissionID, slTransmissionID & slSeqNo & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R", tmAstInfo(llIndexLoop).lgsfCode)
'        End If
'    End If
'    mSetXMLCommandPass0 = blExportSpot
'
'End Function
Private Function mSetXMLCommandPass0(llIndexLoop As Long, slStationID As String, slTransmissionID As String, slUnitID As String, slProgCodeID As String) As Boolean
    Dim blExportSpot As Boolean
    Dim ilRegionExist As Integer
    Dim ilRet As Integer
    
    blExportSpot = False
    If tmAstInfo(llIndexLoop).iRegionType > 0 Then
        ilRegionExist = True
    Else
        ilRegionExist = False
    End If
    If mFindXHT(Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G") Then
        If ilRegionExist Then
            If Not mFindXHT(Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R") Then
                blExportSpot = True
            End If
        End If
    Else
        blExportSpot = True
    End If
    If Not blExportSpot Then
        ilRet = mSetXHT(Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", "U")
        If ilRegionExist Then
            ilRet = mSetXHT(Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R", "U")
        End If
    Else
        ilRet = mSetXHT(Val(slStationID), slTransmissionID, slUnitID, "", Trim$(slProgCodeID), "G", "R")
        '7509 added gsfcode
        ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
        If ilRegionExist Then
            ilRet = mSetXHT(Val(slStationID), slTransmissionID, slUnitID, "", Trim$(slProgCodeID), "R", "R")
            '7509 added gsfcode
            ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slStationID), slTransmissionID, slUnitID, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "R", tmAstInfo(llIndexLoop).lgsfCode)
        End If
    End If
    mSetXMLCommandPass0 = blExportSpot

End Function
'Private Function mSetXMLCommandPass1(llIndexLoop As Long, slXDReceiverID As String, slTransmissionID As String, slUnitHBP As String, slVefCode5 As String, slProgCodeID As String)
'    Dim blExportSpot As Boolean
'    Dim ilRegionExist As Integer
'    Dim ilRet As Integer
'
'    blExportSpot = False
'    If tmAstInfo(llIndexLoop).iRegionType > 0 Then
'        ilRegionExist = True
'    Else
'        ilRegionExist = False
'    End If
'    If Not ilRegionExist And udcCriteria.XSpots(0) Then
'        If smUnitIdByAstCodeForBreak = "Y" Then
'            If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G") Then
'                blExportSpot = True
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), "", Trim$(slProgCodeID), "G", "R")
'                '7509 added gsfcode
'                ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
'            Else
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", "U")
'            End If
'        Else
'            If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G") Then
'                blExportSpot = True
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, "", Trim$(slProgCodeID), "G", "R")
'                '7509 added gsfcode
'                ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
'            Else
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sISCI), Trim$(slProgCodeID), "G", "U")
'            End If
'        End If
'    ElseIf ilRegionExist Then
'        If smUnitIdByAstCodeForBreak = "Y" Then
'            If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G") Then
'                blExportSpot = True
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), "", Trim$(slProgCodeID), "G", "R")
'                '7509 added gsfcode
'                ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
'            Else
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llIndexLoop).lCode)), Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G", "U")
'            End If
'        Else
'            If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G") Then
'                blExportSpot = True
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, "", Trim$(slProgCodeID), "G", "R")
'                '7509 added gsfcode
'                ilRet = mAddXHT(tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G", tmAstInfo(llIndexLoop).lgsfCode)
'            Else
'                ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Mid(slTransmissionID, 3) & slUnitHBP & slVefCode5, Trim$(tmAstInfo(llIndexLoop).sRISCI), Trim$(slProgCodeID), "G", "U")
'            End If
'        End If
'    End If
'    mSetXMLCommandPass1 = blExportSpot
'End Function
'10021
Private Function mSetXMLCommandPassHBP(llIndexLoop As Long, slXDReceiverID As String, slTransmissionID As String, slUnitHBP As String, slVefCode5 As String, slProgCodeID As String) As Boolean
    Dim blExportSpot As Boolean
    Dim blRegionExist As Boolean
    Dim slUnitID As String
    Dim slIsciToUse As String
    
    blExportSpot = False
    If tmAstInfo(llIndexLoop).iRegionType > 0 Then
        blRegionExist = True
        slIsciToUse = Trim$(tmAstInfo(llIndexLoop).sRISCI)
    Else
        blRegionExist = False
        slIsciToUse = Trim$(tmAstInfo(llIndexLoop).sISCI)
        If udcCriteria.XSpots(REGIONALONLY) Then
            'generic, but sending regional only? skip
            GoTo finish
        End If
    End If
    slProgCodeID = Trim$(slProgCodeID)
    slUnitID = mCreateUnitIDForCue(HBPFORM, slUnitHBP & slVefCode5, tmAstInfo(llIndexLoop).lAttCode, slTransmissionID)
    If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, slUnitID, slIsciToUse, slProgCodeID, "G") Then
        blExportSpot = True
        mSetXHT Val(slXDReceiverID), slTransmissionID, slUnitID, "", slProgCodeID, "G", "R"
        '7509 added gsfcode
        mAddXHT tmAstInfo(llIndexLoop).lAttCode, tmAstInfo(llIndexLoop).sFeedDate, Val(slXDReceiverID), slTransmissionID, slUnitID, slIsciToUse, slProgCodeID, "G", tmAstInfo(llIndexLoop).lgsfCode
    Else
        mSetXHT Val(slXDReceiverID), slTransmissionID, slUnitID, slIsciToUse, slProgCodeID, "G", "U"
    End If
finish:
    mSetXMLCommandPassHBP = blExportSpot
End Function

Private Function mSetXMLCommandPassHB(llIndexStart As Long, llIndexEnd As Long, slXDReceiverID As String, slTransmissionID As String, slUnitHB As String, slVefCode5 As String, slProgCodeID As String) As Boolean
    '10021 copied mSetXMLCommandPass2 added type returned
    Dim blExportSpot As Boolean
    Dim llTest As Long
    Dim ilSpotNo As Integer
    Dim slXHTUnitID As String
   ' Dim ilRegionExist As Integer
    Dim ilRet As Integer
    Dim blFound As Boolean
    '9113 not really for 9113, but helped clean up code
    Dim slISCI As String
    Dim slType As String
    
    blExportSpot = False
    ilSpotNo = 0
    'Dan this first loop is if any spots were already sent, set to remove spots from deletion list(?)
    For llTest = llIndexStart To llIndexEnd Step 1
        ilSpotNo = ilSpotNo + 1
        '10021
        slXHTUnitID = mCreateUnitIDForCue(HBFORM, slUnitHB & slVefCode5, tmAstInfo(llTest).lCode, slTransmissionID)
        If smUnitIdByAstCodeForBreak <> "Y" Then
            If ilSpotNo <= 9 Then
                slXHTUnitID = slXHTUnitID & "-" & Trim$(Str$(ilSpotNo))
            Else
                slXHTUnitID = slXHTUnitID & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
            End If
        End If
'        If smUnitIdByAstCodeForBreak <> "Y" Then
'            'slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'            If ilSpotNo <= 9 Then
'                slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'            Else
'                slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
'            End If
'        Else
'            '9113
'            slXHTUnitID = Trim$(Str$(tmAstInfo(llTest).lCode))
'            Do While Len(slXHTUnitID) < 9
'                slXHTUnitID = "0" & slXHTUnitID
'            Loop
'        End If
        blFound = False
        '9113 easier to read
        If tmAstInfo(llTest).iRegionType <= 0 Then
            slISCI = Trim$(tmAstInfo(llTest).sISCI)
        Else
            slISCI = Trim$(tmAstInfo(llTest).sRISCI)
        End If
        If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(slXHTUnitID), slISCI, Trim$(slProgCodeID), "G") Then
            '9452
            If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2 Or bmSendNotCarried) Then
                blExportSpot = True
                Exit For
            End If
        Else
            blFound = True
        End If
        If blFound Then
            'Handle case wher status changed after the original export '9452
            If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged = 2 And Not bmSendNotCarried) Then
                blExportSpot = True
                Exit For
            End If
        End If
    Next llTest
    ilSpotNo = 0
    For llTest = llIndexStart To llIndexEnd Step 1
        ilSpotNo = ilSpotNo + 1
        'Retain spot with status D is set to not aired '9452
        If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2 Or bmSendNotCarried) Then
            '10021
            slXHTUnitID = mCreateUnitIDForCue(HBFORM, slUnitHB & slVefCode5, tmAstInfo(llTest).lCode, slTransmissionID)
            If smUnitIdByAstCodeForBreak <> "Y" Then
                If ilSpotNo <= 9 Then
                    slXHTUnitID = slXHTUnitID & "-" & Trim$(Str$(ilSpotNo))
                Else
                    slXHTUnitID = slXHTUnitID & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
                End If
            End If
'            If smUnitIdByAstCodeForBreak <> "Y" Then
'                'slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'                If ilSpotNo <= 9 Then
'                    slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'                Else
'                    slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
'                End If
'            Else
'                '9113
'                slXHTUnitID = Trim$(Str$(tmAstInfo(llTest).lCode))
'                Do While Len(slXHTUnitID) < 9
'                    slXHTUnitID = "0" & slXHTUnitID
'                Loop
'            End If
            '9113 easier to read
            If tmAstInfo(llTest).iRegionType <= 0 Then
                slISCI = Trim$(tmAstInfo(llTest).sISCI)
            Else
                slISCI = Trim$(tmAstInfo(llTest).sRISCI)
            End If
            If blExportSpot Then
                slType = "R"
            Else
                slType = "U"
            End If
            ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, slISCI, Trim$(slProgCodeID), "G", slType)
            If blExportSpot Then
                ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, slXHTUnitID, slISCI, Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
            End If
        End If
    Next llTest
    mSetXMLCommandPassHB = blExportSpot
End Function
'Private Function mSetXMLCommandPass2(llIndexStart As Long, llIndexEnd As Long, slXDReceiverID As String, slTransmissionID As String, slUnitHB As String, slVefCode5 As String, slProgCodeID As String)
'    Dim blExportSpot As Boolean
'    Dim llTest As Long
'    Dim ilSpotNo As Integer
'    Dim slXHTUnitID As String
'    Dim ilRegionExist As Integer
'    Dim ilRet As Integer
'    Dim blFound As Boolean
'    '9113 not really for 9113, but helped clean up code
'    Dim slISCI As String
'    Dim slType As String
'
'    blExportSpot = False
'    ilSpotNo = 0
'    For llTest = llIndexStart To llIndexEnd Step 1
'        If tmAstInfo(llTest).iRegionType = 2 Then
'            ilRegionExist = True
'        Else
'            ilRegionExist = False
'        End If
'        ilSpotNo = ilSpotNo + 1
'        If smUnitIdByAstCodeForBreak <> "Y" Then
'            'slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'            If ilSpotNo <= 9 Then
'                slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'            Else
'                slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
'            End If
'        Else
'            '9113
'            slXHTUnitID = Trim$(Str$(tmAstInfo(llTest).lCode))
'            Do While Len(slXHTUnitID) < 9
'                slXHTUnitID = "0" & slXHTUnitID
'            Loop
'        End If
'        blFound = False
'        '9113 easier to read
'        If tmAstInfo(llTest).iRegionType <= 0 Then
'            slISCI = Trim$(tmAstInfo(llTest).sISCI)
'        Else
'            slISCI = Trim$(tmAstInfo(llTest).sRISCI)
'        End If
'        If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(slXHTUnitID), slISCI, Trim$(slProgCodeID), "G") Then
'            '9452
'            If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2 Or bmSendNotCarried) Then
'                blExportSpot = True
'                Exit For
'            End If
'        Else
'            blFound = True
'        End If
''        If tmAstInfo(llTest).iRegionType <= 0 Then
''            If smUnitIdByAstCodeForBreak = "Y" Then
''                If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G") Then
''                    If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2) Then
''                        blExportSpot = True
''                        Exit For
''                    End If
''                Else
''                    blFound = True
''                End If
''            Else
''                If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G") Then
''                    If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2) Then
''                        blExportSpot = True
''                        Exit For
''                    End If
''                Else
''                    blFound = True
''                End If
''            End If
''        Else
''            If smUnitIdByAstCodeForBreak = "Y" Then
''                If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G") Then
''                    If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2) Then
''                        blExportSpot = True
''                        Exit For
''                    End If
''                Else
''                    blFound = True
''                End If
''            Else
''                If Not mFindXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G") Then
''                    If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2) Then
''                        blExportSpot = True
''                        Exit For
''                    End If
''                Else
''                    blFound = True
''                End If
''            End If
''        End If
'        If blFound Then
'            'Handle case wher status changed after the original export '9452
'            If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged = 2 And Not bmSendNotCarried) Then
'                blExportSpot = True
'                Exit For
'            End If
'        End If
'    Next llTest
'    ilSpotNo = 0
'    For llTest = llIndexStart To llIndexEnd Step 1
'        ilSpotNo = ilSpotNo + 1
'        'Retain spot with status D is set to not aired '9452
'        If (tgStatusTypes(gGetAirStatus(tmAstInfo(llTest).iStatus)).iPledged <> 2 Or bmSendNotCarried) Then
'            If smUnitIdByAstCodeForBreak <> "Y" Then
'                'slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'                If ilSpotNo <= 9 Then
'                    slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Str$(ilSpotNo))
'                Else
'                    slXHTUnitID = slTransmissionID & slUnitHB & slVefCode5 & "-" & Trim$(Chr(Asc("A") + ilSpotNo - 10))
'                End If
'            Else
'                '9113
'                slXHTUnitID = Trim$(Str$(tmAstInfo(llTest).lCode))
'                Do While Len(slXHTUnitID) < 9
'                    slXHTUnitID = "0" & slXHTUnitID
'                Loop
'            End If
'            '9113 easier to read
'            If tmAstInfo(llTest).iRegionType <= 0 Then
'                slISCI = Trim$(tmAstInfo(llTest).sISCI)
'            Else
'                slISCI = Trim$(tmAstInfo(llTest).sRISCI)
'            End If
'            If blExportSpot Then
'                slType = "R"
'            Else
'                slType = "U"
'            End If
'            ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, slISCI, Trim$(slProgCodeID), "G", slType)
'            If blExportSpot Then
'                ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, slXHTUnitID, slISCI, Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
'            End If
''            If Not blExportSpot Then
''                If tmAstInfo(llTest).iRegionType <= 0 Then
''                    If smUnitIdByAstCodeForBreak = "Y" Then
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G", "U")
''                    Else
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G", "U")
''                    End If
''                Else
''                    If smUnitIdByAstCodeForBreak = "Y" Then
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G", "U")
''                    Else
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G", "U")
''                    End If
''                End If
''            Else
''                If tmAstInfo(llTest).iRegionType <= 0 Then
''                    If smUnitIdByAstCodeForBreak = "Y" Then
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), "", Trim$(slProgCodeID), "G", "R")
''                        '7509 added gsfcode
''                        ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
''                    Else
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, "", Trim$(slProgCodeID), "G", "R")
''                        '7509 added gsfcode
''                        ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sISCI), Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
''                    End If
''                Else
''                    If smUnitIdByAstCodeForBreak = "Y" Then
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), "", Trim$(slProgCodeID), "G", "R")
''                        '7509 added gsfcode
''                        ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, Trim$(Str$(tmAstInfo(llTest).lCode)), Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
''                    Else
''                        ilRet = mSetXHT(Val(slXDReceiverID), slTransmissionID, slXHTUnitID, "", Trim$(slProgCodeID), "G", "R")
''                        '7509 added gsfcode
''                        ilRet = mAddXHT(tmAstInfo(llTest).lAttCode, tmAstInfo(llTest).sFeedDate, Val(slXDReceiverID), slTransmissionID, slXHTUnitID, Trim$(tmAstInfo(llTest).sRISCI), Trim$(slProgCodeID), "G", tmAstInfo(llTest).lgsfCode)
''                    End If
''                End If
''            End If
'        End If
'    Next llTest
'    mSetXMLCommandPass2 = blExportSpot
'End Function

Private Sub mAddSurroundingElement(ilPass As Integer, blIsStart As Boolean)
    Dim slWrite As String
    Dim slCommand As String
    
    If ilPass = 0 Then
        slWrite = "Sites"
    Else
        slWrite = "Insertions"
    End If
    If blIsStart Then
        slCommand = "OT"
    Else
        slCommand = "CT"
        'this gets ready for the next time to write at start
        bmWroteTopElement = False
    End If
   ' mCSIXMLData slCommand, slWrite, ""
    csiXMLData slCommand, slWrite, ""
End Sub
Private Function mGetOrigSpotShortTitle(ilPassForm As Integer, slSection As String, llLstCode As Long, slShortTitle As String) As Integer
    Dim llAdf As Long
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    slShortTitle = ""
    SQLQuery = "SELECT lstAdfCode, lstProd, lstSdfCode"
    'SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
    SQLQuery = SQLQuery & " FROM LST "
    SQLQuery = SQLQuery & " WHERE lstCode =" & Str(llLstCode)
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If igExportSource = 2 Then DoEvents
        If (sgSpfUseProdSptScr <> "P") Then
            llAdf = gBinarySearchAdf(CLng(rst!lstAdfCode))
            If llAdf <> -1 Then
                slShortTitle = Trim$(Left(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6))
                If slShortTitle = "" Then
                    slShortTitle = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
                End If
            End If
            slShortTitle = slShortTitle & "," & Trim$(rst!lstProd)
        Else
            slShortTitle = gGetShortTitle(rst!lstSdfCode)
        End If
        If ilPassForm = 0 Then
            '6744 Dan "CUMULUS" is never returned.  Must've wanted when Cumulus on double head end
           ' If InStr(1, UCase(Trim(slSection)), "CUMULUS", vbBinaryCompare) > 0 Then
            If InStr(1, UCase(Trim(slSection)), "-CU", vbBinaryCompare) > 0 Then
                llAdf = gBinarySearchAdf(CLng(rst!lstAdfCode))
                If llAdf <> -1 Then
                    slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                End If
            End If
        End If
        slShortTitle = UCase$(gFileNameFilter(slShortTitle))
    End If
    mGetOrigSpotShortTitle = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmExportXDigital-mGetOrigSpotShortTitle"
    mGetOrigSpotShortTitle = False
    Exit Function
End Function

Private Sub mRetainDeletions()
    Dim llUpper As Long
    '7509 added to ent
    
    On Error GoTo ErrHand
    bmAllowXMLCommands = True
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked Then
        Exit Sub
    End If
    xhtInfo_rst.Filter = "Status = " & "'D'"
    Do While Not xhtInfo_rst.EOF
        '8357
        If mNotSentOkToDelete(Trim$(xhtInfo_rst!SiteID), Trim$(xhtInfo_rst!UnitID)) Then
            llUpper = UBound(tmRetainDeletions)
            tmRetainDeletions(llUpper).sSiteId = Trim$(xhtInfo_rst!SiteID)
            tmRetainDeletions(llUpper).sTransmissionID = Trim$(xhtInfo_rst!TransmissionID)
            tmRetainDeletions(llUpper).sUnitID = Trim$(xhtInfo_rst!UnitID)
            ReDim Preserve tmRetainDeletions(0 To llUpper + 1) As REATINDELETIONS
            xhtInfo_rst!Status = "R"
            myEnt.Add Format$(xhtInfo_rst!FeedDate, sgShowDateForm), xhtInfo_rst!gsfCode, Deleted
        End If
On Error GoTo 0
        xhtInfo_rst.MoveNext
    Loop
    xhtInfo_rst.Filter = adFilterNone
    Exit Sub
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mRetainDeletions"
    Resume Next
End Sub
'12/20/14: Remove XHT records regardless of agreements that are two or more weeks old
'Private Sub mRemoveOldXHT(llAttCode As Long)
Private Sub mRemoveOldXHT()
    Dim slDate As String
    Dim slXHTDeleteDate As String
    Dim slProp As String

    On Error GoTo ErrHand
    If udcCriteria.XExportType(SPOTINSERTION, "V") = vbUnchecked Then
        Exit Sub
    End If
    '7063 added bmTestForceUpdateXHT for testing
    If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
        slXHTDeleteDate = Format(gDateValue(gObtainPrevMonday(smDate)) - 14, "m/d/yy")
        '12/20/14: Remove XHT records regardless of agreements that are two or more weeks old
        ''7063
        'SQLQuery = "DELETE FROM xht where xhtAttCode = " & llAttCode & " And xhtFeedDate < '" & Format$(slXHTDeleteDate, sgSQLDateForm) & "'"
        ''SQLQuery = "DELETE FROM xht where xhtAttCode = " & llAttCode & " And xhtFeedDate < " & slXHTDeleteDate
        SQLQuery = "DELETE FROM xht where xhtFeedDate < '" & Format$(slXHTDeleteDate, sgSQLDateForm) & "'"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, "Export XDigital-mRemoveOldXHT"
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mRemoveOldXHT"
    Resume Next
'ErrHand1:
'    gHandleError smPathForgLogMsg, "Export XDigital-mRemoveOldXHT"
'    Return
End Sub
Private Function mSafeForTrim(slXmlLine As String, ilMax As Integer) As String
    ' find &pos;(or similar).  If it crosses through the ilMax, we have to get rid of. Loop for multiple
    Dim slRet As String
    Dim ilCurrentAnd As Integer
    Dim ilCurrentSemi As Integer
    Dim ilSafeStart As Integer
    Dim blKeepSearching As Boolean
On Error GoTo ERRORBOX

    ilSafeStart = ilMax - 4
    If InStr(ilSafeStart, slXmlLine, "&") > 0 Then
        slRet = slXmlLine
        blKeepSearching = True
        Do While blKeepSearching
            ilCurrentAnd = InStrRev(slRet, "&", ilMax)
            If ilCurrentAnd >= ilSafeStart Then
                ilCurrentSemi = InStr(ilCurrentAnd, slRet, ";")
                If ilCurrentSemi > ilCurrentAnd Then
                    If ilCurrentSemi <= ilMax Then
                        blKeepSearching = False
                    Else
                        'lose from &
                        slRet = Mid(slRet, 1, ilCurrentAnd - 1) & Mid(slRet, ilCurrentSemi + 1)
                    End If
                Else  'lone & shouldn't happen
                     slRet = Mid(slRet, 0, ilCurrentAnd - 1)
                End If
            Else
                blKeepSearching = False
            End If
        Loop
    Else
        slRet = slXmlLine
    End If
    mSafeForTrim = slRet
    Exit Function
ERRORBOX:
    mSafeForTrim = slXmlLine
End Function
Private Sub mConfirmXHT(slAtts As String, slFeedStartDate As String, slFeedEndDate As String)
    Dim slProp As String
    If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
        'slatt = "4,25,13,"
        If Len(slAtts) > 1 Then
            slAtts = mLoseLastLetterIfComma(slAtts)
            SQLQuery = "UPDATE xht SET xhtStatus = 'Y' where xhtStatus = 'N' And xhtAttCode in (" & slAtts & ") And xhtFeedDate >= '" & Format$(slFeedStartDate, sgSQLDateForm) & "' And xhtFeedDate <= '" & Format$(slFeedEndDate, sgSQLDateForm) & "'"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand
                gHandleError smPathForgLogMsg, "Export XDigital-mConfirmXHT"
            End If
            SQLQuery = "DELETE FROM xht where xhtAttCode in (" & slAtts & ") And xhtStatus = 'D'"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand
                gHandleError smPathForgLogMsg, "Export XDigital-mConfirmXHT"
            End If

        End If
    End If
    Exit Sub
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mConfirmXHT"
    Return
End Sub
Private Function mIgnoreWarning(slMessage As String) As Boolean
    'return true if this message is to be ignored
    'I have mulitple messages in the slMessage, so I have to make sure all of them can be ignored
    Dim blRet As Boolean
    Dim slTestLine As String
    Dim ilMsgCount As Integer
    Dim ilFoundTest As Integer
    'Dim ilPos As Integer
    Dim ilPos As Long

    blRet = False
    ilMsgCount = 0
    ilFoundTest = 0
    slTestLine = "AS THE STATION IS NOT MAPPED TO ANY RECEIVERS SITEID:"
    If InStr(1, slMessage, slTestLine) > 0 Then
        Do
            ilPos = InStr(ilPos + 1, slMessage, "<MSG CODE")
            ilMsgCount = ilMsgCount + 1
        Loop While ilPos > 0
        Do
            ilPos = InStr(ilPos + 1, slMessage, slTestLine)
            ilFoundTest = ilFoundTest + 1
        Loop While ilPos > 0
        '5/21/15 add one for "OK"
        If ilMsgCount = ilFoundTest + 1 And ilMsgCount > 0 Then
            blRet = True
        End If
    End If
    mIgnoreWarning = blRet
End Function

Private Function mBuildDelayAst(ilPassForm As Integer, cprst As ADODB.Recordset) As Integer
    Dim llAst As Long
    Dim blSortRequired As Boolean
    Dim blCheckPreviousWeek As Boolean
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llMoDate As Long
    Dim ilRet As Integer
    Dim slSortDate As String
    Dim slSortTime As String
    Dim slSortSeqNo As String
    Dim cptt_rst As ADODB.Recordset

    On Error GoTo ErrHand
    mBuildDelayAst = True
    ReDim tmDelayAstInfo(0 To 0) As ASTINFO
    If smSupportXDSDelay <> "Y" Then
        Exit Function
    End If
    blSortRequired = False
    llMoDate = gDateValue(gObtainPrevMonday(smDate))
    llStartDate = gDateValue(smDate)
    llEndDate = llStartDate + imNumberDays - 1
    If cprst!attSendDelayToXDS = "Y" Then
        For llAst = 0 To UBound(tmAstInfo) - 1 Step 1
            If (tgStatusTypes(gGetAirStatus(tmAstInfo(llAst).iPledgeStatus)).iPledged = 1) Then
                If (gDateValue(tmAstInfo(llAst).sPledgeDate) >= llStartDate) And (gDateValue(tmAstInfo(llAst).sPledgeDate) <= llEndDate) Then
                    If gDateValue(tmAstInfo(llAst).sPledgeDate) = gDateValue(tmAstInfo(llAst).sFeedDate) Then
                        'Determine if seven day delay or just a time shift
                        If tmAstInfo(llAst).sPdDayFed <> "A" Then
                            'Time shift only
                            tmAstInfo(llAst).sFeedTime = tmAstInfo(llAst).sPledgeStartTime
                            tmAstInfo(llAst).iStatus = 0
                            tmAstInfo(llAst).iAirPlay = 1
                            blSortRequired = True
                        End If
                    Else
                        If gDateValue(tmAstInfo(llAst).sPledgeDate) <= gDateValue(gObtainNextSunday(tmAstInfo(llAst).sFeedDate)) Then
                            tmAstInfo(llAst).sFeedDate = tmAstInfo(llAst).sPledgeDate
                            tmAstInfo(llAst).sFeedTime = tmAstInfo(llAst).sPledgeStartTime
                            tmAstInfo(llAst).iStatus = 0
                            tmAstInfo(llAst).iAirPlay = 1
                            blSortRequired = True
                        End If
                    End If
                Else
                    tmAstInfo(llAst).iStatus = 2
                End If
            End If
        Next llAst
    End If
    'Obtain agreement that covers previous week
    blCheckPreviousWeek = True
    If (gDateValue(cprst!attOnAir) < llMoDate) And (cprst!attSendDelayToXDS <> "Y") Then
        blCheckPreviousWeek = False
    End If
    If blCheckPreviousWeek Then
        SQLQuery = "SELECT cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attVoiceTracked, attXDReceiverID, attSendDelayToXDS FROM Cptt Left Outer Join att On cpttAtfCode = attcode "
        '7701 added
        SQLQuery = SQLQuery & " Left outer join VAT_Vendor_Agreement on attcode = vatAttCode "
        SQLQuery = SQLQuery & " WHERE (cpttVefCode = " & cprst!cpttvefcode & " And cpttStartDate = '" & Format(llMoDate - 7, sgSQLDateForm) & "' and cpttShfCode = " & cprst!shttCode
        If ilPassForm = 0 Then
            '7701
            SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.XDS_ISCI
          '  SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'X'"
        Else
            SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.XDS_Break
           ' SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'B'"
        End If
        SQLQuery = SQLQuery & ")"
        Set cptt_rst = gSQLSelectCall(SQLQuery)
        If Not cptt_rst.EOF Then
            If cptt_rst!attSendDelayToXDS = "Y" Then
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cptt_rst!cpttCode
                tgCPPosting(0).iStatus = cptt_rst!cpttStatus
                tgCPPosting(0).iPostingStatus = cptt_rst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cptt_rst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cptt_rst!attTimeType
                tgCPPosting(0).iVefCode = imVefCode
                tgCPPosting(0).iShttCode = cprst!shttCode
                tgCPPosting(0).sZone = cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(llMoDate - 7, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cptt_rst!cpttAstStatus
                tgCPPosting(0).iNumberDays = 7
                igTimes = 1 'By Week
                If ilPassForm = 0 Then
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmDelayAstInfo(), imAdfCode, True, False, True, , , , , , True)
                    gFilterAstExtendedTypes tmDelayAstInfo
                Else
                    If (smMidnightBasedHours = "Y") And (ilPassForm <> 0) Then
                        '6082 change first 'false' to 'true to get rid of 0 in astcode
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmDelayAstInfo(), imAdfCode, True, False, True, , , True, , , True)
                        gFilterAstExtendedTypes tmDelayAstInfo
                    Else
                        '6082 change first 'false' to 'true to get rid of 0 in astcode
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmDelayAstInfo(), imAdfCode, True, False, True, , , , , , True)
                        gFilterAstExtendedTypes tmDelayAstInfo
                    End If
                End If
                For llAst = 0 To UBound(tmDelayAstInfo) - 1 Step 1
                    If (tgStatusTypes(gGetAirStatus(tmDelayAstInfo(llAst).iPledgeStatus)).iPledged = 1) Then
                        If (gDateValue(tmDelayAstInfo(llAst).sPledgeDate) >= llStartDate) And (gDateValue(tmDelayAstInfo(llAst).sPledgeDate) <= llEndDate) Then
                            tmDelayAstInfo(llAst).sFeedDate = tmDelayAstInfo(llAst).sPledgeDate
                            tmDelayAstInfo(llAst).sFeedTime = tmDelayAstInfo(llAst).sPledgeStartTime
                            tmDelayAstInfo(llAst).iStatus = 0
                            tmDelayAstInfo(llAst).iAirPlay = 1
                            tmAstInfo(UBound(tmAstInfo)) = tmDelayAstInfo(llAst)
                            ReDim Preserve tmAstInfo(0 To UBound(tmAstInfo) + 1) As ASTINFO
                            blSortRequired = True
                        Else
                            tmDelayAstInfo(llAst).iStatus = 2
                        End If
                    End If
                Next llAst
            End If
        End If
    End If
    If blSortRequired Then
        ReDim tmDelaySort(0 To 0) As DELAYSORT
        ReDim tmDelayAstInfo(0 To UBound(tmAstInfo)) As ASTINFO
        For llAst = 0 To UBound(tmAstInfo) - 1 Step 1
            tmDelayAstInfo(llAst) = tmAstInfo(llAst)
            slSortDate = gDateValue(tmAstInfo(llAst).sFeedDate)
            Do While Len(slSortDate) < 7
                slSortDate = "0" & slSortDate
            Loop
            slSortTime = gTimeToLong(tmAstInfo(llAst).sFeedTime, False)
            Do While Len(slSortTime) < 7
                slSortTime = "0" & slSortTime
            Loop
            slSortSeqNo = Trim$(Str(llAst))
            Do While Len(slSortSeqNo) < 7
                slSortSeqNo = "0" & slSortSeqNo
            Loop
            tmDelaySort(UBound(tmDelaySort)).sKey = slSortDate & "|" & slSortTime & "|" & slSortSeqNo
            tmDelaySort(UBound(tmDelaySort)).lAstIndex = llAst
            ReDim Preserve tmDelaySort(0 To UBound(tmDelaySort) + 1) As DELAYSORT
        Next llAst
        If UBound(tmDelaySort) - 1 >= 1 Then
            ArraySortTyp fnAV(tmDelaySort(), 0), UBound(tmDelaySort), 0, LenB(tmDelaySort(0)), 0, LenB(tmDelaySort(0).sKey), 0
        End If
        For llAst = 0 To UBound(tmDelaySort) - 1 Step 1
            tmAstInfo(llAst) = tmDelayAstInfo(tmDelaySort(llAst).lAstIndex)
        Next llAst
    End If
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "frmExportXDigital-mBuildDelayAst"
    mBuildDelayAst = False
    Exit Function
End Function
Private Function mReturnSiteIds(slMessage As String) As String
    '7508
    'return 7,11,41,
    'strip #s from <SetAuthorizationsResult TransmissionID=225 Count=5><msgs><msg code=-2 ID=7 ...etc.
    'Handles spot insertion errors.  1) invalid program name isn't handled--it's a vehicle and we send by vehicle so let it fall through as "", which means whole vehicle gets marked as not sent
    '2) hb returns invalid site id like this <MSG CODE=-2 ID=ESPNH10B3P1 NAME=CUECODE>INVALID SITEID=39294, ITEM IGNORED</MSG>
    '3) isci returns invalid site id like this <MSG CODE=-3 ID=935 NAME=SITEID>SITEID=935 DOES NOT EXIST, INSERTIONS TO THIS SITE IGNORED.</MSG>
    '4) isci also can return 'not mapped'. but those are covered in mignored and so aren't caught here <MSG CODE=-2 ID=20150518000100103 NAME=UNITID>COULD NOT CREATE REGIONAL FILEDELIVERY RECORDS AS THE STATION IS NOT MAPPED TO ANY RECEIVERS SITEID: 167 REGIONALSPOT: RA_SOLE-SOLU-1507</MSG>
    'this code catches 2 and 3
    Dim ilPos As Long
    Dim ilEnd As Long
    Dim slRet As String
    Dim slTemp As String
    
    slRet = ""
    ilPos = 1
On Error GoTo errbox
    'invalid site id
    Do While ilPos > 0
        '-2 too restrictive.  Already tested if 'error', so ignore the code #
        ilPos = InStr(ilPos, slMessage, "<msg code=-")
       ' ilPos = InStr(ilPos, slMessage, "<msg code=-2")
        If ilPos > 0 Then
            ilPos = InStr(ilPos, slMessage, "SITEID=")
            If ilPos > 0 Then
                ilEnd = InStr(ilPos + 1, slMessage, " ")
                If ilEnd > ilPos Then
                    slTemp = Mid(slMessage, ilPos + 7, ilEnd - ilPos - 7)
                    slTemp = mLoseLastLetterIfComma(slTemp)
                    If InStr("," & slRet, "," & slTemp & ",") = 0 Then
                        slRet = slRet & slTemp & ","
                    End If
                End If
            End If
        End If
    Loop
    If Len(slRet) = 0 Then
        bmFailedToReadReturn = True
    End If
    mReturnSiteIds = slRet
    Exit Function
errbox:
    bmFailedToReadReturn = True
    myExport.WriteError "Error in mReturnSiteIds, could not read the returned warning message.", True, False
    mReturnSiteIds = ""
End Function
Private Function mSendAndTestReturn(ilRoutine As XDSType, Optional slDoNotReturn As String = "") As Boolean
    '7508 consolidate into one function
    'return false if warning or error
    '6966 but still true even if errors on first try but successful on resend.
    Dim blRet As Boolean
    Dim c As Integer
    Dim slRet As String
    Dim slStatus As String
    Dim slRoutine As String
    Dim blReturnIds As Boolean
    Dim blReturnSiteIds As Boolean
    
    blReturnSiteIds = False
    blReturnIds = False
    bmAllowXMLCommands = True
    slStatus = ""
    slDoNotReturn = ""
    blRet = True
    Select Case ilRoutine
        Case XDSType.Insertions
            slRoutine = "Spot Insertions"
            blReturnSiteIds = True
        Case XDSType.Authorizations
            slRoutine = "Authorizations"
            blReturnIds = True
        Case XDSType.FILEDELIVERY
            slRoutine = "File Delivery"
        Case XDSType.Stations
            slRoutine = "Stations"
        Case XDSType.Vehicles
            slRoutine = "Programs"
            blReturnIds = True
    End Select
    bmAllowXMLCommands = True
    slStatus = ""
    blRet = True
    If Not mSendBasic(False, False, slRoutine, slStatus) Then
        If bmIsError Then
            'dan 5/7/15
            If imMaxRetries > 0 Then
                mSetResults "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted.", MESSAGERED
                myExport.WriteWarning "Error in sending " & slRoutine & " " & imMaxRetries & " retries will be attempted."
                For c = 1 To imMaxRetries - 1
                    '9375 30 second pause.
                    gSleep 30
                    If mSendBasic(False, True, slRoutine, slStatus) Then
                        Exit For
                    ElseIf bmIsError = False Then
                        blRet = False
                        Exit For
                    End If
                Next c
                If bmIsError Then
                    blRet = mSendBasic(True, True, slRoutine, slStatus)
                End If
            Else
                mSetResults "Error in sending " & slRoutine, MESSAGERED
                myExport.WriteWarning "Error in sending " & slRoutine
            End If
            'resending fixed the issue
            If bmIsError = False Then
                bmAlertAboutReExport = True
                mSetResults "Error in sending " & slRoutine & " corrected. Export Ok and continuing.", MESSAGERED
                myExport.WriteWarning "Error in sending " & slRoutine & " corrected. Export Ok and continuing."
                If blRet = False Then
                    If blReturnIds Then
                        slDoNotReturn = mReturnIds(slStatus)
                    ElseIf blReturnSiteIds Then
                        slDoNotReturn = mReturnSiteIds(slStatus)
                    End If
                End If
            Else
                blRet = False
            End If
        Else
            blRet = False
            If blReturnIds Then
                slDoNotReturn = mReturnIds(slStatus)
            ElseIf blReturnSiteIds Then
                slDoNotReturn = mReturnSiteIds(slStatus)
            End If
            'dan 5715 already written in mSendBasic
           ' myExport.WriteWarning "Warning in sending " & slRoutine & ": " & slStatus
        End If
    Else
    
    End If
    mSendAndTestReturn = blRet
    Exit Function
End Function
Private Function mAdjustAtts(slGood As String, slBad As String) As String
    'change the bad to atts and pull them out of good.  Return the bad atts as 'slBad', and the good atts as 'slGood'
    'both strings come in as ",x,y,"
    Dim slRet As String
    Dim slRemoveThis As String
    Dim slDonts() As String
    Dim c As Integer
    Dim j As Integer
    Dim ilUpper As Integer
    
    slRemoveThis = ""
    slBad = mLoseLastLetter(slBad)
    ilUpper = UBound(tmSiteIdToAttCode)
    If Len(slBad) > 0 And ilUpper > 0 Then
        slDonts = Split(slBad, ",")
        For c = 0 To UBound(slDonts)
            For j = 0 To ilUpper - 1
                If slDonts(c) = tmSiteIdToAttCode(j).sSite Then
                    slRemoveThis = slRemoveThis & tmSiteIdToAttCode(j).sAtt & ","
                    Exit For
                End If
            Next j
        Next c
        slBad = slRemoveThis
        If Len(slGood) > 0 Then
            slRet = mAdjustUpdates(slGood, slRemoveThis) & ","
        Else
            slRet = ""
        End If
    Else
        slRet = slGood
    End If
    mAdjustAtts = slRet
End Function
Private Sub mAddSiteAndAtt(slAtt As String, slSite As String)
    Dim ilUpper As Integer
                
    ilUpper = UBound(tmSiteIdToAttCode)
    tmSiteIdToAttCode(ilUpper).sAtt = slAtt
    tmSiteIdToAttCode(ilUpper).sSite = Trim$(slSite)
    ReDim Preserve tmSiteIdToAttCode(ilUpper + 1)
End Sub
Private Function mLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    mLoseLastLetterIfComma = slNewString
End Function


Private Sub mBuildHeadEndZoneAdjTable()
    Dim slSQLQuery As String
    Dim saf_rst As ADODB.Recordset
    Dim slHeadEndZone As String
    
    imHDAdj(0) = 0
    imHDAdj(1) = 0
    imHDAdj(2) = 0
    imHDAdj(3) = 0
    '9629 removed this block
'    If smMidnightBasedHours <> "Y" Then
'        Exit Sub
'    End If
    slHeadEndZone = ""
    slSQLQuery = "Select safXDSHeadEndZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set saf_rst = gSQLSelectCall(slSQLQuery)
    If Not saf_rst.EOF Then
        slHeadEndZone = saf_rst!safXDSHeadEndZone
    End If
    '9629 add trim
    'If (slHeadEndZone = "") Then
    If (Trim$(slHeadEndZone) = "") Then
        slHeadEndZone = "E"
    End If
    
    'Adjust time as times are relative to location of head end
    If slHeadEndZone = "E" Then
        imHDAdj(0) = 0
        imHDAdj(1) = 1
        imHDAdj(2) = 2
        imHDAdj(3) = 3
    ElseIf slHeadEndZone = "C" Then
        imHDAdj(0) = -1
        imHDAdj(1) = 0
        imHDAdj(2) = 1
        imHDAdj(3) = 2
    ElseIf slHeadEndZone = "M" Then
        imHDAdj(0) = -2
        imHDAdj(1) = -1
        imHDAdj(2) = 0
        imHDAdj(3) = 1
    ElseIf slHeadEndZone = "P" Then
        imHDAdj(0) = -3
        imHDAdj(1) = -2
        imHDAdj(2) = -1
        imHDAdj(3) = 0
    End If
End Sub

Private Function mStationAdj(slZone As String) As Integer

    mStationAdj = 0
    If slZone = "" Then
        Exit Function
    End If
    Select Case Left(slZone, 1)
        Case "E"
            mStationAdj = imHDAdj(0)
        Case "C"
            mStationAdj = imHDAdj(1)
        Case "M"
            mStationAdj = imHDAdj(2)
        Case "P"
            mStationAdj = imHDAdj(3)
    End Select
End Function
'Private Function mGetEventId(llCefCode As Long, slProgram As String, slOldCue As String) As String
'    Dim slRet As String
'    Dim slTemp As String
'    Dim ilPos As Integer
'    'Note that if I don't find anything, I will write out as if normal (the program code has to be 'event' to get here)
'    'also if program code > 8 characters.
'    slProgram = "EVENT"
'    slRet = slOldCue
'    If llCefCode > 0 Then
'        mGetCefComment llCefCode, slTemp
'        slTemp = Trim$(slTemp)
'        ilPos = InStr(slTemp, ":")
'        If ilPos > 0 Then
'            slProgram = Mid(slTemp, 1, ilPos - 1)
'            If Len(slProgram) < 9 Then
'                slRet = Mid(slTemp, ilPos + 1)
'            Else
'                slProgram = "EVENT"
'            End If
'        End If
'    End If
'    mGetEventId = slRet
'End Function
'Private Function mGetEventId(llCefCode As Long, slProgram As String, slOldCue As String, tlEventIds() As EventIdCueAndCode) As String
'    Dim slRet As String
'    Dim slTemp As String
'    Dim ilPos As Integer
'    Dim c As Integer
'    Dim slExtras() As String
'    Dim slTempCue As String
'    Dim slTempCode As String
'    'Note that if I don't find anything, I will write out as if normal (the program code has to be 'event' to get here)
'    'also if program code > 8 characters.
'    slProgram = "EVENT"
'    slRet = slOldCue
'    ReDim tlEventIds(0 To 0) As EventIdCueAndCode
'    If llCefCode > 0 Then
'        mGetCefComment llCefCode, slTemp
'        slTemp = Trim$(slTemp)
'        ilPos = InStr(slTemp, "~")
'        If ilPos > 0 Then
'            slExtras = Split(slTemp, "~")
'            'added extra to end
'            ReDim tlEventIds(0 To UBound(slExtras) + 1) As EventIdCueAndCode
'            For c = 0 To UBound(slExtras)
'                slTempCode = "EVENT"
'                slTempCue = slOldCue
'                If Len(slExtras(c)) > 0 Then
'                    ilPos = InStr(slExtras(c), ":")
'                    If ilPos > 0 Then
'                        slTempCode = Mid(slExtras(c), 1, ilPos - 1)
'                        If Len(slTempCode) < 9 Then
'                            slTempCue = Mid(slExtras(c), ilPos + 1)
'                        Else
'                            slTempCode = "EVENT"
'                        End If
'                    End If
'                    tlEventIds(c).sCode = gXMLNameFilter(slTempCode)
'                    tlEventIds(c).sCue = gXMLNameFilter(slTempCue)
'                End If
'            Next c
'            If Len(tlEventIds(0).sCue) > 0 Then
'                slRet = tlEventIds(0).sCue
'            End If
'            If Len(tlEventIds(0).sCode) > 0 Then
'                slProgram = tlEventIds(0).sCode
'            End If
'        Else
'            ilPos = InStr(slTemp, ":")
'            If ilPos > 0 Then
'                slProgram = Mid(slTemp, 1, ilPos - 1)
'                If Len(slProgram) < 9 Then
'                    slRet = Mid(slTemp, ilPos + 1)
'                    slRet = gXMLNameFilter(slRet)
'                    slProgram = gXMLNameFilter(slProgram)
'                Else
'                    slProgram = "EVENT"
'                End If
'            End If
'        End If
'    End If
'    mGetEventId = slRet
'End Function
Private Function mParseEventId(llCefCode As Long, slProgram As String, slOldCue As String, tlEventIds() As EventIdCueAndCode, dlCuesAndCodes As Dictionary) As String
    Dim slRet As String
    Dim slTemp As String
    Dim ilPos As Integer
    Dim c As Integer
    Dim slExtras() As String
    Dim slTempCue As String
    Dim slTempCode As String
    'Note that if I don't find anything, I will write out as if normal (the program code has to be 'event' to get here)
    'also if program code > 8 characters.
    slProgram = "EVENT"
    slRet = slOldCue
    ReDim tlEventIds(0 To 0) As EventIdCueAndCode
    If dlCuesAndCodes.Exists(llCefCode) Then
        slTemp = dlCuesAndCodes.Item(llCefCode)
        slTemp = Trim$(slTemp)
        ilPos = InStr(slTemp, "~")
        If ilPos > 0 Then
            slExtras = Split(slTemp, "~")
            'added extra to end
            ReDim tlEventIds(0 To UBound(slExtras) + 1) As EventIdCueAndCode
            For c = 0 To UBound(slExtras)
                slTempCode = "EVENT"
                slTempCue = slOldCue
                If Len(slExtras(c)) > 0 Then
                    ilPos = InStr(slExtras(c), ":")
                    If ilPos > 0 Then
                        slTempCode = Mid(slExtras(c), 1, ilPos - 1)
                        If Len(slTempCode) < 9 Then
                            slTempCue = Mid(slExtras(c), ilPos + 1)
                        Else
                            slTempCode = "EVENT"
                        End If
                    End If
                    tlEventIds(c).sCode = gXMLNameFilter(slTempCode)
                    tlEventIds(c).sCue = gXMLNameFilter(slTempCue)
                End If
            Next c
            If Len(tlEventIds(0).sCue) > 0 Then
                slRet = tlEventIds(0).sCue
            End If
            If Len(tlEventIds(0).sCode) > 0 Then
                slProgram = tlEventIds(0).sCode
            End If
        Else
            ilPos = InStr(slTemp, ":")
            If ilPos > 0 Then
                slProgram = Mid(slTemp, 1, ilPos - 1)
                If Len(slProgram) < 9 Then
                    slRet = Mid(slTemp, ilPos + 1)
                    slRet = gXMLNameFilter(slRet)
                    slProgram = gXMLNameFilter(slProgram)
                Else
                    slProgram = "EVENT"
                End If
            End If
        End If
    End If
    mParseEventId = slRet
End Function
Private Function mParseEventIdForZone(llCefCode As Long, slProgram As String, slEventZones() As String, dlCuesAndCodes As Dictionary) As String
    'returns new program, or EVENT
    Dim slRet As String
    Dim slTemp As String
    Dim ilPos As Integer
    Dim c As Integer
    Dim slExtras() As String
    Dim slTempCode As String
    Const ZONEMINIMUM As Integer = 3
    'if something goes wrong, put event into the array, which is set to 3 (matches Pacific as farthest from Eastern
    'also if program code > 8 characters.
    slRet = "EVENT"
    ReDim slEventZones(0 To 0) As String
    If dlCuesAndCodes.Exists(llCefCode) Then
        slTemp = Trim$(dlCuesAndCodes.Item(llCefCode))
        'strip out program
        ilPos = InStr(slTemp, ":")
        'only one colon
        If ilPos = InStrRev(slTemp, ":") Then
            If ilPos > 0 Then
                slTempCode = Mid(slTemp, 1, ilPos - 1)
                If Len(slTempCode) < 9 Then
                    slRet = gXMLNameFilter(slTempCode)
                    'now just the string of cues
                    slTemp = Mid(slTemp, ilPos + 1)
                    'divide the cues
                    ilPos = InStr(slTemp, "~")
                    If ilPos > 0 Then
                        slExtras = Split(slTemp, "~")
                        ReDim slEventZones(0 To UBound(slExtras))
                        For c = 0 To UBound(slExtras)
                            If Len(slExtras(c)) > 0 Then
                                slEventZones(c) = gXMLNameFilter(slExtras(c))
                            End If
                        Next c
                    End If
                End If
            End If
        End If
    End If
    'this will block array from being read later in code
    If UBound(slEventZones) < ZONEMINIMUM Then
        slRet = "EVENT"
    End If
    mParseEventIdForZone = slRet
End Function
Private Function mEventIdsWriteExtraInserts(slProgram As String, slCue As String, slVehicleName As String, slStartDate As String, slEndDate As String, slTransId As String, slLength As String, slEventIdsIscis() As String) As Boolean
    Dim ilIsciIndex As Integer
    Dim ilUpperIscis As Integer
    Dim blRet As Boolean
    
    blRet = False
    If Len(slProgram) > 0 And Len(slCue) > 0 Then
        blRet = True
        ilUpperIscis = UBound(slEventIdsIscis)
        slVehicleName = gXMLNameFilter(slVehicleName)
        mCSIXMLData "OT", "Insert", ""
        mCSIXMLData "CD", "ProgramCode", slProgram
        If Len(slVehicleName) > 0 Then
            mCSIXMLData "CD", "ProgramName", slVehicleName
        End If
        mCSIXMLData "CD", "Cue", slCue
        mCSIXMLData "CD", "StartDate", slStartDate
        mCSIXMLData "CD", "EndDate", slEndDate
        mCSIXMLData "CD", "TransmissionID", slTransId
        mCSIXMLData "OT", "SpotSet", "duration=" & """" & slLength & """"
        For ilIsciIndex = 0 To ilUpperIscis - 1
            mCSIXMLData "CD", "ISCI", slEventIdsIscis(ilIsciIndex)
        Next ilIsciIndex
        mCSIXMLData "CT", "SpotSet", ""
    End If
    mEventIdsWriteExtraInserts = blRet
End Function
Private Sub mEventIdsAddIsci(slISCI As String, slEventIdsIscis() As String)
    Dim ilUpper As Integer
    
    ilUpper = UBound(slEventIdsIscis)
    slEventIdsIscis(ilUpper) = slISCI
    ReDim Preserve slEventIdsIscis(ilUpper + 1)
End Sub
Private Sub mAddToCount(blisReExport As Boolean, ilIsciArrayBound As Integer, llTotalExport As Long, llReExportSent As Long, llNewExportSent As Long, ilNeedExport As Integer)
    Dim ilNewNumber As Integer
    
    ilNewNumber = ilIsciArrayBound
    llTotalExport = llTotalExport + ilNewNumber
    If bmAllowXMLCommands Or bmReExportForce Then
        ilNeedExport = ilNeedExport + ilNewNumber
        '7256
        If blisReExport Then
            llReExportSent = llReExportSent + ilNewNumber
        Else
            llNewExportSent = llNewExportSent + ilNewNumber
        End If
    End If
End Sub
Private Function mEventIDXMLSiteTags(ilIncrement As Integer, slXDReceiverID As String, slTransmissionID As String, slUnitHB As String, slVefCode5 As String) As Boolean
    Dim ilUnitId As Integer
    Dim blRet As Boolean
    Dim slNewUnit As String
    
 On Error GoTo errbox
    blRet = True
    '02002 must be become 02102
    If Len(slUnitHB) = 5 And ilIncrement < 10 And ilIncrement > 0 Then
           slNewUnit = Mid(slUnitHB, 1, 2) & ilIncrement & Mid(slUnitHB, 4)
    Else
        blRet = False
        slNewUnit = slUnitHB
    End If
    mCSIXMLData "OT", "Sites", ""
'        If smUnitIdByAstCode = "Y" Then
'            slUnitIDAstCode = Trim$(Str$(llAstCode))
'            Do While Len(slUnitIDAstCode) < 9
'                slUnitIDAstCode = "0" & slUnitIDAstCode
'            Loop
'            mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slUnitIDAstCode & """"
'        Else
    mCSIXMLData "OT", "Site", "SiteID = " & """" & slXDReceiverID & """" & " UnitID=" & """" & slTransmissionID & slNewUnit & slVefCode5 & """"
    mCSIXMLData "CT", "Site", ""
    mCSIXMLData "CT", "Sites", ""
    mCSIXMLData "CT", "Insert", ""
    mEventIDXMLSiteTags = blRet
    Exit Function
errbox:
    mEventIDXMLSiteTags = False
End Function
'Private Sub mFilterLibraryEventIds(tlAstInfo() As ASTINFO, dlEventCodeAndCue As Dictionary)
'    '8299
'    'O: tlAstInfo--removes those spots that are not defined in the library and so shouldn't go to XDS (If 'event' is program code and not a game.)
'    'O: dlEventCodeAndCue-- save code and cue (PRM:PR05~CRV:CR05) with key being the cefCommentCode
'    Dim llCount As Long
'    Dim llAst As Long
'    Dim slCueAndCode As String
'
'    dlEventCodeAndCue.RemoveAll
'    ReDim tlTempAstInfo(0 To UBound(tlAstInfo)) As ASTINFO
'    llCount = 0
'    For llAst = 0 To UBound(tlAstInfo) - 1 Step 1
'        If mGetEventIdCueAndCode(tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode) Then
'            tlTempAstInfo(llCount) = tlAstInfo(llAst)
'            llCount = llCount + 1
'            If dlEventCodeAndCue.Exists(tmAstInfo(llAst).lEvtIDCefCode) = False Then
'                dlEventCodeAndCue.Add tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode
'            End If
'        End If
'    Next llAst
'    If llCount < UBound(tlAstInfo) Then
'        ReDim tlAstInfo(0 To llCount) As ASTINFO
'        For llAst = 0 To llCount Step 1
'            tlAstInfo(llAst) = tlTempAstInfo(llAst)
'        Next llAst
'    End If
'    Erase tlTempAstInfo
'End Sub
Private Sub mFilterLibraryEventIds(blAlreadySet As Boolean, tlAstInfo() As ASTINFO, dlEventCodeAndCue As Dictionary)
    '8299
    'O: tlAstInfo--removes those spots that are not defined in the library and so shouldn't go to XDS (If 'event' is program code and not a game.)
    'O: dlEventCodeAndCue-- save code and cue (PRM:PR05~CRV:CR05) with key being the cefCommentCode
    Dim llCount As Long
    Dim llAst As Long
    Dim slCueAndCode As String
    
    ReDim tlTempAstInfo(0 To UBound(tlAstInfo)) As ASTINFO
    llCount = 0
    If blAlreadySet = False Then
        dlEventCodeAndCue.RemoveAll
        For llAst = 0 To UBound(tlAstInfo) - 1 Step 1
            If mGetEventIdCueAndCode(tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode) Then
                tlTempAstInfo(llCount) = tlAstInfo(llAst)
                llCount = llCount + 1
                If dlEventCodeAndCue.Exists(tmAstInfo(llAst).lEvtIDCefCode) = False Then
                    dlEventCodeAndCue.Add tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode
                End If
            End If
        Next llAst
    Else
        For llAst = 0 To UBound(tlAstInfo) - 1 Step 1
            If dlEventCodeAndCue.Exists(tmAstInfo(llAst).lEvtIDCefCode) Then
                tlTempAstInfo(llCount) = tlAstInfo(llAst)
                llCount = llCount + 1
            '8371  previous station didn't have it...add it now
            ElseIf mGetEventIdCueAndCode(tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode) Then
                tlTempAstInfo(llCount) = tlAstInfo(llAst)
                llCount = llCount + 1
                If dlEventCodeAndCue.Exists(tmAstInfo(llAst).lEvtIDCefCode) = False Then
                    dlEventCodeAndCue.Add tmAstInfo(llAst).lEvtIDCefCode, slCueAndCode
                End If
            End If
        Next llAst
    End If
    If llCount < UBound(tlAstInfo) Then
        ReDim tlAstInfo(0 To llCount) As ASTINFO
        For llAst = 0 To llCount Step 1
            tlAstInfo(llAst) = tlTempAstInfo(llAst)
        Next llAst
    End If
    Erase tlTempAstInfo
End Sub
Private Function mGetEventIdCueAndCode(llCefCode As Long, slCueAndCode As String) As Boolean
    '8299 true if has event cue and code
    'O: CueAndCode
    Dim blRet As Boolean
    Dim slTemp As String
    
    blRet = False
    slCueAndCode = ""
    If llCefCode > 0 Then
        mGetCefComment llCefCode, slTemp
        slTemp = Trim$(slTemp)
        If UCase(slTemp) <> "-XDS" Then
            blRet = True
            slCueAndCode = slTemp
        End If
    'blank?  count as true so will cause error in send!
    Else
        blRet = True
    End If
    mGetEventIdCueAndCode = blRet
End Function
Private Sub mRetainSiteIdAndUnitId(slSiteId As String, slUnitID As String)
    Dim ilUpper As Integer
    
    If Len(slUnitID) > 0 And Len(slSiteId) > 0 Then
        ilUpper = UBound(tmSentListForDeletionCompare)
        tmSentListForDeletionCompare(ilUpper).sSiteId = slSiteId
        tmSentListForDeletionCompare(ilUpper).sUnitID = slUnitID
        ReDim Preserve tmSentListForDeletionCompare(ilUpper + 1)
    End If
End Sub
Private Function mNotSentOkToDelete(slSiteId As String, slUnitID) As Boolean
    Dim blRet As Boolean
    Dim ilIndex As Integer
    Dim ilPos As Integer
    Dim slUnitIdToTest As String
    
    ilPos = InStr(1, slUnitID, "-", vbTextCompare)
    If ilPos > 0 Then
        slUnitIdToTest = Left(slUnitID, ilPos - 1)
    Else
        slUnitIdToTest = slUnitID
    End If
    blRet = True
    For ilIndex = 0 To UBound(tmSentListForDeletionCompare) - 1
        If Trim$(tmSentListForDeletionCompare(ilIndex).sSiteId) = slSiteId And Trim$(tmSentListForDeletionCompare(ilIndex).sUnitID) = slUnitIdToTest Then
            blRet = False
            Exit For
        End If
    Next ilIndex
    mNotSentOkToDelete = blRet
End Function
Private Sub mFillVehicle()
    
    Dim ilVff As Integer
    Dim ilVpf As Integer
    Dim slXDXMLForm As String
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim llVef As Long
    Dim llGridRow As Long
    Dim llCol As Long
    
    imAnyHBorHBP = False
    ilRet = gPopVff()
    ReDim imMergeVefCode(0 To 0) As Integer
    
    On Error GoTo ErrHand
    chkAll.Value = 0
    grdVeh.Redraw = False
    mClearGrid
    grdVeh.Row = 0
    For llCol = VEHINDEX To SPLITINDEX Step 1
        grdVeh.Col = llCol
        grdVeh.CellBackColor = vbHighlight
    Next llCol
    
    llGridRow = grdVeh.FixedRows
    'Set the column headers background color to light blue
    With grdVeh
        For llCol = .FixedCols To .Cols - 1
            .Col = llCol
            .CellBackColor = LIGHTBLUE
        Next
    End With
    grdVeh.BackColorFixed = LIGHTBLUE
    slNowDate = Format(gNow(), sgSQLDateForm)
    SQLQuery = "SELECT DISTINCT attVefCode FROM att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode  WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND (vatwvtVendorId = " & Vendors.XDS_Break & " OR vatwvtVendorId = " & Vendors.XDS_ISCI & ") "
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        If igExportSource = 2 Then DoEvents
        llVef = gBinarySearchVef(CLng(rst!attvefCode))
        If llVef <> -1 Then
            ilVff = gBinarySearchVff(tgVehicleInfo(llVef).iCode)
            ilVpf = gBinarySearchVpf(CLng(tgVehicleInfo(llVef).iCode))
            '8163 added vehicle state
            If (ilVff <> -1) And (ilVpf <> -1) And tgVehicleInfo(llVef).sState = "A" Then
                If igExportSource = 2 Then DoEvents
                If (Trim$(tgVffInfo(ilVff).sXDProgCodeID) <> "") Or (tgVpfOptions(ilVpf).iInterfaceID > 0) Then
                    If (Trim$(UCase(tgVffInfo(ilVff).sXDProgCodeID)) <> "MERGE") Then
                        'lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                        'lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(llVef).iCode
                        mAddToGrid llGridRow, llVef
                        slXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
                        If (slXDXMLForm = "A") Or (slXDXMLForm = "S") Then
                            imAnyHBorHBP = True
                        End If
                    Else
                        imMergeVefCode(UBound(imMergeVefCode)) = tgVehicleInfo(llVef).iCode
                        ReDim Preserve imMergeVefCode(0 To UBound(imMergeVefCode) + 1) As Integer
                    End If
                End If
            End If
        End If
        rst.MoveNext
    Loop
    mFindAlertsForGrdVeh
    mVehSortCol VEHINDEX
    'mVehSortCol LOGINDEX
    grdVeh.Row = 0
    grdVeh.Col = VEHCODEINDEX
    grdVeh.Redraw = True
    Exit Sub

ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportXDigital-mFillVehicle"
    Exit Sub
End Sub
    
    
Private Sub mFindAlertsForGrdVeh()
 
    Dim rst As ADODB.Recordset
    Dim slMoWeekDate As String
    Dim ilVehCode As Integer
    Dim ilLoop As Integer
    Dim blNeedsLogGened As Boolean
    Dim blRet As Boolean
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slLogNeeded As String
    Dim slPGMNeeded As String
    Dim slSplitNeeded As String
    Dim blRstSetUsed As Boolean
    
    
    If edcDate.Text = "" Then
        Exit Sub
    End If
    If Trim$(txtNumberDays.Text) = "" Then
        Exit Sub
    End If
    '11/3/17
    If (smEDCDate = edcDate.Text) And (smTxtNumberDays = txtNumberDays.Text) Then
        If cmdExport.Enabled = False Then
            cmdExport.Enabled = True
            cmdExportTest.Enabled = True
            cmdCancel.Caption = "&Cancel"
        End If
        Exit Sub
    End If
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    smEDCDate = edcDate.Text
    smTxtNumberDays = txtNumberDays.Text
    
    slLogNeeded = "N"
    slPGMNeeded = "N"
    slSplitNeeded = "N"
    grdVeh.TextMatrix(0, LOGINDEX) = "Gen"
    grdVeh.TextMatrix(0, PGMINDEX) = "Pgm"
    grdVeh.TextMatrix(0, SPLITINDEX) = "Split"
    'mSetGridTitles
    slStartDate = edcDate.Text
    slEndDate = DateAdd("d", CInt(txtNumberDays.Text) - 1, slStartDate)
    slMoWeekDate = gObtainPrevMonday(edcDate.Text)
    smEndDate = DateAdd("d", CInt(txtNumberDays.Text) - 1, slMoWeekDate)
    blRstSetUsed = False
    grdVeh.Visible = False
    Do
        For ilLoop = grdVeh.FixedRows To grdVeh.Rows - 1
            
            If Trim(grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)) <> "" Then
                ilVehCode = Trim(grdVeh.TextMatrix(ilLoop, VEHCODEINDEX))
                smVefName = grdVeh.TextMatrix(ilLoop, VEHINDEX)
                blNeedsLogGened = False
            '*** Check for log alerts ***
            SQLQuery = "Select * from AUF_Alert_User where "
            SQLQuery = SQLQuery & "aufStatus = 'R' "
            SQLQuery = SQLQuery & "and aufType = 'L' "
            SQLQuery = SQLQuery & "and aufSubType <> 'M' "
            SQLQuery = SQLQuery & "and aufSubType <> '' "
            SQLQuery = SQLQuery & "and aufMoWeekDate = " & "'" & Format$(slMoWeekDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery & "and aufVefCode = " & ilVehCode
            Set rst = gSQLSelectCall(SQLQuery)
                blRstSetUsed = True
            If rst.EOF Then
                'No alerts found; now check to see if the log for the given week has been generated
                    
                    If ilVehCode = 3 Then
                        blNeedsLogGened = blNeedsLogGened
                    End If
                    blNeedsLogGened = mLogNeedsToBeGenerated(ilVehCode, slStartDate, slEndDate)
            Else
                    blNeedsLogGened = True
            End If
            If blNeedsLogGened Then
                grdVeh.TextMatrix(0, LOGINDEX) = "Gen *"
                With grdVeh
                    .Row = ilLoop
                    .Col = LOGINDEX
                    '.CellFontName = "Monotype Sorts"
                    .TextMatrix(ilLoop, LOGINDEX) = ""
                    .TextMatrix(ilLoop, LOGSORTINDEX) = "A"
                    If .CellBackColor <> vbRed Then
                        .CellBackColor = vbRed
                        .CellForeColor = vbRed
                    End If
                End With
                
                slLogNeeded = "Y"
            Else
               With grdVeh
                    .Row = ilLoop
                    .Col = LOGINDEX
                    '.CellFontName = "Monotype Sorts"
                    .TextMatrix(ilLoop, LOGINDEX) = ""
                    .TextMatrix(ilLoop, LOGSORTINDEX) = "B"
                    If .CellBackColor <> vbWhite Then
                        .CellBackColor = vbWhite
                        .CellForeColor = vbWhite
                    End If
                End With
            End If
            '*** Check for program change alerts ***
            SQLQuery = "Select * from AUF_Alert_User where "
            SQLQuery = SQLQuery & "aufStatus = 'R' "
            SQLQuery = SQLQuery & "and aufType = 'P' "
            '2/15/18: A= Agreement changed
            'SQLQuery = SQLQuery & "and aufSubType <> '' "
            SQLQuery = SQLQuery & "and aufSubType = 'A' "
            SQLQuery = SQLQuery & "and aufMoWeekDate <= " & "'" & Format$(smEndDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery & "and aufVefCode = " & ilVehCode
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                grdVeh.TextMatrix(0, PGMINDEX) = "Pgm *"
                With grdVeh
                    .Row = ilLoop
                    .Col = PGMINDEX
                    '.CellFontName = "Monotype Sorts"
                    .TextMatrix(ilLoop, PGMINDEX) = ""
                    .TextMatrix(ilLoop, PGMSORTINDEX) = "A"
                    If .CellBackColor <> vbRed Then
                        .CellBackColor = vbRed 'Red
                        .CellForeColor = vbRed 'Red
                    End If
                End With
                slPGMNeeded = "Y"
            Else
               With grdVeh
                    .Row = ilLoop
                    .Col = PGMINDEX
                    '.CellFontName = "Monotype Sorts"
                    .TextMatrix(ilLoop, PGMINDEX) = ""
                    .TextMatrix(ilLoop, PGMSORTINDEX) = "B"
                    If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "0" Then
                        If .CellBackColor <> vbWhite Then
                            .CellBackColor = vbWhite
                            .CellForeColor = vbWhite
                        End If
                    End If
                End With
            End If
    
            '*** Check for split copy ***
                If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
                    blRet = gSplitFillDefined(ilVehCode, slStartDate, slEndDate)
                    If Not blRet Then
                        grdVeh.TextMatrix(0, SPLITINDEX) = "Split *"
                        With grdVeh
                            .Row = ilLoop
                            .Col = SPLITINDEX
                            '.CellFontName = "Monotype Sorts"
                            .TextMatrix(ilLoop, SPLITINDEX) = ""
                            .TextMatrix(ilLoop, SPLITSORTINDEX) = "A"
                            If .CellBackColor <> vbRed Then
                                .CellBackColor = vbRed 'Red
                                .CellForeColor = vbRed 'Red
                            End If
                        End With
                        slSplitNeeded = "Y"
                    Else
                       With grdVeh
                            .Row = ilLoop
                            .Col = SPLITINDEX
                            '.CellFontName = "Monotype Sorts"
                            .TextMatrix(ilLoop, SPLITINDEX) = ""
                            .TextMatrix(ilLoop, SPLITSORTINDEX) = "B"
                            If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "0" Then
                                If .CellBackColor <> vbWhite Then
                                    .CellBackColor = vbWhite
                                    .CellForeColor = vbWhite
                                End If
                            End If
                        End With
                    End If
                End If

            End If
        Next
        slMoWeekDate = DateAdd("d", 7, slMoWeekDate)
    Loop While DateValue(gAdjYear(slMoWeekDate)) < DateValue(gAdjYear(smEndDate))
    grdVeh.Visible = True
    grdVeh.Redraw = True
    If blRstSetUsed Then
        rst.Close
    End If
    mCreateMessage slLogNeeded, slPGMNeeded, slSplitNeeded
    gSetMousePointer grdVeh, grdVeh, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot - mFindAlertsForgrdVeh: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        gLogMsg gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
    End If
    Exit Sub
End Sub
Private Function mLogNeedsToBeGenerated(iVefCode As Integer, sStartDate As String, sEndDate As String) As Boolean

    Dim blFound As Boolean
    Dim slLLD As String
    Dim llLLD As Long
    Dim llVpf As Long
    Dim rst_Vpf As ADODB.Recordset

                
    On Error GoTo ErrHand
    mLogNeedsToBeGenerated = True
    '11/26/17
    llVpf = gBinarySearchVpf(CLng(iVefCode))
    If llVpf <> -1 Then
        slLLD = tgVpfOptions(llVpf).sLLD
        If Trim$(slLLD) = "" Then
            Exit Function
        End If
        llLLD = gDateValue(slLLD)
        If llLLD < gDateValue(sEndDate) Then
            If gProgramDefined(iVefCode, DateAdd("d", 1, slLLD), sEndDate) Then
                Exit Function
            End If
        End If
    Else
        SQLQuery = "SELECT vpfLLD"
        SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
        SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & iVefCode & ")"
        Set rst_Vpf = gSQLSelectCall(SQLQuery)
        If Not rst_Vpf.EOF Then
            If IsNull(rst_Vpf!vpfLLD) Or (Trim$(rst_Vpf!vpfLLD) = "") Then
                Exit Function
            Else
                If Not gIsDate(rst_Vpf!vpfLLD) Then
                    Exit Function
                Else
                    'set sLLD to last log date
                    slLLD = Format$(rst_Vpf!vpfLLD, sgShowDateForm)
                    llLLD = gDateValue(slLLD)
                    'If llLLD < gDateValue(sStartDate) Then
                    '    Exit Function
                    'End If
                    If llLLD < gDateValue(sEndDate) Then
                        If gProgramDefined(iVefCode, DateAdd("d", 1, slLLD), sEndDate) Then
                            Exit Function
                        End If
                    End If
                End If
            End If
        Else
            Exit Function
        End If
    End If
    mLogNeedsToBeGenerated = False
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot - mFindLastLogDate: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        gLogMsg gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
    End If
    Exit Function
End Function
Private Sub mVehSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
        slStr = Trim$(grdVeh.TextMatrix(llRow, VEHINDEX))
        If slStr <> "" Then
            If ilCol = LOGINDEX Then
                slSort = UCase$(Trim$(grdVeh.TextMatrix(llRow, LOGSORTINDEX)))
            ElseIf ilCol = PGMINDEX Then
                slSort = UCase$(Trim$(grdVeh.TextMatrix(llRow, PGMSORTINDEX)))
            ElseIf ilCol = SPLITINDEX Then
                slSort = UCase$(Trim$(grdVeh.TextMatrix(llRow, SPLITSORTINDEX)))
            Else
                slSort = UCase$(Trim$(grdVeh.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdVeh.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastVehColSorted) Or ((ilCol = imLastVehColSorted) And (imLastVehSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVeh.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVeh.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastVehColSorted Then
        imLastVehColSorted = SORTINDEX
    Else
        imLastVehColSorted = -1
        imLastVehSort = -1
    End If
    gGrid_SortByCol grdVeh, VEHINDEX, SORTINDEX, imLastVehColSorted, imLastVehSort
    imLastVehColSorted = ilCol
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mVehSortCol"
    Exit Sub
End Sub

Private Sub mClearGrid()
    
    Dim llRow As Long
    Dim llCol As Long
    
    On Error GoTo ErrHand
    imBypassAll = True
    chkAll.Value = vbUnchecked
    imBypassAll = False
    gGrid_Clear grdVeh, True
    
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
        grdVeh.Row = llRow
        For llCol = 0 To VEHCODEINDEX Step 1
            grdVeh.Col = llCol
            If grdVeh.CellBackColor <> vbWhite Then
                grdVeh.CellBackColor = vbWhite
            End If
            grdVeh.TextMatrix(llRow, llCol) = ""
        Next llCol
    Next llRow
    lmLastClickedRow = -1
    imLastVehColSorted = -1
    imLastVehSort = -1
    lmScrollTop = grdVeh.FixedRows
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mClearGrid"
    Exit Sub
End Sub

Private Sub mAddToGrid(llRow As Long, llVeh As Long)

    Dim llCol As Long
    
    On Error GoTo ErrHand
    If llRow >= grdVeh.Rows Then
        grdVeh.AddItem ""
    End If
    grdVeh.Row = llRow
    For llCol = VEHINDEX To SPLITINDEX Step 1
        grdVeh.Col = llCol
        grdVeh.CellBackColor = vbWhite
        grdVeh.CellForeColor = vbWindowText
    Next llCol
    grdVeh.TextMatrix(llRow, VEHINDEX) = Trim$(tgVehicleInfo(llVeh).sVehicle)
    grdVeh.TextMatrix(llRow, VEHCODEINDEX) = Trim$(tgVehicleInfo(llVeh).iCode)
    grdVeh.TextMatrix(llRow, LOGINDEX) = ""
    grdVeh.TextMatrix(llRow, PGMINDEX) = ""
    grdVeh.TextMatrix(llRow, SPLITINDEX) = ""
    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
    llRow = llRow + 1
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mAddToGrid"
    Exit Sub
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdVeh.ColWidth(SORTINDEX) = 0
    grdVeh.ColWidth(SELECTEDINDEX) = 0
    grdVeh.ColWidth(VEHCODEINDEX) = 0
    grdVeh.ColWidth(LOGSORTINDEX) = 0
    grdVeh.ColWidth(PGMSORTINDEX) = 0
    grdVeh.ColWidth(SPLITSORTINDEX) = 0
    grdVeh.ColWidth(LOGINDEX) = grdVeh.Width * 0.1
    grdVeh.ColWidth(PGMINDEX) = grdVeh.Width * 0.1
    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        grdVeh.ColWidth(SPLITINDEX) = grdVeh.Width * 0.1
    Else
        grdVeh.ColWidth(SPLITINDEX) = 0
    End If
    grdVeh.ColWidth(VEHINDEX) = grdVeh.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To SPLITINDEX Step 1
        If ilCol <> VEHINDEX Then
            grdVeh.ColWidth(VEHINDEX) = grdVeh.ColWidth(VEHINDEX) - grdVeh.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdVeh
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdVeh.TextMatrix(0, VEHINDEX) = "Vehicle"
    grdVeh.TextMatrix(1, VEHINDEX) = "Name"
    grdVeh.TextMatrix(0, LOGINDEX) = "Gen"
    grdVeh.TextMatrix(1, LOGINDEX) = "Log"
    grdVeh.TextMatrix(0, PGMINDEX) = "Pgm"
    grdVeh.TextMatrix(1, PGMINDEX) = "Chg"
    grdVeh.TextMatrix(0, SPLITINDEX) = "Split"
    grdVeh.TextMatrix(1, SPLITINDEX) = "Fill"
End Sub

Private Function mGetGrdSelCount() As Long

    Dim llRow As Long
    Dim llCol As Long
    Dim llCount As Long
    
    On Error GoTo ErrHand
    llCount = 0
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1
        If Trim(grdVeh.TextMatrix(llRow, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(llRow, VEHCODEINDEX)
                llCount = llCount + 1
            End If
        End If
    Next llRow
    mGetGrdSelCount = llCount
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mGetGrdSelCount"
End Function


Private Function mFindDuplVeh(iVehCode As Integer) As Boolean

    Dim llRow As Long
    Dim llCol As Long
    
    On Error GoTo ErrHand
    mFindDuplVeh = False
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1
        If Trim(grdVeh.TextMatrix(llRow, VEHINDEX)) <> "" Then
            If Val(grdVeh.TextMatrix(llRow, VEHCODEINDEX)) = iVehCode Then
                mFindDuplVeh = True
                Exit Function
            End If
        End If
    Next llRow
    Exit Function
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportXDigital-mFindDuplVeh"
End Function


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdVeh.Row = llRow
    For llCol = VEHINDEX To VEHINDEX Step 1
        grdVeh.Col = llCol
        If grdVeh.TextMatrix(llRow, LOGSORTINDEX) = "A" And (llCol = LOGINDEX) Then
            grdVeh.CellBackColor = vbRed
            grdVeh.CellForeColor = vbRed
        ElseIf grdVeh.TextMatrix(llRow, PGMSORTINDEX) = "A" And (llCol = PGMINDEX) Then
            grdVeh.CellBackColor = vbRed
            grdVeh.CellForeColor = vbRed
        ElseIf grdVeh.TextMatrix(llRow, SPLITSORTINDEX) = "A" And (llCol = SPLITINDEX) Then
            grdVeh.CellBackColor = vbRed
            grdVeh.CellForeColor = vbRed
        Else
            If grdVeh.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
                grdVeh.CellBackColor = vbWhite
                grdVeh.CellForeColor = vbWindowText
            Else
                grdVeh.CellBackColor = vbHighlight
                grdVeh.CellForeColor = vbWhite
            End If
        End If
    Next llCol
End Sub
Private Sub grdVeh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdVeh.RowHeight(0) Then
        grdVeh.Col = grdVeh.MouseCol
        mVehSortCol grdVeh.Col
        grdVeh.Row = 0
        grdVeh.Col = VEHCODEINDEX
        Exit Sub
    End If
    'D.S. 07-28-17
    'ilFound = gGrid_GetRowCol(grdVeh, X, Y, llCurrentRow, llCol)
    llCurrentRow = grdVeh.MouseRow
    llCol = grdVeh.MouseCol
    If llCurrentRow < grdVeh.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdVeh.FixedRows Then
        If grdVeh.TextMatrix(llCurrentRow, VEHINDEX) <> "" Then
            grdVeh.TopRow = lmScrollTop
            llTopRow = grdVeh.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdVeh.TextMatrix(grdVeh.Row, VEHCODEINDEX) <> "" Then
                    If grdVeh.TextMatrix(grdVeh.Row, SELECTEDINDEX) <> "1" Then
                        grdVeh.TextMatrix(grdVeh.Row, SELECTEDINDEX) = "1"
                    Else
                        grdVeh.TextMatrix(grdVeh.Row, SELECTEDINDEX) = "0"
                    End If
                    mPaintRowColor grdVeh.Row
                End If
            Else
                For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
                    If grdVeh.TextMatrix(llRow, VEHINDEX) <> "" Then
                        grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
                        If grdVeh.TextMatrix(llRow, VEHCODEINDEX) <> "" Then
                            If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                                If llRow = llCurrentRow Then
                                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                Else
                                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
                                End If
                            ElseIf lmLastClickedRow < llCurrentRow Then
                                If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            Else
                                If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                                    grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                End If
                            End If
                            mPaintRowColor llRow
                        End If
                    End If
                Next llRow
                grdVeh.TopRow = llTopRow
                grdVeh.Row = llCurrentRow
            End If
            lmLastClickedRow = llCurrentRow
            mShowStations
        End If
    End If
    smGridTypeAhead = ""
    mSetCommands
End Sub

Public Sub mCreateMessage(sLog As String, sPGM As String, sSplit As String)

    Dim blMakeMessVisible As Boolean
    
    On Error GoTo ErrHand
    blMakeMessVisible = False
    lblNote.Visible = False
    lblNote.ForeColor = vbRed
    
    'Log, PGM and Splits need checking
    If sLog = "Y" And sPGM = "Y" And sSplit = "Y" Then
        lblNote.Caption = "* Red Box: Generate Log, Check Programming and Create Network Split Fills before running Export."
        blMakeMessVisible = True
    End If
    
    'Log Permutations
    If sLog = "Y" And sPGM = "N" And sSplit = "N" Then
        lblNote.Caption = "* Red Box: Generate Logs before running Export."
        blMakeMessVisible = True
    End If
    
    If sLog = "Y" And sPGM = "Y" And sSplit = "N" Then
        lblNote.Caption = "* Red Box: Generate Logs and Check Programming before running Export."
        blMakeMessVisible = True
    End If
    
    If sLog = "Y" And sPGM = "N" And sSplit = "Y" Then
        lblNote.Caption = "* Red Box: Generate Logs and Network Split Fills before running Export."
        blMakeMessVisible = True
    End If
    
    'PGM Chg Permutations
     If sLog = "N" And sPGM = "Y" And sSplit = "N" Then
        lblNote.Caption = "* Red Box: Check Programming before running Export."
        blMakeMessVisible = True
    End If
    
    If sLog = "N" And sPGM = "Y" And sSplit = "Y" Then
        lblNote.Caption = "* Red Box: Check Programming and Create Network Split Fills before running Export."
        blMakeMessVisible = True
    End If
    
    'Split Network Permutations
    If sLog = "N" And sPGM = "N" And sSplit = "Y" Then
        lblNote.Caption = "* Red Box: Create Network Split Fills before running Export."
        blMakeMessVisible = True
    End If
    
    If blMakeMessVisible Then
        lblNote.Visible = True
    End If
    Exit Sub

ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot - mFindLastLogDate: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        gLogMsg gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
    End If
    Exit Sub
End Sub

Private Sub mShowStations()
    lbcStation.Clear
    If mGetGrdSelCount() = 1 Then
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    Else
        edcTitle3.Visible = False
        chkAllStation.Visible = False
        lbcStation.Visible = False
    End If
    imBypassAll = True
    chkAll.Value = vbUnchecked
    imBypassAll = False
End Sub

Private Sub mSetLogPgmSplitColumns()
    If IsDate(edcDate.Text) = False Then
        'edcDate.SetFocus
        Exit Sub
    End If
    If Trim$(txtNumberDays.Text) = "" Then
        Exit Sub
    End If
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    grdVeh.Redraw = False
    mFindAlertsForGrdVeh
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    imLastVehColSorted = -1
    imLastVehSort = -1
    mVehSortCol VEHINDEX
    'mVehSortCol LOGINDEX
    grdVeh.Row = 0
    grdVeh.Col = VEHCODEINDEX
    grdVeh.Redraw = True
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdExportTest.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    gSetMousePointer grdVeh, grdVeh, vbDefault
End Sub

Private Sub txtNumberDays_Change()
    tmcDelay.Enabled = False
    tmcDelay.Interval = 3000
    tmcDelay.Enabled = True
End Sub

Private Sub txtNumberDays_GotFocus()
    '11/3/17
    'tmcDelay.Enabled = False
    
    'cmdExport.Enabled = False
End Sub

Private Sub txtNumberDays_LostFocus()
    tmcDelay.Enabled = False
    tmcDelay.Interval = 500
    tmcDelay.Enabled = True
End Sub

Private Sub mSetCommands()

    Dim ilEnable As Integer
    Dim llRow As Long

    ilEnable = False
    If (edcDate.Text <> "") And (txtNumberDays.Text <> "") Then
        For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
            If grdVeh.TextMatrix(llRow, VEHINDEX) <> "" Then
                If grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                    ilEnable = True
                    Exit For
                End If
            End If
        Next llRow
    End If
    
    cmdExport.Enabled = ilEnable
End Sub
'7675
Public Function mSendISCIDeletes(llAttCode As Long, slFeedStartDate As String, slFeedEndDate As String) As Boolean
    'true if no issue
    'if problem, get out.  only on trying to delete and it doesn't work do I resume next
    Dim blRet As Boolean
    Dim llUpper As Long
    Dim blContinue As Boolean
    Dim slPrevTransmissionID As String
    Dim slPrevSiteId As String
    Dim slPrevUnitId As String
    Dim slPrevProgCodeId As String
    'for V81
    Dim slDoNotReturn As String
    
    Dim slProp As String
    
    blContinue = False
    blRet = True
On Error GoTo ErrHand1
    SQLQuery = "SELECT * FROM xht"
    SQLQuery = SQLQuery & " WHERE xhtAttCode = " & llAttCode
    SQLQuery = SQLQuery & " AND xhtFeedDate >= '" & Format(slFeedStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND xhtFeedDate <= '" & Format(slFeedEndDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " Order By xhtCode"
    Set xht_rst = cnn.Execute(SQLQuery)
    Do While Not xht_rst.EOF
        'Regionals these are the only ones we care about!
        If (slPrevTransmissionID = Trim$(xht_rst!xhtTransmissionId)) And (slPrevSiteId = Trim$(xht_rst!xhtSiteId)) And (slPrevUnitId = Trim$(xht_rst!xhtunitid)) And (slPrevProgCodeId = Trim$(xht_rst!xhtProgCodeID)) Then
            blContinue = True
            llUpper = UBound(tmRetainDeletions)
            tmRetainDeletions(llUpper).sSiteId = Trim$(xht_rst!xhtSiteId)
            tmRetainDeletions(llUpper).sTransmissionID = Trim$(xht_rst!xhtTransmissionId)
            tmRetainDeletions(llUpper).sUnitID = Trim$(xht_rst!xhtunitid)
            ReDim Preserve tmRetainDeletions(0 To llUpper + 1) As REATINDELETIONS
            'now delete regional!
            If udcCriteria.XGenType(0, slProp) Or bmTestForceUpdateXHT Then
                SQLQuery = "DELETE from xht where XhtCode = " & xht_rst!xhtcode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    GoSub ErrHand:
                End If
            End If
        'generics
        Else
            slPrevTransmissionID = Trim$(xht_rst!xhtTransmissionId)
            slPrevSiteId = Trim$(xht_rst!xhtSiteId)
            slPrevUnitId = Trim$(xht_rst!xhtunitid)
            slPrevProgCodeId = Trim$(xht_rst!xhtProgCodeID)
        End If
        xht_rst.MoveNext
    Loop
    If blContinue Then
        '7797
        If bmWroteTopElement Then
            bmWroteTopElement = False
            csiXMLData "CT", "Sites", ""
        End If
        mSendDeleteCommands
        If Not mSendAndTestReturn(XDSType.Insertions, slDoNotReturn) Then
            blRet = False
            myExport.WriteWarning "Problem above sending deletes for attcode #" & llAttCode
        End If
    End If
    ReDim tmRetainDeletions(0)
Cleanup:
    
    mSendISCIDeletes = blRet
Exit Function
ErrHand1:
    gHandleError smPathForgLogMsg, "Export XDigital-mSendISCIDeletes"
    blRet = False
    GoTo Cleanup
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, "Export XDigital-mSendISCIDeletes"
    blRet = False
    Resume Next
End Function
Private Function mCreateUnitIDForCue(ilPassForm As Integer, slCurrentUnitIdWithVefCode5 As String, llAstCode As Long, slTransmissionID As String) As String
    '10021 9818
    Dim slRet As String
    If smUnitIdByAstCodeForBreak = "Y" Then
        slRet = Trim$(Str$(llAstCode))
        Do While Len(slRet) < 9
            slRet = "0" & slRet
        Loop
        If imSharedHeadEndCue > 0 Then
            slRet = imSharedHeadEndCue & slRet
        End If
    Else
        If imSharedHeadEndCue > 0 Then
            slRet = imSharedHeadEndCue & Mid(slTransmissionID, 4) & slCurrentUnitIdWithVefCode5
        Else
            If ilPassForm = HBPFORM Then
                slRet = Mid(slTransmissionID, 3) & slCurrentUnitIdWithVefCode5
            Else
                slRet = slTransmissionID & slCurrentUnitIdWithVefCode5
            End If
        End If
    End If
    mCreateUnitIDForCue = Trim$(slRet)
End Function

