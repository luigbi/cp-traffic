VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form FrmExportMarketron 
   Caption         =   "Export Marketron"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9615
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   9240
      Top             =   2760
   End
   Begin V81Affiliate.CSI_Calendar CSI_Calendar1 
      Height          =   300
      Left            =   3465
      TabIndex        =   1
      Top             =   135
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      Text            =   "7/10/2020"
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9225
      Top             =   3675
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
      Left            =   8895
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   15
      Top             =   4290
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1845
      Width           =   3825
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   1845
      Width           =   1635
   End
   Begin VB.CommandButton cmdExportTest 
      Caption         =   "Export in Test Mode"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   5010
      Width           =   1665
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   6
      Top             =   4245
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportMarketron.frx":0000
      Left            =   4200
      List            =   "AffExportMarketron.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2115
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   4245
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3375
      Left            =   6195
      TabIndex        =   8
      Top             =   765
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportMarketron.frx":0004
      Left            =   120
      List            =   "AffExportMarketron.frx":0006
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2115
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   615
      Top             =   4650
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5475
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   5010
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   10
      Top             =   5010
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label lacResult 
      Height          =   405
      Left            =   75
      TabIndex        =   12
      Top             =   4560
      Width           =   9240
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6570
      TabIndex        =   7
      Top             =   345
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Export Start Date (Monday of week)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2955
   End
   Begin VB.Menu mnuGuide 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuTest 
         Caption         =   "Test as if Sent"
      End
   End
End
Attribute VB_Name = "FrmExportMarketron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************
'*  frmExportMarketron
'*
'*  Created October 2010 by Dan Michaelson
'*  Copied from FrmExportXDigital
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private imExportMode As Integer     '0=Standard export; 1=Test Mode
Private smStartDate As String     'Export Date
Private smEndDate As String
Private imVefCode As Integer
Private imAdfCode As Integer
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private crf_rst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private hmCsf As Integer
Private bmCsfOpenStatus As Boolean
Private bmAtLeastOneExport As Boolean
Private smFields(13) As String
Private smIniPath As String

Private lmEqtCode As Long

'marketron T, file F, both B
Private smOutputType As String
Private Const FIELDCOUNT As Integer = 13
Private Const SPOTID As Integer = 0
Private Const Advertiser As Integer = 1
Private Const PRODUCTNAME As Integer = 2
Private Const PRODUCTCODE As Integer = 3
Private Const startDate As Integer = 4
Private Const endDate As Integer = 5
Private Const startTime As Integer = 6
Private Const endTime As Integer = 7
Private Const LENGTH As Integer = 8
Private Const ISCICODE As Integer = 9
Private Const COPYTITLE As Integer = 10
Private Const COPYENDDATE As Integer = 12
Private Const COPYCOMMENT As Integer = 11
Private Const XMLDATE As String = "yyyy-mm-dd"
'Private Const LOGFILE As String = "MarketronExportLog.txt"
Private Const FORMNAME As String = "FrmExportMarketron"
Private Const FILEFACTS As String = "MarketronFacts"
Private Const FILEERROR As String = "MarketronExport"
Private Const FILEDEBUG As String = "MarketronDebug"
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
'6633
'Const XMLERRORFILE As String = "XMLErrorResponse.txt"
Private smXmlErrorFile As String
Private smXmlFileFindError As String

Private lmMaxWidth As Long
'replaces logfile
Private smPathForgLogMsg As String
Private myErrors As CLogger
'5349
Private omExport As CMarketron
Private bmStationNotExist As Boolean
'5050
Dim rsFacts As ADODB.Recordset

    'dan M I replaced mFillVehicle with mFillMarketronVehicle to limit the vehicles that appear in the list box.
Private Sub mFillVehicle()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim llVef As Long
    
    ilRet = gPopVff()
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = vbUnchecked
    'For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
    '    lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
    '    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    'Next iLoop
    slNowDate = Format(gNow(), sgSQLDateForm)
    '7701
    SQLQuery = "SELECT DISTINCT attVefCode FROM att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode  WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND attExportType <> 0 AND vatwvtVendorId = " & Vendors.NetworkConnect
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        llVef = gBinarySearchVef(CLng(rst!attvefCode))
        If llVef <> -1 Then
            '8162
            If tgVehicleInfo(llVef).sState = "A" Then
                lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                lbcVehicles.ItemData(lbcVehicles.NewIndex) = rst!attvefCode
            End If
        End If
        rst.MoveNext
    Loop
End Sub
'7701 no longer valid
'Private Sub mFillMarketronVehicle()
'    Dim slSql As String
'    Dim myRs As ADODB.Recordset
'    lbcVehicles.Clear
'    chkAll.Value = 0
'    slSql = "select distinct attVefCode, VefName from att,VEF_Vehicles WHERE attExportToMarketron = 'Y' AND attVefCode = vefCode and attoffAir >= '" & Format(smStartDate, sgSQLDateForm) & "' "
'    Set myRs = gSQLSelectCall(slSql)
'    If Not (myRs.EOF Or myRs.BOF) Then
'        myRs.MoveFirst
'        Do While Not myRs.EOF
'            lbcVehicles.AddItem Trim$(myRs!vefName)
'            lbcVehicles.ItemData(lbcVehicles.NewIndex) = myRs!attvefCode
'            myRs.MoveNext
'        Loop
'    Else
'        cmdExport.Enabled = False
'        mSetResults "No agreements have been set up to use Marketron.  This form cannot be activated.", RGB(255, 0, 0)
'    End If
'End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
        If lbcVehicles.ListCount > 1 Then
            edcTitle3.Visible = False
            chkAllStation.Visible = False
            lbcStation.Visible = False
            lbcStation.Clear
        Else
            edcTitle3.Visible = True
            chkAllStation.Visible = True
            lbcStation.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub chkAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
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

End Sub

Private Sub cmdExport_Click()
    '6911 removed
   ' mCleanAet
    imExportMode = 0
    mExport
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload FrmExportMarketron
End Sub

Private Sub cmdExportTest_Click()
    imExportMode = 1
    mExport
End Sub


Private Sub CSI_Calendar1_Change()
    '8162
    tmcDelay.Enabled = True
End Sub

Private Sub CSI_Calendar1_Validate(Cancel As Boolean)
    '8162
    tmcDelay_Timer
'    If LenB(CSI_Calendar1.Text) = 0 Then
'        gMsgBox "Date must be set"
'        'CSI_Calendar1.SetFocus
'    ElseIf Weekday(CSI_Calendar1.Text) <> vbMonday Then
'        CSI_Calendar1.Text = gObtainPrevMonday(CSI_Calendar1.Text)
'    End If
'    '8162
'    If lbcStation.Visible Then
'        lbcVehicles_Click
'    End If
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        udcCriteria.Left = Label1.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = CSI_Calendar1.Top + (3 * CSI_Calendar1.Height) / 4
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcStation.Clear
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
                End If
            Next ilLoop
            chkAll.Value = vbChecked
            If lbcVehicles.ListCount = 1 Then
                imVefCode = lbcVehicles.ItemData(0)
                edcTitle3.Visible = True
                chkAllStation.Visible = True
                chkAllStation.Value = vbUnchecked
                lbcStation.Visible = True
                mFillStations
                chkAllStation.Value = vbChecked
            End If
        End If
        If igTestSystem Then
            udcCriteria.MOutput(0, "E") = False
            udcCriteria.MOutput(0, "V") = vbUnchecked
            udcCriteria.MOutput(1, "V") = vbChecked
        End If
        If igExportSource = 2 Then
            slNowStart = gNow()
            CSI_Calendar1.Text = sgExporStartDate
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "MarketronResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "Marketron Result List, Started: " & slNowStart
            ' pass global so glogMsg will write messages to sgExportResultName
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "MarketronResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "Marketron Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "Marketron Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            '6394 clear values
            hgExportResult = 0
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If

End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts FrmExportMarketron
    gCenterForm FrmExportMarketron
    mAdjustForm
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub
Private Sub mAdjustForm()
    Dim llLeft As Long
    
    llLeft = lbcVehicles.Left
    'ckcOutput(0).Left = llLeft
    Label1.Left = llLeft
    chkAll.Left = llLeft
    'ckcOutput(1).Left = CSI_Calendar1.Left
    udcCriteria.Left = llLeft
End Sub
Private Sub Form_Load()
    mInit
End Sub
Private Sub mInit()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    bmStationNotExist = False
    cmdExportTest.Visible = False
    Screen.MousePointer = vbHourglass
    FrmExportMarketron.Caption = "Export Marketron - " & sgClientName
    smStartDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    CSI_Calendar1.Text = smStartDate
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    smIniPath = gXmlIniPath(True)
    Set myErrors = New CLogger
    myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEERROR)
    myErrors.CleanThisFolder = messages
    smPathForgLogMsg = FILEERROR & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    If LenB(smIniPath) = 0 Then
        cmdExport.Enabled = False
        mSetResults "Xml.ini doesn't exist.  This form cannot be activated.", MESSAGERED
        myErrors.WriteWarning "Xml.ini doesn't exit. Export Marketron cannot be activated."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '6234
    If Not mSetExportClass(smIniPath) Then
        cmdExport.Enabled = False
        mSetResults "Xml.ini has no values for Marketron, or values cannot be read.  This form cannot be activated.", MESSAGERED
        myErrors.WriteWarning "Xml.ini has no values for Marketron or the values cannot be read."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    lbcStation.Clear
    mFillVehicle
    'limit to those vehicles that have an agreement that is Marketron...may be too slow.
    'mFillMarketronVehicle
    chkAll.Value = vbChecked
    Screen.MousePointer = vbDefault
    'csi internal guide-for testing help
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuGuide.Visible = True
    End If
End Sub
Private Function mSetExportClass(slIniPath As String) As Boolean
    'return true if values exist in ini file, not if created myExport
    Dim slRet As String
    Dim blRet As Boolean
    Dim slServicePage As String
    Dim slHost As String
    Dim slPassword As String
    Dim slUserName As String
'    Dim myXml As MSXML2.DOMDocument
'    Dim myElem As MSXML2.IXMLDOMElement
    
    blRet = False
    slUserName = ""
    slPassword = ""
On Error GoTo ERRORBOX
    'not sure if this code will work with proxy; so if proxy set, don't run.
    gLoadFromIni "MARKETRON", "ProxyServer", slIniPath, slRet
    If slRet <> "Not Found" Then
        mSetExportClass = True
        Exit Function
    End If
    gLoadFromIni "MARKETRON", "Host", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetExportClass = blRet
        Exit Function
    End If
    slHost = slRet
    gLoadFromIni "MARKETRON", "WebServiceURL", slIniPath, slRet 'import: WebServiceRcvURL
    If slRet = "Not Found" Then
        slRet = ""
    End If
'    If Len(slRet) = 0 Then
'        mSetExportClass = blRet
'        Exit Function
'    End If
    slServicePage = slRet
    gLoadFromIni "MARKETRON", "Authentication", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetExportClass = blRet
        Exit Function
    End If
    blRet = True
    '7878
'    Set myXml = New MSXML2.DOMDocument
'    If Not myXml.loadXML(slRet) Then
'        mSetExportClass = blRet
'        Exit Function
'    End If
'    Set myElem = myXml.selectSingleNode("//Username")
'    If Not myElem Is Nothing Then
'        slUserName = myElem.Text
'    End If
'    Set myElem = myXml.selectSingleNode("//Password")
'    If Not myElem Is Nothing Then
'        slPassword = myElem.Text
'    End If
    slUserName = gParseXml(slRet, "Username", 0)
    slPassword = gParseXml(slRet, "Password", 0)
    If Len(slPassword) > 0 And Len(slUserName) > 0 Then
        Set omExport = New CMarketron
        With omExport
            If StrComp(slHost, "Test", vbTextCompare) = 0 Then
                .isTest = True
            End If
                'export needs either this or for reading stations.  needs 'webserviceurl' to send exports.
               ' slServicePage = "/mx/orders/OrderServices.asmx"
            .SoapUrl = slHost '& slServicePage
            .ExportPage = slServicePage
            'couldn't set address
            If Len(.ErrorMessage) > 0 Then
                mSetResults "Couldn't set secondary calls to Marketron", MESSAGEBLACK
                myErrors.WriteWarning "Couldn't set secondary calls to Marketron: " & omExport.ErrorMessage & ".  Export will continue."
                Set omExport = Nothing
                GoTo Cleanup
            End If
            .Password = slPassword
            .UserName = slUserName
            'don't create debug file for just station calls
            '.LogPath = .CreateLogName(sgMsgDirectory & FILEDEBUG)
        End With
    End If
Cleanup:
'    Set myElem = Nothing
'    Set myXml = Nothing
    mSetExportClass = blRet
    Exit Function
ERRORBOX:
    blRet = False
    myErrors.WriteError "mSetExportClass-" & Err.Description
    GoTo Cleanup
End Function
Private Sub mCleanAet()
    Dim Sql As String
    Dim slDate As String
    Const MONTHSTOSAVE As Integer = 8
    
    slDate = DateAdd("m", -MONTHSTOSAVE, Date)
    slDate = Format(slDate, sgSQLDateForm)
    Sql = "Delete from aet where aetStatus = 'M' and aetFeedDate < '" & slDate & "'"
    If gSQLWaitNoMsgBox(Sql, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHandler:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mCleaAet"
        Exit Sub
    End If
    Exit Sub
    
ErrHandler:
    gHandleError smPathForgLogMsg, FORMNAME & "-mCleaAet"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    Erase tmCPDat
    Erase tmAstInfo
    rsFacts.Close
    Set myErrors = Nothing
    Set FrmExportMarketron = Nothing
End Sub
Private Sub lbcStation_Click()
    If imAllStationClick Then
        Exit Sub
    End If
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdExportTest.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
End Sub

Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcStation.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdExportTest.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        chkAllStation.Value = vbUnchecked
    End If
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        imAllClick = True
        chkAll.Value = vbUnchecked
        imAllClick = False
    End If
    For iLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(iLoop) Then
            imVefCode = lbcVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    If iCount = 1 Then
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    Else
        edcTitle3.Visible = False
        chkAllStation.Visible = False
        lbcStation.Visible = False
    End If
End Sub
Private Function mIsCsiLoggedIn() As Boolean
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If bgLimitedGuide Then
            mIsCsiLoggedIn = False
        Else
            mIsCsiLoggedIn = True
        End If
    End If
End Function
Private Function mExportSpots() As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim blOkStation As Boolean
    Dim blOkVehicle As Boolean
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilVpf As Integer
    Dim slVehicleName As String
    Dim slStationName As String
    Dim llTotalExport As Long
    Dim llSpotExport As Long
    Dim blAtLeastOneSpot As Boolean
    Dim blGrabFromAdv As Boolean
    Dim slProductCodeFromAdv As String
    Dim slResults(FIELDCOUNT - 1) As String
    Dim blAtLeastOneStation As Boolean
    '3/28/12  ttp 5263
    Dim blWarnNoSister As Boolean
    '6003
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilAllowedDays(0 To 6) As Integer
    '8/1/14: Compliant not required
    'Dim ilCompliant As Integer
    '5050
    Dim slAdv As String
    '6423
    Dim blErrorShown As Boolean
    '6633
    Dim slError As String
    '7458
    Dim myEnt As CENThelper
    Dim slEntSuccess As StatusEnum
    '8173
    Dim ilResendCount As Integer
    Dim slProp As String
    '9589
    Dim myRemapper As cRemapper
    
    On Error GoTo ErrHand
    
    Set myRemapper = New cRemapper
    '9851
    'myRemapper.Start
    myRemapper.StartRemapping
    myRemapper.isImport = False

    blErrorShown = False
    bmAtLeastOneExport = False
    llTotalExport = 0
    imExporting = True
    imVefCode = 0
    '6003
    gPopDaypart
    'ttp 5050
    ' this is same as smoutputype <> 'T'.  They chose 'generate file'
    If udcCriteria.MOutput(1, slProp) = vbChecked Then
        Set rsFacts = mPrepRecordsetDebug()
    End If
'6554 remove testing station!
'    If udcCriteria.MOutput(0) = vbChecked And Not omExport Is Nothing Then
'        lacResult.Caption = "Gathering Station Information from Marketron."
'        If Not omExport.GetStations() Then
'            myErrors.WriteError "Couldn't check stations: " & omExport.ErrorMessage
'            Set omExport = Nothing
'        End If
'    End If
    'just 1 vehicle chosen?
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            If imVefCode = 0 Then
                imVefCode = lbcVehicles.ItemData(ilVef)
            Else
                imVefCode = -1
                Exit For
            End If
        End If
    Next ilVef
    lacResult.Caption = "Beginning export."
    blAtLeastOneSpot = False
    If igExportSource = 2 Then DoEvents
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .TypeEnt = Exportunposted3rdparty
        .ThirdParty = Vendors.NetworkConnect
        .User = igUstCode
        .ErrorLog = smPathForgLogMsg
    End With
    '3/28/12  ttp 5263 add attMulticast
    '7701
    SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, shttClusterGroupId, shttMasterCluster, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attTimeType, attMultiCast "
    '6/24/11 Dan add for multicasting
   ' SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attTimeType" 'attdropdate,attOffAir,attOnAir, attGenCP, cpttStartDate,
   ' SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, shttStationID, shttClusterGroupId, shttMasterCluster, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attTimeType" 'attdropdate,attOffAir,attOnAir, attGenCP, cpttStartDate,
    SQLQuery = SQLQuery & " FROM shtt, cptt, vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode "
    SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
    SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
    SQLQuery = SQLQuery & " AND vefCode = cpttVefCode"
    '10/29/14: Bypass Service agreements
    SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.NetworkConnect
    If imVefCode > 0 Then
        SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
    End If
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(smStartDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
    Set cprst = gSQLSelectCall(SQLQuery)
    While Not cprst.EOF
        If igExportSource = 2 Then DoEvents
        If lbcStation.ListCount > 0 Then
            blOkStation = False
            For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If lbcStation.Selected(ilLoop) Then
                    If lbcStation.ItemData(ilLoop) = cprst!shttCode Then
                        blOkStation = True
                        Exit For
                    End If
                End If
            Next ilLoop
        Else
            blOkStation = True
        End If
        If blOkStation Then
            blOkVehicle = False
            For ilVef = 0 To lbcVehicles.ListCount - 1
                If igExportSource = 2 Then DoEvents
                If lbcVehicles.Selected(ilVef) Then
                    If lbcVehicles.ItemData(ilVef) = cprst!cpttvefcode Then
                        '12/11/17: Clear abf
                        If (imVefCode <> lbcVehicles.ItemData(ilVef)) And (imVefCode > 0) Then
                            If (lbcStation.ListCount <= 0) Or (chkAllStation.Value = vbChecked) Then
                                'gClearAbf imVefCode, 0, smStartDate, gObtainNextSunday(smStartDate)
                            End If
                        End If
                        imVefCode = lbcVehicles.ItemData(ilVef)
                        blOkVehicle = True
                        Exit For
                    End If
                End If
            Next ilVef
        End If
        If blOkStation And blOkVehicle Then
            '6/23/11 Dan multicasting
            '7701-made global
            If Not gSlave(cprst!shttclustergroupId, cprst!shttMasterCluster, cprst!attMulticast, blWarnNoSister, smPathForgLogMsg) Then
'            If Not mSlave(cprst!shttClusterGroupID, cprst!shttMasterCluster, cprst!attMulticast, blWarnNoSister) Then
                '5349
                slStationName = mGetStationName(cprst!shttCode)
                '6554 remove testing station!
                'must be sending to marketron
'                If Not omExport Is Nothing And udcCriteria.MOutput(0) = vbChecked Then
'                    If Not omExport.TestStation(slStationName) Then
'                        If Len(omExport.ErrorMessage) > 0 Then
'                            myErrors.WriteWarning "Issue in TestStation: " & omExport.ErrorMessage
'                            'error? then don't try to test stations in future
'                            Set omExport = Nothing
'                        Else
'                            blOkStation = False
'                            bmStationNotExist = True
'                            myErrors.WriteWarning slStationName & " does not exist on Marketron Network and will not be sent."
'                        End If
'                    End If
'                End If
                If blOkStation Then
                    llSpotExport = 0
                    ilVpf = gBinarySearchVpf(CLng(imVefCode))
                    If ilVpf <> -1 Then
                        On Error GoTo ErrHand
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                        tgCPPosting(0).iStatus = cprst!cpttStatus
                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                        tgCPPosting(0).iVefCode = imVefCode
                        tgCPPosting(0).iShttCode = cprst!shttCode
                        tgCPPosting(0).sZone = cprst!shttTimeZone
                        tgCPPosting(0).sDate = Format$(smStartDate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                        'Create AST records
                        igTimes = 1 'By Week
                        imAdfCode = -1
                        If igExportSource = 2 Then DoEvents
                       '7458
                        slVehicleName = mGetVehicleName(imVefCode) 'tmAstInfo(ilIndex).iVefCode
                        With myEnt
                            .Vehicle = imVefCode
                            .Station = cprst!shttCode
                            .Agreement = cprst!cpttatfCode
                            .fileName = mCreateXmlFileName(slVehicleName, slStationName, imVefCode)
                            .ProcessStart
                        End With
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                        '3/10/16: Remove MG; Replacements and Bonus
                        gFilterAstExtendedTypes tmAstInfo
                        ilIndex = LBound(tmAstInfo)
                        slVehicleName = mGetVehicleName(imVefCode) 'tmAstInfo(ilIndex).iVefCode
                        '6423
                        If ilIndex = UBound(tmAstInfo) Then
                            If Not blErrorShown Then
                                mSetResults "Spots Missing, see MarketronExportLog", MESSAGERED
                                blErrorShown = True
                            End If
                            myErrors.WriteWarning "Spots missing for: " & slStationName & " " & slVehicleName
                        Else
                            lacResult.Caption = "Exporting " & slStationName & ", " & slVehicleName
                            '6640
                            myErrors.WriteFacts "Exporting " & slStationName & ", " & slVehicleName
                        End If
 '                      Replaced by 6423
'                        'Dan using demo database, gGetAstInfo would sometimes return nothing, causing error later
'                        If ilIndex = UBound(tmAstInfo) Then
'                            mSetResults "Couldn't build spots.", RGB(255, 0, 0)
'                            mExportSpots = 0
'                            Exit Function
'                        End If
'                        slVehicleName = mGetVehicleName(tmAstInfo(ilIndex).iVefCode)
'                       ' slStationName = mGetStationName(tmAstInfo(ilIndex).iShttCode)
'                        '5868
'                        'Call mSetResults("Exporting " & slStationName & ", " & slVehicleName, 0)
'                        lacResult.Caption = "Exporting " & slStationName & ", " & slVehicleName
                        
                        'loop all spots
                        Do While ilIndex < UBound(tmAstInfo)
                            With tmAstInfo(ilIndex)
                                If igExportSource = 2 Then DoEvents
                                '6003
                                '8/1/14: Compliant not required
                                'gGetLineParameters False, tmAstInfo(ilIndex), slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), ilCompliant
                                '7470 marketron may send pledge dates
                               ' gGetLineParameters False, tmAstInfo(ilIndex), slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays()
                                gGetLineParameters False, tmAstInfo(ilIndex), slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), True
                            ' this date control unneeded: always just a week's worth of spots, and we always export that.
                            ' continue to use feed date here -Dan and Dick 3/14/13
                                If (gDateValue(gAdjYear(.sFeedDate)) >= gDateValue(gAdjYear(smStartDate))) And (gDateValue(gAdjYear(.sFeedDate)) <= gDateValue(gAdjYear(smEndDate))) Then
                                    ' 2 means don't air!
                                    If (tgStatusTypes(gGetAirStatus(.iStatus)).iPledged <> 2) Then
                                        If Not blAtLeastOneSpot Then
                                            blAtLeastOneSpot = True
                                            If Not mStartXmlFile(slVehicleName, Trim$(cprst("shttCallLetters")), .iVefCode) Then
                                                gMsgBox "Problem in mStartXmlFile. ", vbCritical
                                                mExportSpots = False
                                                '7458
                                                GoTo Cleanup
                                                'Exit Function
                                            End If
                                        End If
                                        '7458
                                        If Not myEnt.Add(.sFeedDate, .lgsfCode) Then
                                            myErrors.WriteWarning myEnt.ErrorMessage
                                        End If
                                        llSpotExport = llSpotExport + 1
                                        '9589
                                        myRemapper.ExportingDate = CDate(slStartDate)
                                        slResults(SPOTID) = myRemapper.Remap(.lCode)
                                        ' slResults(SPOTID) = Trim$(.lCode)
                                        slResults(startDate) = Format(slStartDate, XMLDATE)
                                        slResults(endDate) = Format(slEndDate, XMLDATE)
                                        'Dan M 8/02/11 endtime sometimes blank
                                        slResults(startTime) = Format(slStartTime, "HH:mm:ss")
                                        If slResults(startTime) = "00:00:00" Then
                                            slResults(startTime) = "00:00:01"
                                        End If
                                        slResults(endTime) = Format(slEndTime, "HH:mm:ss")
                                        If Len(Trim(slResults(endTime))) = 0 Then
                                            slResults(endTime) = slResults(startTime)
                                        ElseIf slResults(endTime) = "00:00:00" Then
                                            slResults(endTime) = "23:59:59"
                                        End If
                                        slResults(LENGTH) = Trim$(tmAstInfo(ilIndex).iLen)
                                        '0 =none, 1=split copy, 2=blackout
                                        If tmAstInfo(ilIndex).iRegionType <> 0 Then
                                            slResults(PRODUCTNAME) = gXMLNameFilter(.sRProduct)
                                            slResults(ISCICODE) = gXMLNameFilter(.sRISCI)
                                            slResults(COPYTITLE) = gXMLNameFilter(.sRCreativeTitle)
                                            slResults(COPYCOMMENT) = gXMLNameFilter(mGetCSFComment(.lRCrfCsfCode))
                                        Else
                                            slResults(PRODUCTNAME) = gXMLNameFilter(.sProd)
                                            slResults(ISCICODE) = gXMLNameFilter(.sISCI)
                                            slResults(COPYTITLE) = gXMLNameFilter(mCreativeTitle(.lCpfCode))
                                            slResults(COPYCOMMENT) = gXMLNameFilter(mGetCSFComment(.lCrfCsfCode))
                                        End If
                                        'couldn't find product code in cif or chf? use adv.  Couldn't find copyenddate? use end date of file
                                        blGrabFromAdv = mCifInfo(ilIndex, slResults(PRODUCTCODE), slResults(COPYENDDATE))
                                        If Len(slResults(COPYENDDATE)) = 0 Then
                                            slResults(COPYENDDATE) = Format(smEndDate, XMLDATE)
                                        End If
                                        '5050
                                        slAdv = mAdvName(ilIndex, blGrabFromAdv, slProductCodeFromAdv)
                                        slResults(Advertiser) = gXMLNameFilter(slAdv)
                                        'slResults(Advertiser) = mAdvName(ilIndex, blGrabFromAdv, slProductCodeFromAdv)
                                        If blGrabFromAdv Then
                                            slResults(PRODUCTCODE) = gXMLNameFilter(slProductCodeFromAdv)
                                        End If
                                        mCSIXMLData "OT", "Spot", vbNullString
                                        For ilLoop = 0 To FIELDCOUNT - 1
                                            mCSIXMLData "CD", smFields(ilLoop), slResults(ilLoop)
                                        Next ilLoop
                                        mCSIXMLData "CT", "Spot", vbNullString
                                        'ttp 4922 removed 6911
'                                        If smOutputType <> "F" Or mIsCsiLoggedIn() Then
'                                            'Sql error on inputting to aet will STOP export and, after msgbox, will say 'user terminated'
'                                             If Not mStoreBackup(ilIndex) Then
'                                                imTerminate = True
'                                             End If
'                                        End If
                                        'ttp 5050
                                        ' this is same as smoutputype <> 'T'.  They chose 'generate file'
                                        If udcCriteria.MOutput(1, slProp) = vbChecked Then
                                                rsFacts.AddNew Array("Station", "Vehicle", "AstCode", "ISCI", "advertiser", "Date", "Time"), Array(slStationName, slVehicleName, slResults(SPOTID), slResults(ISCICODE), slAdv, slResults(startDate), slResults(startTime))
                                        End If
                                    '7458
                                    ElseIf Not myEnt.Add(.sFeedDate, .lgsfCode, Asts) Then
                                        myErrors.WriteWarning myEnt.ErrorMessage
                                    End If  'include spot
                                End If  'if date is ok
                            End With
                            ilIndex = ilIndex + 1
                            If imTerminate Then
                                imExporting = False
                                Exit Function
                            End If
                        Loop 'each spot
                        llTotalExport = llTotalExport + llSpotExport
                        If blAtLeastOneSpot Then
                            '5868
'                            If smOutputType <> "F" Then
'                                mSetResults "   Sending to Marketron", 0
'                            End If
                            If Not mCloseXml() Then
                                '8173
                                For ilResendCount = 1 To 2
                                    myErrors.WriteWarning "Could not send.  Attempting resend #" & ilResendCount, True
                                    If mResendXml() Then
                                        blAtLeastOneStation = True
                                        '7458
                                        slEntSuccess = Successful
                                        If smOutputType <> "F" Then
                                            ilRet = gUpdateLastExportDate(imVefCode, smEndDate)
                                        End If
                                        Exit For
                                    End If
                                    If ilResendCount = 2 Then
                                        slError = mGetXmlError()
                                        ' set to -1 to know that an error occurred, not that no spots needed to be exported
                                        llTotalExport = -1
                                        mExportSpots = False
                                        Call mSetResults("Failed to send for " & slStationName & " and " & slVehicleName, RGB(255, 0, 0))
                                        myErrors.WriteWarning "Failed to send for " & slStationName & " and " & slVehicleName & vbCrLf & "  " & slError
                                        '7458
                                        slEntSuccess = EntError
                                    End If
                                Next ilResendCount
                            Else
                                blAtLeastOneStation = True
                                '7458
                                slEntSuccess = Successful
                                If smOutputType <> "F" Then
                                    '5868
                                    'mSetResults "   Sent to Marketron: " & llSpotExport & " spots.", RGB(0, 0, 0), True
                                    ilRet = gUpdateLastExportDate(imVefCode, smEndDate)
                                End If
                            End If
                            '8173 moved from 'mCloseXml'
                            csiXMLEnd
                            '7458
                            If smOutputType <> "F" Or omExport.isTest Then
                                If Not myEnt.CreateEnts(slEntSuccess) Then
                                    myErrors.WriteWarning myEnt.ErrorMessage
                                End If
                            End If
                        '7458
                        Else
                            myEnt.ClearWhenDontSend
                        End If 'at least one spot
                        For ilLoop = 0 To FIELDCOUNT - 1
                            slResults(ilLoop) = vbNullString
                        Next ilLoop
                    End If  'found vpf record
                End If '5349 safe station to send
                If igExportSource = 2 Then DoEvents
            Else
                If blWarnNoSister Then
                    mSetResults Trim$(cprst!shttCallLetters) & " is not set to have a master station, but the agreement is set as multicast.  Not exported.", RGB(255, 0, 0)
                Else
                    '5868
                   ' mSetResults Trim$(cprst!shttCallLetters) & " is a non-master station and will not be exported.", RGB(0, 155, 0)
                    lacResult.Caption = Trim$(cprst!shttCallLetters) & " is a non-master station and will not be exported."
                End If
            End If  'not slave
        End If  'station and vehicle...xml single file
        cprst.MoveNext
        blAtLeastOneSpot = False
    Wend
    If imTerminate Then
        imExporting = False
        Exit Function
    End If
    
    '12/11/17: Clear Abf
    If (lbcStation.ListCount <= 0) Or (chkAllStation.Value = vbChecked) Then
        'gClearAbf imVefCode, 0, smStartDate, gObtainNextSunday(smStartDate)
    End If
    
    If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
        gClearASTInfo True
    Else
        gClearASTInfo False
    End If
'    If llTotalExport > 0 Then
'        Call mSetResults("Total Spots Exported = " & llTotalExport, 0, True)
'        bmAtLeastOneExport = True
'    Else
'        Call mSetResults("No Spots needed to be Exported.", RGB(0, 155, 0))
'        bmAtLeastOneExport = False
'    End If
    ' dan 6/21/11 for multicasting--use ignore llTotalExport = -1
    If llTotalExport > 0 Then
        Call mSetResults("Total Spots Exported = " & llTotalExport, 0, False)
        ' at least one went through...use for writing a message later.
        bmAtLeastOneExport = True
    ElseIf llTotalExport = 0 Then
        Call mSetResults("No Spots needed to be Exported", RGB(0, 155, 0))
    ElseIf blAtLeastOneStation Then
        bmAtLeastOneExport = True
    End If
    If udcCriteria.MOutput(0, slProp) = vbChecked Then
        mClearAlerts gDateValue(smStartDate), gDateValue(smEndDate)
    End If
    mExportSpots = True
    'ttp 5050
    If udcCriteria.MOutput(1, slProp) = vbChecked Then
        If mWriteFacts() Then
            'send to caption?
        End If
    End If
Cleanup:
    If Not rsFacts Is Nothing Then
        If (rsFacts.State And adStateOpen) <> 0 Then
            rsFacts.Close
        End If
        Set rsFacts = Nothing
    End If
    Set myEnt = Nothing
Exit Function
mExportSpotsErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mExportSpots"
    mExportSpots = False
    GoTo Cleanup
    Exit Function
 End Function
Private Function mStoreBackup(ilIndex As Integer) As Boolean
    '4922
    Dim Sql As String
    Dim myRs As ADODB.Recordset
    Dim blRet As Boolean
    
    blRet = True
'    gPackDate tmAstInfo(ilIndex).sFeedDate, iFeedDate(0), iFeedDate(1)
'    gPackTime tmAstInfo(ilIndex).sFeedTime, iFeedTime(0), iFeedTime(1)
    Sql = " SELECT aetCode FROM  aet WHERE aetAstCode = " & tmAstInfo(ilIndex).lCode
    On Error GoTo ErrHandler
    Set myRs = gSQLSelectCall(Sql)
    If Not myRs.EOF Then
       ' MsgBox "Update"
       'update fails because index key for feed date set as non-modifiable
       ' Sql = " UPDATE AET set aetStatus = 'N', aetFeedDate = " & Format(tmAstInfo(ilIndex).sFeedDate, sgSQLDateForm) & " , aetFeedTime = " & Format(tmAstInfo(ilIndex).sFeedTime, sgSQLTimeForm) & " where aetcode = " & myRs.Fields("aetCode").Value
        Sql = " DELETE FROM AET where aetcode = " & myRs.Fields("aetCode").Value
        If gSQLWaitNoMsgBox(Sql, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHandler:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, FORMNAME & "-mStoreBackup"
            blRet = False
            GoTo Cleanup
        End If
    End If
       ' MsgBox "insert"
    Sql = " INSERT into aet (aetStatus,aetAstCode,aetFeedDate,aetFeedTime) VALUES ( 'M', " & tmAstInfo(ilIndex).lCode & ", '" & Format(tmAstInfo(ilIndex).sFeedDate, sgSQLDateForm) & "', '" & Format(tmAstInfo(ilIndex).sFeedTime, sgSQLTimeForm) & "' )"
    If gSQLWaitNoMsgBox(Sql, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHandler:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mStoreBackup"
        blRet = False
        GoTo Cleanup
    End If
Cleanup:
     mStoreBackup = blRet
     If Not myRs Is Nothing Then
         If (myRs.State And adStateOpen) <> 0 Then
             myRs.Close
        End If
        Set myRs = Nothing
    End If
    Exit Function
ErrHandler:
    gHandleError smPathForgLogMsg, FORMNAME & "-mStoreBackup"
    blRet = False
    GoTo Cleanup
End Function
Private Function mCloseXml() As Boolean
    '6633
    If myErrors.myFile.FILEEXISTS(smXmlErrorFile) Then
On Error GoTo ERRORNODELETE
        'dantest must allow
        myErrors.myFile.DeleteFile smXmlErrorFile, True
    End If
    mCSIXMLData "CT", "SpotList", ""
    mCSIXMLData "CT", "NetworkOrder", ""
    mCSIXMLData "CT", "Payload", ""
    mCSIXMLData "CT", "ProcessOrderNetworkRequest", ""
    mCSIXMLData "CT", "ProcessOrder", ""
    mCloseXml = csiXMLWrite(1)
    'dantest must remove!
   ' mCloseXml = False
    '8173 remove
    'csiXMLEnd
    Exit Function
ERRORNODELETE:
    myErrors.WriteWarning "Warning- could not delete xml error file " & Err.Description
    Resume Next
End Function
Private Function mResendXml() As Boolean
    If myErrors.myFile.FILEEXISTS(smXmlErrorFile) Then
On Error GoTo ERRORNODELETE
        myErrors.myFile.DeleteFile smXmlErrorFile, True
    End If
    mResendXml = csiXMLResend(1)
    'dantest must remove!
   ' mResendXml = False
    Exit Function
ERRORNODELETE:
    myErrors.WriteWarning "Warning- could not delete xml error file " & Err.Description
    Resume Next
End Function
Private Function mCreativeTitle(llCpfCode As Long) As String
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    
On Error GoTo ErrHandler
    If llCpfCode > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select cpfCreative FROM cpf_copy_prodct_isci WHERE cpfCode = " & llCpfCode
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            mCreativeTitle = Trim$(Format(myRs!cpfCreative))
        End If
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
            Set myRs = Nothing
        End If
    End If
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mCreativeTitle"
    mCreativeTitle = vbNullString
    GoTo Cleanup
End Function
Private Function mAdvName(ilIndex As Integer, blGetFromAdv As Boolean, slAdvProductCode As String) As String
    Dim ilAdf As Integer
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    Dim ilAdfComp As Integer
    
On Error GoTo ErrHandler
    slAdvProductCode = vbNullString
    ilAdf = tmAstInfo(ilIndex).iAdfCode
    If ilAdf > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select adfName, adfmnfcomp1,adfmnfcomp2 FROM adf_advertisers WHERE adfCode = " & ilAdf
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            If igExportSource = 2 Then DoEvents
           ' mAdvName = gXMLNameFilter(Format(myRs!adfname))
            '5050
            mAdvName = myRs!adfName
            If blGetFromAdv Then
                If Format(myRs!adfmnfcomp1) > 0 Then
                    ilAdfComp = Format(myRs!adfmnfcomp1)
                ElseIf Format(myRs!adfmnfcomp2) > 0 Then
                    ilAdfComp = Format(myRs!adfmnfcomp2)
                End If
                If ilAdfComp > 0 Then
                    myRs.Close
                   Sql = "Select mnfName FROM mnf_Multi_Names WHERE mnfCode = " & ilAdfComp
                    Set myRs = gSQLSelectCall(Sql)
                    If Not (myRs.EOF Or myRs.BOF) Then
                        slAdvProductCode = Trim$(Format(myRs!mnfName))
                    End If
                End If
            End If
        End If
    Else
        mAdvName = vbNullString
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
            Set myRs = Nothing
        End If
    End If
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mAdvName"
    mAdvName = vbNullString
    GoTo Cleanup
End Function
Private Function mOpenCSF() As Boolean

    Dim ilRet As Integer
    
    hmCsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCsf, "", sgDBPath & "CSF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on CSF.BTR"
        ilRet = btrClose(hmCsf)
        btrDestroy hmCsf
        mOpenCSF = False
        Exit Function
    End If
    
    mOpenCSF = True

End Function
Private Sub mCloseCSF()

    Dim ilRet As Integer
    If bmCsfOpenStatus Then
        ilRet = btrClose(hmCsf)
        If ilRet <> BTRV_ERR_NONE Then
            gMsgBox "btrClose Failed on CSF.BTR"
        End If
        btrDestroy hmCsf
    End If
End Sub

Private Function mGetCSFComment(lCSFCode As Long) As String
    Dim ilRet As Integer, i As Integer, ilLen As Integer, ilActualLen As Integer
    Dim ilRecLen As Integer
    Dim tlCSF As CSF
    Dim tlCsfSrchKey As LONGKEY0
    Dim slComment As String
    Dim slTemp As String
    Dim blOneChar As Byte
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    mGetCSFComment = ""
    If (lCSFCode <= 0) Or (bmCsfOpenStatus = False) Then
        Exit Function
    End If
    tlCsfSrchKey.lCode = lCSFCode
    tlCSF.sComment = ""
    ilRecLen = Len(tlCSF) '5011
    ilRet = btrGetEqual(hmCsf, tlCSF, ilRecLen, tlCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If
    If igExportSource = 2 Then DoEvents
    slComment = gStripChr0(tlCSF.sComment)
    If slComment <> "" Then
        ' Strip off any trailing non ascii characters.
        ilLen = Len(slComment)
        ' Find the first valid ascii character from the end and assume the rest of the string is good.
        For i = ilLen To 1 Step -1
            If igExportSource = 2 Then DoEvents
            blOneChar = Asc(Mid(slComment, i, 1))
            If blOneChar >= 32 Then
                ' The first valid ASCII character has been found.
                slTemp = Left(slComment, i)
                Exit For
            End If
        Next i
        ilActualLen = i
        ' Scan through and remove any non ASCII characters. This was causing a problem for the web site.
        slComment = ""
        For i = 1 To ilActualLen
            If igExportSource = 2 Then DoEvents
            blOneChar = Asc(Mid(slTemp, i, 1))
            If blOneChar >= 32 Then
                slComment = slComment + Mid(slTemp, i, 1)
            Else
                slComment = slComment + " "
            End If
        Next i
        mGetCSFComment = slComment
    End If

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in ExportMarketron-mGetCSFComment: "
        'gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "MarketronExportLog.Txt", False
        myErrors.WriteError "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, False, True
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Function
Private Function mCifInfo(ilIndex As Integer, slProductCode As String, slRotEndDate As String) As Boolean
'returns if productCode should come from adv table
'rule: get mnfName from cif unless it is 0 or cif is 0.  Then use sdfcode to chfCode to get mnfName (except for blackout). If that fails also, or is blackout
' go to adv
    Dim llCif As Long
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    Dim llSdf As Long
    Dim ilCifComp As Integer
    Dim ilChfComp As Integer
    Dim slRet As String
    
    slRotEndDate = vbNullString
    slProductCode = vbNullString
On Error GoTo ErrHandler
    mCifInfo = False
    If tmAstInfo(ilIndex).iRegionType = 0 Then
        llCif = tmAstInfo(ilIndex).lCifCode
    Else
        llCif = tmAstInfo(ilIndex).lRCifCode
    End If
    ilCifComp = 0
    If llCif > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select cifmnfcomp2, cifmnfcomp1,cifRotEndDate FROM cif_Copy_Inventory WHERE cifCode = " & llCif
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            If Format(myRs!cifmnfcomp1) > 0 Then
                ilCifComp = Format(myRs!cifmnfcomp1)
            ElseIf Format(myRs!cifmnfcomp2) > 0 Then
                ilCifComp = Format(myRs!cifmnfcomp2)
            End If
            slRotEndDate = Format(myRs!cifRotEndDate, XMLDATE)
        End If
        If igExportSource = 2 Then DoEvents
        myRs.Close
    End If
    If ilCifComp > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select mnfName FROM mnf_Multi_Names WHERE mnfCode = " & ilCifComp
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            slProductCode = Trim$(Format(myRs!mnfName))
        End If
        myRs.Close
    'failed to get from cif. blackout done, other get from sdf-chf-mnf
    ElseIf tmAstInfo(ilIndex).iRegionType < 2 Then
        llSdf = tmAstInfo(ilIndex).lSdfCode
        ilChfComp = 0
        slRet = vbNullString
        If llSdf > 0 Then
            If igExportSource = 2 Then DoEvents
            Sql = "SELECT sdfchfCode FROM sdf_Spot_Detail where sdfCode = " & llSdf
            Set myRs = gSQLSelectCall(Sql)
            If Not (myRs.EOF Or myRs.BOF) Then
                If myRs(0) > 0 Then
                    Sql = "Select chfmnfcomp2, chfmnfcomp1 FROM chf_Contract_Header WHERE chfCode = " & myRs(0)
                    myRs.Close
                    Set myRs = gSQLSelectCall(Sql)
                    If Not (myRs.EOF Or myRs.BOF) Then
                        If Format(myRs!chfmnfcomp1) > 0 Then
                           ilChfComp = Format(myRs!chfmnfcomp1)
                        ElseIf Format(myRs!chfmnfcomp2) > 0 Then
                           ilChfComp = Format(myRs!chfmnfcomp2)
                        End If
                    End If
                    myRs.Close
                Else
                    myRs.Close
                End If
            Else
                myRs.Close
            End If
        End If
        If ilChfComp > 0 Then
            If igExportSource = 2 Then DoEvents
            Sql = "Select mnfName FROM mnf_Multi_Names where mnfcode = " & ilChfComp
            Set myRs = gSQLSelectCall(Sql)
            If Not (myRs.EOF Or myRs.BOF) Then
                slProductCode = Trim$(Format(myRs!mnfName))
            End If
        End If
    Else
        mCifInfo = True
    End If
    If LenB(slProductCode) > 0 Then
        slProductCode = gXMLNameFilter(slProductCode)
    Else
        mCifInfo = True
    End If

Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
            Set myRs = Nothing
        End If
    End If
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mCifInfo"
    mCifInfo = vbNullString
    GoTo Cleanup
End Function
Private Function mStartXmlFile(ByVal slVehicleName As String, slCallLetters As String, ilVefCode As Integer) As Boolean
    Dim slOrderId As String
    Dim slEnd As String
    Dim slfixedCallBand As String
    
    If imExportMode = 1 Then
        slEnd = Chr$(13) + Chr$(10)
    End If
    'for 7458
    slOrderId = mCreateXmlFileName(slVehicleName, slCallLetters, ilVefCode)
'    'change Kzzz-fm to Kzzzfm
    slfixedCallBand = Replace(slCallLetters, "-", "")
'    slOrderId = slVehicleName & "-" & ilVefCode & "-" & slfixedCallBand & "-" & Format$(smStartDate, "yyyymmdd")
'    slOrderId = mSafeFileName(slOrderId)
 On Error GoTo ERRWRITE
    If igExportSource = 2 Then DoEvents
    '6807
    'If csiXMLStart(smIniPath, "Marketron", smOutputType, sgExportDirectory & slOrderId, slEnd) <> 0 Then
    If csiXMLStart(smIniPath, "Marketron", smOutputType, sgExportDirectory & slOrderId, slEnd, smXmlErrorFile) <> 0 Then
        mStartXmlFile = False
        Exit Function
    End If
    'file name couldn't have '\', but order id does
   ' slOrderId = gXMLNameFilter(slVehicleName) & "\" & slfixedCallBand & "\" & Format$(smStartDate, "yyyymmdd")
    '5993
    slVehicleName = gXMLNameFilter(slVehicleName)
    slOrderId = mSafeFileName(slVehicleName) & "-" & ilVefCode & "\" & slfixedCallBand & "\" & Format$(smStartDate, "yyyymmdd")
    mWriteHeader slVehicleName, slCallLetters, slOrderId
    If igExportSource = 2 Then DoEvents
    mStartXmlFile = True
    Exit Function
ERRWRITE:
    mStartXmlFile = False
End Function
Public Function mCreateXmlFileName(ByVal slVehicleName As String, slCallLetters As String, ilVefCode As Integer) As String
    Dim slRet As String
    Dim slfixedCallBand As String
    
    'change Kzzz-fm to Kzzzfm
    slfixedCallBand = Replace(slCallLetters, "-", "")
    slRet = slVehicleName & "-" & ilVefCode & "-" & slfixedCallBand & "-" & Format$(smStartDate, "yyyymmdd")
    slRet = mSafeFileName(slRet)
    mCreateXmlFileName = slRet
End Function
Private Function mSafeFileName(slOldName As String) As String
    Dim slTempName As String
    If igExportSource = 2 Then DoEvents
    slTempName = Replace(slOldName, "?", "-")
    slTempName = Replace(slTempName, "/", "-")
    slTempName = Replace(slTempName, "\", "-")
    slTempName = Replace(slTempName, "%", "-")
    slTempName = Replace(slTempName, "*", "-")
    slTempName = Replace(slTempName, ":", "-")
    slTempName = Replace(slTempName, "|", "-")
    slTempName = Replace(slTempName, """", "-")
    slTempName = Replace(slTempName, ".", "-")
    slTempName = Replace(slTempName, "<", "-")
    slTempName = Replace(slTempName, ">", "-")
    If igExportSource = 2 Then DoEvents
    mSafeFileName = slTempName
End Function
Private Sub mWriteHeader(slVehicleName As String, slCallAndBand As String, slOrderId As String)
    Dim slAnswers() As String
    Dim slCall As String
    Dim slBand As String
    
    slAnswers = Split(slCallAndBand, "-")
    If UBound(slAnswers) = 1 Then
        slCall = slAnswers(0)
        slBand = slAnswers(1)
    Else
        slCall = slAnswers(0)
        slBand = ""
    End If
    Erase slAnswers
    If Len(slCall) > 10 Then
        slCall = Mid(slCall, 1, 10)
    End If
    If Len(slBand) > 6 Then
        slCall = Mid(Trim$(slBand), 1, 6)
    End If
    If igExportSource = 2 Then DoEvents
    mCSIXMLData "CD", "SchemaName", "External Network Order Schema"
    mCSIXMLData "CD", "SchemaVersion", "1.0.0.0"
    mCSIXMLData "CD", "OrderID", gXMLNameFilter(slOrderId)
    mCSIXMLData "CD", "NetworkName", mNetworkName()
    mCSIXMLData "CD", "NetworkSoftware", "Counterpoint"
    mCSIXMLData "CD", "ProgramName", gXMLNameFilter(slVehicleName)
    mCSIXMLData "CD", "CallLetters", gXMLNameFilter(slCall)
    mCSIXMLData "CD", "Band", slBand
    mCSIXMLData "CD", "StartDate", Format(smStartDate, XMLDATE)
    mCSIXMLData "CD", "EndDate", Format(smEndDate, XMLDATE)
    mCSIXMLData "OT", "SpotList", ""
    If igExportSource = 2 Then DoEvents
End Sub
Private Function mNetworkName() As String
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    Dim slMyString As String
 On Error GoTo ERRORHAND
    Sql = "Select spfGClient from SPF_Site_Options"
    Set myRs = gSQLSelectCall(Sql)
    If Not (myRs.EOF Or myRs.BOF) Then
        myRs.MoveFirst
        slMyString = Format(myRs!spfgClient)
        mNetworkName = gXMLNameFilter(slMyString)
    End If
    Exit Function
ERRORHAND:
    mNetworkName = vbNullString
End Function
Private Sub mFillStations()
    Dim slNowDate As String
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    '8162 gAdjYear(CSI_Calendar1.Text)
    slNowDate = Format(gAdjYear(CSI_Calendar1.Text), sgSQLDateForm)
    If Not IsDate(slNowDate) Then
        slNowDate = Format(gNow(), sgSQLDateForm)
    End If
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    'Dan M marketron only
    SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.NetworkConnect
    '8162
    SQLQuery = SQLQuery & " AND attonAir <= '" & slNowDate & "' AND attdropdate > '" & slNowDate & "' AND attoffair > '" & slNowDate & "'"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcStation.AddItem Trim$(rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
        rst.MoveNext
    Wend
    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mFillStations"
End Sub

Private Sub mSetResults(Msg As String, FGC As Long, Optional blRemovePrevious = False)
'   multicasting...if get an error, leave red! Use gAddmsgtolistbox to add scroll bar
'    If blRemovePrevious Then
'        lbcMsg.RemoveItem (lbcMsg.ListIndex)
'
'  '  Else
'     '   lbcMsg.AddItem "   Exporting spot #" & llCurrentSpot
'    End If
'    lbcMsg.AddItem Msg
'    lbcMsg.ListIndex = lbcMsg.ListCount - 1
'    lbcMsg.ForeColor = FGC
'    If igExportSource = 2 Then DoEvents
    If blRemovePrevious Then
        lbcMsg.RemoveItem (lbcMsg.ListIndex)
    End If
    'dan 6/15/2011 add vertical scroll bar as needed
    gAddMsgToListBox FrmExportMarketron, lmMaxWidth, Msg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    If lbcMsg.ForeColor <> RGB(255, 0, 0) Then
        lbcMsg.ForeColor = FGC
    End If
    If igExportSource = 2 Then DoEvents

End Sub

Private Function mGetVehicleName(iVefCode As Integer) As String
    Dim llLoop As Integer
    mGetVehicleName = ""
    For llLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(llLoop).iCode = iVefCode Then
            mGetVehicleName = Trim(tgVehicleInfo(llLoop).sVehicle)
            Exit For
        End If
    Next
End Function
Private Function mGetStationName(iShttCode As Integer) As String
    Dim llLoop As Integer
    mGetStationName = ""
    For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(llLoop).iCode = iShttCode Then
            mGetStationName = Trim(tgStationInfo(llLoop).sCallLetters)
            Exit For
        End If
    Next
End Function

Private Sub mCSIXMLData(slInCommand As String, slInTag As String, slInData As String)
    Dim slCommand As String
    Dim slTag As String
    Dim slData As String
    Dim ilRet As Integer
    ReDim slFields(0 To 2) As String

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
    Dim slOutputType As String
    Dim ilVef As Integer
    Dim ilVehicleSelected As Integer
    Dim ilVff As Integer
    Dim slProp As String
    
    On Error GoTo ErrHand
    imTerminate = False
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    lmMaxWidth = 0
    gClearListScrollBar lbcMsg
    lacResult.Caption = ""
    If (udcCriteria.MOutput(0, slProp) = vbUnchecked) And (udcCriteria.MOutput(1, slProp) = vbUnchecked) Then
        gMsgBox "Please choose an output method--Send to Marketron or Generate to File"
        'ckcOutput(0).SetFocus
        smOutputType = vbNullString
        Exit Sub
    ElseIf (udcCriteria.MOutput(0, slProp) = vbChecked) And (udcCriteria.MOutput(1, slProp) = vbChecked) Then
        smOutputType = "B"
    ElseIf udcCriteria.MOutput(0, slProp) = vbChecked Then
        smOutputType = "T"
    Else
        smOutputType = "F"
    End If
    '5349
    If Not omExport Is Nothing Then
        If omExport.isTest Then
            smOutputType = "F"
        End If
    End If
    '6633 can't find logpath? write out issue, but continue export
    '6807
    'smXmlErrorFile = mGetXmlErrorFile("Marketron", smIniPath, smXmlFileFindError)
    smXmlErrorFile = gGetXmlErrorFile(smXmlFileFindError)
    If Len(smXmlErrorFile) = 0 Then
        myErrors.WriteWarning "Warning, cannot read errors from web server: " & smXmlFileFindError
    End If
    smStartDate = gAdjYear(CSI_Calendar1.Text)
    sNowDate = gObtainPrevMonday(gNow)
    If gDateValue(smStartDate) < gDateValue(gAdjYear(sNowDate)) Then
        Beep
        gMsgBox "Date cannot be previous to this week.", vbCritical
        'CSI_Calendar1.SetFocus
        Exit Sub
    End If
    smEndDate = DateAdd("d", 6, smStartDate)
    ilVehicleSelected = False
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            ilVehicleSelected = True
            Exit For
        End If
    Next ilVef
    If (Not ilVehicleSelected) Then
        Beep
        gMsgBox "Vehicle must be selected.", vbCritical
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    imExporting = True
    mSaveCustomValues
    If Not gPopCopy(smStartDate, "Export Marketron") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    mSetFields
'    '5349 get list of stations
'    If Not omExport Is Nothing Then
'        If Not omExport.GetStations() Then
'            If Len(omExport.ErrorMessage) > 0 Then
'                myErrors.WriteWarning "Issue in mExport-GetStations: " & omExport.ErrorMessage
'                'error? then don't try to test stations in future
'                Set omExport = Nothing
'            End If
'        End If
'    End If
   ' gLogMsg " !! Exporting Spots, Start Date of: " & smStartDate, "MarketronExportLog.Txt", False
    myErrors.WriteFacts "Exporting Spots, Start Date of: " & smStartDate, True
    bmCsfOpenStatus = mOpenCSF()
    '8686
    bgTaskBlocked = False
    sgTaskBlockedName = "Marketron Export"
    ilRet = mExportSpots()
    gCloseRegionSQLRst
    If imTerminate Then
        Call mSetResults("** User Terminated **", RGB(255, 0, 0))
        'gLogMsg "** User Terminated **", "MarketronExportLog.Txt", False
        myErrors.WriteFacts "User Terminated"
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        'cmdCancel.SetFocus
        GoTo Cleanup
    ElseIf (ilRet = False) Then
        Call mSetResults("Export Failed", RGB(255, 0, 0))
        'gLogMsg "** Terminated - mExportSpots returned False **", "MarketronExportLog.Txt", False
        myErrors.WriteWarning "Terminated - mExportSpots returned False"
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        'cmdCancel.SetFocus
        GoTo Cleanup
    End If
    If bmStationNotExist Then
        mSetResults "Some stations could not be sent.  See log.", MESSAGERED
    End If
 '   On Error GoTo ErrHand:
   'gLogMsg "** Completed Export of Marketron **", "MarketronExportLog.Txt", False
    myErrors.WriteFacts "Completed Export of Marketron", True
    cmdCancel.Caption = "&Done"
   ' gLogMsg "", "MarketronExportLog.Txt", False
    If bmAtLeastOneExport Then
        'Call mSetResults("Export Completed Successfully", RGB(0, 155, 0))
        If lbcMsg.ForeColor = RGB(255, 0, 0) Then
            mSetResults "Some exports were not successful.", RGB(0, 155, 0)
            ilRet = gCustomEndStatus(lmEqtCode, 2, "")
        Else
            Call mSetResults("Export Completed Successfully", RGB(0, 155, 0))
            ilRet = gCustomEndStatus(lmEqtCode, 1, "")
        End If
        If smOutputType <> "T" Then
            lacResult.Caption = "Exports placed into: " & sgExportDirectory
        Else
            lacResult.Caption = ""
        End If
    Else
        lacResult.Caption = ""
        ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    End If
    '8686
    If bgTaskBlocked And igExportSource <> 2 Then
         mSetResults "Some spots were blocked during export.", MESSAGERED
         gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
         myErrors.WriteWarning "Some spots were blocked during export.", True
         lacResult.Caption = "Please refer to the Messages folder for file TaskBlocked_" & sgTaskBlockedDate & ".txt."
    End If
Cleanup:
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    cmdExport.Enabled = False
    cmdExportTest.Enabled = False
    Screen.MousePointer = vbDefault
    mCloseCSF
    Exit Sub
ErrHand:
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mExport"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    GoTo Cleanup
End Sub
Private Sub mSetFields()
    smFields(SPOTID) = "SpotID"
    smFields(Advertiser) = "Advertiser"
    smFields(PRODUCTNAME) = "ProductName"
    smFields(PRODUCTCODE) = "ProductCode"
    smFields(startDate) = "StartDate"
    smFields(endDate) = "EndDate"
    smFields(startTime) = "StartTime"
    smFields(endTime) = "EndTime"
    smFields(LENGTH) = "Length"
    smFields(ISCICODE) = "ISCICode"
    smFields(COPYTITLE) = "CopyTitle"
    smFields(COPYCOMMENT) = "CopyComment"
    smFields(COPYENDDATE) = "CopyEndDate"
    
End Sub

Private Sub mClearAlerts(llSDate As Long, llEDate As Long)
    Dim ilVef As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim llStartDate As Long
    Dim ilRet As Integer
    
    slDate = gObtainPrevMonday(Format(llSDate, "m/d/yy"))
    llStartDate = gDateValue(slDate)
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            imVefCode = lbcVehicles.ItemData(ilVef)
            For llDate = llStartDate To llEDate Step 7
                slDate = Format$(llDate, "m/d/yy")
                ilRet = gAlertClear("A", "F", "S", imVefCode, slDate)
                ilRet = gAlertClear("A", "R", "S", imVefCode, slDate)
            Next llDate
        End If
    Next ilVef
    ilRet = gAlertForceCheck()
End Sub

Private Sub mnuTest_Click()
        mnuTest.Checked = True
        omExport.isTest = True
        omExport.SoapUrl = "Test/mx/order/OrderServcies.asmx"
End Sub

Private Sub tmcDelay_Timer()
    '8162
    tmcDelay.Enabled = False
    If IsDate(CSI_Calendar1.Text) Then
        If LenB(CSI_Calendar1.Text) = 0 Then
            gMsgBox "Date must be set"
            'CSI_Calendar1.SetFocus
        ElseIf Weekday(CSI_Calendar1.Text) <> vbMonday Then
            CSI_Calendar1.Text = gObtainPrevMonday(CSI_Calendar1.Text)
        End If
        '8162
        If lbcStation.Visible Then
            lbcVehicles_Click
        End If
    Else
        tmcDelay.Enabled = True
    End If

End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload FrmExportMarketron
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
        For ilLoop = 0 To lbcVehicles.ListCount - 1
            If lbcVehicles.Selected(ilLoop) Then
                ilVefCode(UBound(ilVefCode)) = lbcVehicles.ItemData(ilLoop)
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            End If
        Next ilLoop
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("A", "Marketron", "1", Trim$(CSI_Calendar1.Text), "1", ilVefCode(), ilShttCode())
    End If
End Sub
Private Function mPrepRecordsetDebug() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "ISCI", adChar, 40
            .Append "Advertiser", adChar, 30
            .Append "Date", adDate
            .Append "Time", adChar, 8
            .Append "Station", adChar, 40
            .Append "Vehicle", adChar, 40
            .Append "AstCode", adInteger
        End With
    myRs.Open
    myRs!Station.Properties("optimize") = True
    myRs.Sort = "Station,Vehicle,Date,Time"
    Set mPrepRecordsetDebug = myRs
End Function
Private Function mWriteFacts() As Boolean
    '5050
    Dim myFacts As CLogger
    Dim slPreviousStation As String
    Dim slPreviousVehicle As String
    
    slPreviousStation = ""
    slPreviousVehicle = ""
On Error GoTo errbox
    Set myFacts = New CLogger
    With myFacts
        .LogPath = .CreateLogName(sgExportDirectory & "MarketronFacts")
        .WriteFacts "Marketron export for " & smStartDate, True
        rsFacts.Filter = adFilterNone
        Do While Not rsFacts.EOF
            If slPreviousStation <> rsFacts!Station Then
                slPreviousStation = rsFacts!Station
                .WriteFacts "STATION " & slPreviousStation
                slPreviousVehicle = ""
            End If
            If slPreviousVehicle <> rsFacts!Vehicle Then
                slPreviousVehicle = rsFacts!Vehicle
                .WriteFacts " VECHICLE " & slPreviousVehicle
            End If
            .WriteFacts "       " & Trim$(rsFacts!Advertiser) & " " & Trim$(rsFacts!ISCI) & " " & rsFacts!Date & " " & Trim$(rsFacts!TIME) & " astcode: " & rsFacts!astCode
            rsFacts.MoveNext
        Loop
    End With
    Set myFacts = Nothing
    mWriteFacts = True
    Exit Function
errbox:
    myErrors.WriteError "mWriteFacts-" & Err.Description
    Set myFacts = Nothing
    mWriteFacts = False
End Function
'Private Function mGetXmlErrorFile(slSection As String, slXMLINIInputFile As String, slError As String) As String
''6633
'    Dim slRet As String
'    Dim slFolderPath As String
'    'return "" if doesn't exist in ini, or folder doesn't exit.
'    'return slError for why not returning.
'    slError = ""
'    slFolderPath = ""
'    gLoadFromIni slSection, "LogFile", slXMLINIInputFile, slRet
'    If slRet <> "Not Found" Then
'        slFolderPath = mPathOfFile(slRet)
'        If Len(slFolderPath) > 0 Then
'            '8886
'            'If Dir(slFolderPath, vbDirectory) = vbNullString Then
'            If Not gFolderExist(slFolderPath) Then
'                slError = " The path defined in 'LogFile' in " & slXMLINIInputFile & "- " & slSection & " is not valid: " & slFolderPath
'                slFolderPath = ""
'            Else
'                slFolderPath = gSetPathEndSlash(slFolderPath, False) & XMLERRORFILE
'            End If
'        Else
'            slError = " Could not read value for 'LogFile in " & slXMLINIInputFile & ": " & slSection
'            slFolderPath = ""
'        End If
'    Else
'        slError = " 'LogFile' does not exist in " & slXMLINIInputFile & ": " & slSection
'    End If
'    mGetXmlErrorFile = slFolderPath
'End Function
Private Function mGetXmlError() As String
    Dim myErrorText As TextStream
    Dim slStatus As String
    
    If myErrors.myFile.FILEEXISTS(smXmlErrorFile) Then
On Error GoTo ERRORNOOPEN
        Set myErrorText = myErrors.myFile.OpenTextFile(smXmlErrorFile, ForReading, False)
        slStatus = myErrorText.ReadAll
        myErrorText.Close
        slStatus = mParseXmlError(slStatus)
'dantest must allow
        myErrors.myFile.DeleteFile smXmlErrorFile
    End If
Cleanup:
    mGetXmlError = slStatus
    Set myErrorText = Nothing
    Exit Function
ERRORNOOPEN:
    slStatus = "Issue reading xml error file"
    myErrors.WriteError "could not read xml error file " & Err.Description, True
    GoTo Cleanup
End Function
Private Function mParseXmlError(slMessage As String) As String
    '7509 change to long
    Dim ilPos As Long
    Dim slRet As String
    Dim ilEnd As Long
    
    slRet = slMessage
    ilPos = InStr(slMessage, "<MESSAGES")
    If ilPos > 0 Then
        ilPos = InStr(ilPos, slMessage, ">")
        If ilPos > 0 Then
            ilEnd = InStr(ilPos, slMessage, "</MESSAGE>")
            If ilEnd > ilPos Then
                slRet = Mid(slMessage, ilPos + 1, ilEnd - ilPos - 1)
            End If
        End If
    '8173, returning MSGS on timeout Let's just get msg
    Else
        ilPos = InStr(slMessage, "<MSG ")
        If ilPos > 0 Then
            ilEnd = InStr(ilPos, slMessage, "</MSG>")
            If ilEnd > ilPos Then
                slRet = Mid(slMessage, ilPos + 4, ilEnd - ilPos - 4)
            End If
        End If
    End If
    mParseXmlError = slRet
End Function
Private Function mPathOfFile(slFile As String) As String
    Dim ilPos As Integer
    ilPos = InStrRev(slFile, "\")
    If ilPos > 0 Then
        'mPathOfFile = Left$(slFile, ilPos)
        mPathOfFile = Mid(slFile, 1, ilPos)
    Else
        mPathOfFile = ""
    End If
End Function
