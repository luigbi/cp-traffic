VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form FrmExportWegener 
   Caption         =   "Export Wegener"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportWegener.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9615
   Begin VB.CheckBox chkRunCheckUtility 
      Caption         =   "Run Check Utility"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox edcDays 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3285
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "# of Days"
      Top             =   165
      Width           =   930
   End
   Begin VB.TextBox edcStartDate 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Export Start Date"
      Top             =   165
      Width           =   1530
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9225
      Top             =   285
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1770
      TabIndex        =   0
      Top             =   135
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "02/12/2024"
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   3
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Result"
      Top             =   2445
      Width           =   4740
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   2430
      Width           =   3870
   End
   Begin VB.ListBox lbcGroupDef 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "AffExportWegener.frx":08CA
      Left            =   8625
      List            =   "AffExportWegener.frx":08D1
      Sorted          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox edcMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   420
      Left            =   1770
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   510
      Visible         =   0   'False
      Width           =   4770
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
      Left            =   9300
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   1005
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4275
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4995
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      ItemData        =   "AffExportWegener.frx":08E2
      Left            =   4650
      List            =   "AffExportWegener.frx":08E4
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2670
      Width           =   4755
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportWegener.frx":08E6
      Left            =   120
      List            =   "AffExportWegener.frx":08E8
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2685
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8670
      Top             =   990
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6060
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5910
      TabIndex        =   6
      Top             =   5430
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   7
      Top             =   5430
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label lacResult 
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   5340
      Width           =   5580
   End
End
Attribute VB_Name = "FrmExportWegener"
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
Option Explicit
Option Compare Text

Private hmWegExpErrLog As Integer
Private hmWegRegSpotErrLog As Integer
Private hmFrom As Integer
'Private smFields(1 To 31) As String
Private smFields(0 To 30) As String
Private smDate As String     'Export Date
Private imImportStationInfo As Boolean
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private smVehicleGroupPrefix As String
Private imGameNo() As Integer
Private lmGsfCode() As Long
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private smExportPath As String
Private imTerminate As Integer
Private imFirstTime As Integer
Private lmMaxWidth As Long
Private smExportDirectory As String
Private hmCSV As Integer
Private smMP2FilePath As String
Private smVehicleGroupName As String
Private smCustomGroupName As String
Private imCustomGroupNo As Integer
Private tmRegionBreakSpots() As REGIONBREAKSPOTS
Private tmTempRegionBreakSpots() As REGIONBREAKSPOTS
Private tmSplitNetRegion() As SPLITNETREGION
Private tmRegionDefinition() As REGIONDEFINITION
Private tmSplitCategoryInfo() As SPLITCATEGORYINFO
'3/3/18: Retain station list
Private tmStationSplitCategoryInfo() As SPLITCATEGORYINFO
Private imRowWithStation() As Integer  'Rows with condensed station list
Private imAllowedShttCode() As Integer

Private tmMergeRegionDefinition() As REGIONDEFINITION
Private tmMergeSplitCategoryInfo() As SPLITCATEGORYINFO
Private tmCustomGroupNames() As CUSTOMGROUPNAMES
Private tmWegenerImport() As WEGENERIMPORT
Private tmWegenerVehInfo() As WEGENERVEHINFO
Private smGroupName() As String
Private smWegenerGroupChar As String
Private imPortDefined(0 To 3) As Integer
Private tmWegenerFormatSort() As WEGENERFORMATSORT
Private tmWegenerTimeZoneSort() As WEGENERTIMEZONESORT
Private tmWegenerPostalSort() As WEGENERPOSTALSORT
Private tmWegenerMarketSort() As WEGENERMARKETSORT
Private tmWegenerMSAMarketSort() As WEGENERMARKETSORT
Private tmWegenerIndex() As WEGENERINDEX
Private cprst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private raf_rst As ADODB.Recordset
Private vff_rst As ADODB.Recordset
Private rsf_rst As ADODB.Recordset
Private err_rst As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset
Private tmModelLST As LST
Private tmLst As LST
'Dan M 11/01/10 search for xml.ini once and store in new variable
Private smIniPathFileName As String
Private lmEqtCode As Long
'12/1/17: Test max region combinations in break
Private bmRegionMaxExceeded As Boolean
Private lmPrevSdfCode As Long
Private Type WEGENERINFOINDEX
    iImport As Integer
    lSerialNo As Long
End Type
Dim tmWegenerInfoIndex() As WEGENERINFOINDEX

Private Sub mFillVehicle()
    Dim iLoop As Integer
    Dim llVpf As Long
    
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llVpf = gBinarySearchVpf(CLng(tgVehicleInfo(iLoop).iCode))
        If llVpf <> -1 Then
            If tgVpfOptions(llVpf).sWegenerExport = "Y" Then
                lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
            End If
        End If
    Next iLoop
End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
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

Private Sub cmdExport_Click()
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slExportType As String
    Dim slXMLFileName As String
    Dim slOutputType As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilPos As Integer
    Dim ilVff As Integer
    
    On Error GoTo ErrHand
    
    If imExporting = True Then
        Exit Sub
    End If
    imExporting = True
    '12/1/17: Test max region combinations in break
    bmRegionMaxExceeded = False
    lmPrevSdfCode = -1
    lbcMsg.Clear
    slNowDate = Format$(gNow(), "m/d/yy")
    If udcCriteria.WGenerate(1) = vbChecked Then
        If lbcVehicles.ListIndex < 0 Then
            igExportReturn = 2
            imExporting = False
            Exit Sub
        End If
        If edcDate.Text = "" Then
            imExporting = False
            gMsgBox "Date must be specified.", vbOKOnly
            'edcDate.SetFocus
            Exit Sub
        End If
        If gIsDate(edcDate.Text) = False Then
            imExporting = False
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            'edcDate.SetFocus
            Exit Sub
        Else
            smDate = Format(edcDate.Text, sgShowDateForm)
        End If
        imNumberDays = Val(txtNumberDays.Text)
        If imNumberDays <= 0 Then
            imExporting = False
            gMsgBox "Number of days must be specified.", vbOKOnly
            'txtNumberDays.SetFocus
            Exit Sub
        End If
        Select Case Weekday(gAdjYear(smDate))
            Case vbMonday
                If imNumberDays > 7 Then
                    gMsgBox "Number of days can not exceed 7.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbTuesday
                If imNumberDays > 6 Then
                    gMsgBox "Number of days can not exceed 6.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbWednesday
                If imNumberDays > 5 Then
                    gMsgBox "Number of days can not exceed 5.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbThursday
                If imNumberDays > 4 Then
                    gMsgBox "Number of days can not exceed 4.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbFriday
                If imNumberDays > 3 Then
                    gMsgBox "Number of days can not exceed 3.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbSaturday
                If imNumberDays > 2 Then
                    gMsgBox "Number of days can not exceed 2.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
            Case vbSunday
                If imNumberDays > 1 Then
                    gMsgBox "Number of days can not exceed 1.", vbOKOnly
                    'txtNumberDays.SetFocus
                    Exit Sub
                End If
        End Select
        If gDateValue(gAdjYear(smDate)) < gDateValue(gAdjYear(slNowDate)) Then
            imExporting = False
            Beep
            gMsgBox "Date must be on or after today's date " & slNowDate, vbCritical
            'edcDate.SetFocus
            Exit Sub
        End If
        If udcCriteria.WRunLetter = "" Then
            imExporting = False
            gMsgBox "Run letter must be specified.", vbOKOnly
            'txtRunLetter.SetFocus
            Exit Sub
        End If
    End If
    'If (udcCriteria.edcWStationInfo = "") And (udcCriteria.WGenerate(0) = vbChecked) Then
    If (udcCriteria.edcWStationInfo = "") And (imImportStationInfo) Then
        imExporting = False
        gMsgBox "Import file path must be specified.", vbOKOnly
        'txtStationInfo.SetFocus
        Exit Sub
    End If
    'If (rbcSpots(0).Value = False) And (rbcSpots(1).Value = False) Then
    '    Beep
    '    gMsgBox "Please Specify Export Spots Type.", vbCritical
    '    Exit Sub
    'End If
    gGetPoolAdf
    smExportDirectory = udcCriteria.WExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    '7342 Dan
    lbcMsg.Clear
    '8886
    'If Dir(smExportDirectory, vbDirectory) = vbNullString Then
    If Not gFolderExist(smExportDirectory) Then
        mAddMsgToList "Chosen directory does not exist.  Export file will be written to generic export folder."
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)

    Screen.MousePointer = vbHourglass
    mSaveCustomValues
    If Not gPopCopy(smDate, "Export Wegener") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    
    ' JD 01-25-24 Added support for the new Wegener Check Utility
    Call RunCheckUtility
    If Not imExporting Then ' Gets set to false if the user
        Exit Sub
    End If
    
    ilRet = gPopVff()
    smVehicleGroupPrefix = ""
    For ilVff = LBound(tgVffInfo) To UBound(tgVffInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        ilPos = InStr(1, tgVffInfo(ilVff).sGroupName, "_", vbTextCompare)
        If (ilPos >= 2) And (ilPos <= 4) Then
            smVehicleGroupPrefix = UCase(Left(tgVffInfo(ilVff).sGroupName, ilPos))
            Exit For
        End If
    Next ilVff
    smWegenerGroupChar = ""
    SQLQuery = "SELECT spfWegenerGroupChar"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        smWegenerGroupChar = Trim$(rst!spfWegenerGroupChar)
    End If
    If smWegenerGroupChar = "" Then
        smWegenerGroupChar = "W"
    End If
    imExporting = True
    'If udcCriteria.WGenerate(0) = vbChecked Then
    If imImportStationInfo Then
        edcMsg.Text = "Reading Station Info...."
        edcMsg.Visible = True
        If igExportSource = 2 Then DoEvents
        ilRet = mReadStationReceiverRecords()
        edcMsg.Visible = False
        If igExportSource = 2 Then DoEvents
        If ilRet <> 0 Then
            imExporting = False
            Screen.MousePointer = vbDefault
            If ilRet = 1 Then
                igExportReturn = 2
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                Exit Sub
            ElseIf ilRet = 2 Then
                igExportReturn = 2
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                Exit Sub
            Else
                If igExportSource = 2 Then
                    ilRet = vbYes
                Else
                    ilRet = gMsgBox("Some Stations Not Defined within the Affiliate system, Continue anyway", vbYesNo + vbQuestion, "Information")
                    If ilRet = vbNo Then
                        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                        Exit Sub
                    End If
                End If
            End If
            imExporting = True
            Screen.MousePointer = vbHourglass
        End If
        On Error GoTo 0
        edcMsg.Text = "Updating Station Info...."
        edcMsg.Visible = True
        If igExportSource = 2 Then DoEvents
        ilRet = mUpdateShttUsedForWegener()
        edcMsg.Visible = False
        If igExportSource = 2 Then DoEvents
        If (Not ilRet) And (igExportSource <> 2) Then
            Screen.MousePointer = vbDefault
            ilRet = gMsgBox("Unable to set Used for Wegener with all Stations, Continue anyway", vbYesNo + vbQuestion, "Information")
            If ilRet = vbNo Then
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                imExporting = False
                Exit Sub
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    'dan moved 7342
    'lbcMsg.Clear
    lacResult.Caption = ""
    If udcCriteria.WGenerate(1) = vbChecked Then
        'If rbcSpots(0).Value = True Then
            slExportType = "!! Exporting All Spots, "
        'Else
        '    slExportType = "!! Exporting Regional Spots, "
        'End If
        'ilRet = 0
        slToFile = sgMsgDirectory & "WegenerExportLog.Txt"
        On Error GoTo mFileErr
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            slDateTime = gFileDateTime(slToFile)
            slFileDate = Format$(slDateTime, "m/d/yy")
            If gDateValue(gAdjYear(slFileDate)) = gDateValue(gAdjYear(slNowDate)) Then  'Append
                gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "WegenerExportLog.Txt", False
            Else
                gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "WegenerExportLog.Txt", True
            End If
        Else
            gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "WegenerExportLog.Txt", False
        End If
        On Error GoTo ErrHand
        ilRet = mExportSpots()
        gCloseRegionSQLRst
        edcMsg.Visible = False
        If igExportSource = 2 Then DoEvents
        If (ilRet = False) Then
            gLogMsg "** Terminated - mExportSpots returned False **", "WegenerExportLog.Txt", False
            ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
            imExporting = False
            Screen.MousePointer = vbDefault
            'cmdCancel.SetFocus
            Exit Sub
        End If
        If imTerminate Then
            gLogMsg "** User Terminated **", "WegenerExportLog.Txt", False
            ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
            imExporting = False
            Screen.MousePointer = vbDefault
            'cmdCancel.SetFocus
            Exit Sub
        End If
        On Error GoTo ErrHand:
        'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
        gLogMsg "** Completed Export of Wegener **", "WegenerExportLog.Txt", False
        'Close #hmMsg
        If slOutputType <> "T" Then
            lacResult.Caption = "Exports placed into: " & smExportPath
        Else
            lacResult.Caption = ""
        End If
    End If
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    imExporting = False
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsg "", "WegenerExportLog.Txt", False
    '12/1/17: Test max region combinations in break
    If bmRegionMaxExceeded And igExportSource <> 2 Then
        gMsgBox "Region Break Definition Exceeded: See WegenerExportLog.txt. All definitions have not been exported", vbCritical
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerExportLog.txt", "Export Wegener-mcmdExport_Click"
    Exit Sub
mFileErr:
    ilRet = Err.Number
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload FrmExportWegener
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        udcCriteria.Left = edcStartDate.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = edcStartDate.Top + edcStartDate.Height / 2
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
                End If
            Next ilLoop
            chkAll.Value = vbChecked
        End If
        If igExportSource = 2 Then
            slNowStart = gNow()
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "WegenerResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "Wegener Result List, Started: " & slNowStart
           ' pass global so glogMsg will write messages to sgExportResultName
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "WegenerResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "Wegener Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "Wegener Result List, Completed: " & slNowEnd
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
    Me.Height = Screen.Height / 1.2 '1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    edcMsg.Move (Me.Width - edcMsg.Width) / 2, edcStartDate.Top + edcStartDate.Height + 120 '(Me.Height - Msg.Height) / 2
    gSetFonts Me
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim slSvIniPathFileName As String
    
    Screen.MousePointer = vbHourglass
    FrmExportWegener.Caption = "Export Wegener - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcDate.Text = smDate
    txtNumberDays.Text = 1
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    imImportStationInfo = True  'Replace udcCriteria.WGenerate(0) as it is visible = False
    'If Len(sgImportDirectory) > 0 Then
    '    txtStationInfo.Text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
    'Else
    '    txtStationInfo.Text = ""
    'End If
    mFillVehicle
    chkAll.Value = vbChecked
    slSvIniPathFileName = sgIniPathFileName
    'dan m 11/01/10 xml.ini gotten from general procedure; checking different folders
    'sgIniPathFileName = sgStartupDirectory & "\XML.Ini"
    smIniPathFileName = gXmlIniPath()
    sgIniPathFileName = smIniPathFileName
    
    'ilRet = gPopAvailNames()
    
    If Not gLoadOption("Wegener", "MP2FilePath", smMP2FilePath) Then
        smMP2FilePath = ""
    End If
    If Not gLoadOption("Wegener", "Export", smExportPath) Then
        smExportPath = sgExportDirectory
    End If
    smExportPath = gSetPathEndSlash(smExportPath, True)
    sgIniPathFileName = slSvIniPathFileName
    ilRet = gPopAvailNames()
    '7342
    lacResult.Caption = "Exports placed into: " & mLoseLastLetter(smExportPath)
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    Erase tmRegionBreakSpots
    Erase tmTempRegionBreakSpots
    Erase tmSplitNetRegion
    Erase tmRegionDefinition
    Erase tmSplitCategoryInfo
        '3/3/18
    Erase tmStationSplitCategoryInfo
    Erase imRowWithStation
    Erase imAllowedShttCode

    Erase tmMergeRegionDefinition
    Erase tmMergeSplitCategoryInfo
    Erase tmCustomGroupNames
    Erase tmWegenerImport
    Erase tmWegenerVehInfo
    Erase smGroupName
    Erase tmWegenerFormatSort
    Erase tmWegenerTimeZoneSort
    Erase tmWegenerMarketSort
    Erase tmWegenerMSAMarketSort
    Erase tmWegenerPostalSort
    Erase tmWegenerIndex
    Erase imGameNo
    Erase lmGsfCode
    cprst.Close
    lst_rst.Close
    raf_rst.Close
    vff_rst.Close
    rsf_rst.Close
    err_rst.Close
    rst_Gsf.Close
    Set FrmExportWegener = Nothing
End Sub



Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
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
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(iLoop) Then
            imVefCode = lbcVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
End Sub

Private Sub edcDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub


Private Function mExportSpots() As Integer
    'Export all spots with its general copy for the specified vehicle and days
    'Each vehicle will create a separate export file.
    'All days will be within the same export file
    'The spots are obtained from LST instead of AST as Wegener will create the spots for each station
    'If any spot within a break has region copy, then all spots within that break must be export
    'along with each region definition for the spot.
    'This part of the export must be after all the general copy is exported for the vehicle and days.
    '
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llStartLstDate As Long
    Dim llEndLstDate As Long
    Dim slWeekNo As String
    Dim slYearNo As String
    Dim slVehName As String
    Dim slVehGroupName As String
    Dim slVehExportID As String
    Dim llODate As Long
    'Dim slDate As String
    'Dim llDate As Long
    Dim slXMLFileName As String
    Dim slCSVFileName As String
    Dim slNowDT As String
    Dim slTimeID As String
    Dim llEventID As Long
    Dim slEventID As String
    Dim llBreakNo As Long
    Dim llLogTime As Long
    Dim llLstLogDate As Long
    Dim llLstLogTime As Long
    Dim ilPos As Integer
    Dim slOutputType As String
    Dim ilRegionWithInBreak As Integer
    Dim ilLoop As Integer
    Dim slRunLetter As String
    Dim ilAnyExports As Integer
    Dim ilPositionNo As Integer
    Dim slCSVRecord As String
    Dim llAdf As Long
    Dim slAdvtName As String
    Dim slGrpName As String
    Dim slRCartNo As String
    Dim slRProduct As String
    Dim slRISCI As String
    Dim slRCreativeTitle As String
    Dim llRCrfCsfCode As Long
    Dim llRCpfCode As Long
    Dim ilCifAdfCode As Integer
    Dim slISCI As String
    Dim slStdStartDate As String
    Dim llVefIndex As Long
    Dim ilGsf As Integer
    Dim slZone As String
    Dim ilVefZone As Integer
    Dim ilZone As Integer
    Dim llLocalAdj As Long
    Dim blSplitNetworkSpot As Boolean
    Dim blCreateFill As Boolean
    Dim slSplitNetISCI As String
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
    '7458
    Dim myEnt As CENThelper
    'Handle multi-vehicles with the same vehicle group name
    Dim ilFile As Integer
    Dim blFound As Boolean
    ReDim slFileNamesExported(0 To 0) As String
    
    On Error GoTo ErrHand
    'D.S. 11/17/17 - TTP #8687
    bgTaskBlocked = False
    mExportSpots = True
    ilAnyExports = False
    slRunLetter = Trim$(udcCriteria.WRunLetter)
    slSDate = smDate
    slEDate = DateAdd("d", imNumberDays - 1, smDate)
    llSDate = gDateValue(gAdjYear(slSDate))
    llEDate = gDateValue(gAdjYear(slEDate))
    'slWeekNo = Format(slSDate, "ww")
    'If Len(slWeekNo) = 1 Then
    '    slWeekNo = "0" & slWeekNo
    'End If
    'slYearNo = Format(slSDate, "yy")
    'If Len(slYearNo) = 1 Then
    '    slYearNo = "0" & slYearNo
    'End If
    slStdStartDate = gObtainYearStartDate(slSDate)
    slWeekNo = Trim$(Str$(DateDiff("ww", slStdStartDate, slSDate, vbMonday) + 1))
    If Len(slWeekNo) = 1 Then
        slWeekNo = "0" & slWeekNo
    End If
    slYearNo = Format(gObtainEndStd(slSDate), "yy")
    If Len(slYearNo) = 1 Then
        slYearNo = "0" & slYearNo
    End If
    llEventID = 0
    imCustomGroupNo = 0
    'smCustomGroupName = "W" & (Asc(UCase$(slRunLetter)) - Asc("A") + 1) & slWeekNo
    smCustomGroupName = smWegenerGroupChar & (Asc(UCase$(slRunLetter)) - Asc("A") + 1) & slWeekNo
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .ThirdParty = Vendors.Wegener_Compel
        .TypeEnt = Exportunposted3rdparty
        .User = igUstCode
        .Station = 0
        .Agreement = 0
        .ErrorLog = "WegenerExportLog.txt"
        If Len(.ErrorMessage) > 0 Then
            'gLogMsgWODT "W", hmCSV, myEnt.ErrorMessage
            gLogMsg myEnt.ErrorMessage, "WegenerExportLog.Txt", False
        End If
    End With
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mExportSpots = False
                Exit For
            End If
            slVehName = Trim$(lbcVehicles.List(ilVef))
            imVefCode = lbcVehicles.ItemData(ilVef)
            slZone = "E"
            llLocalAdj = 0
            llStartLstDate = llSDate
            llEndLstDate = llEDate
            ilVefZone = gBinarySearchVef(CLng(imVefCode))
            If ilVefZone <> -1 Then
                For ilZone = LBound(tgVehicleInfo(ilVefZone).sZone) To UBound(tgVehicleInfo(ilVefZone).sZone) Step 1
                    If igExportSource = 2 Then DoEvents
                    If UCase$(Left$(Trim$(tgVehicleInfo(ilVefZone).sZone(ilZone)), 1)) = "E" Then
                        slZone = UCase$(Left$(tgVehicleInfo(ilVefZone).sZone(tgVehicleInfo(ilVefZone).iBaseZone(ilZone)), 1))
                        If (tgVehicleInfo(ilVefZone).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVefZone).sFed(ilZone)) <> "") Then
                            llLocalAdj = CLng(3600) * tgVehicleInfo(ilVefZone).iLocalAdj(ilZone)
                            If llLocalAdj > 0 Then
                                llStartLstDate = llStartLstDate - 1
                            Else
                                llEndLstDate = llEndLstDate + 1
                            End If
                        End If
                        Exit For
                    End If
                Next ilZone
            End If
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
            SQLQuery = SQLQuery + " WHERE (vffVefCode = " & imVefCode & ")"
            Set vff_rst = gSQLSelectCall(SQLQuery)
            If Not vff_rst.EOF Then
                edcMsg.Text = "Generating General Schedule for " & slVehName & "..."
                edcMsg.Visible = True
                If igExportSource = 2 Then DoEvents
                ilAnyExports = True
                ReDim imGameNo(0 To 1) As Integer
                ReDim lmGsfCode(0 To 1) As Long
                imGameNo(0) = 0
                lmGsfCode(0) = 0
                '7458
                With myEnt
                    .Vehicle = imVefCode
                    .ProcessStart
                End With
                llVefIndex = gBinarySearchVef(CLng(imVefCode))
                If llVefIndex <> -1 Then
                    If tgVehicleInfo(llVefIndex).sVehType = "G" Then
                        ReDim imGameNo(0 To 0) As Integer
                        ReDim lmGsfCode(0 To 0) As Long
                        SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & imVefCode & " AND gsfAirDate >= '" & Format$(slSDate, sgSQLDateForm) & "'" & " AND gsfAirDate <= '" & Format$(slEDate, sgSQLDateForm) & "'" & ")"
                        Set rst_Gsf = gSQLSelectCall(SQLQuery)
                        Do While Not rst_Gsf.EOF
                            If igExportSource = 2 Then DoEvents
                            imGameNo(UBound(imGameNo)) = rst_Gsf!gsfGameNo
                            lmGsfCode(UBound(lmGsfCode)) = rst_Gsf!gsfCode
                            ReDim Preserve imGameNo(0 To UBound(imGameNo) + 1) As Integer
                            ReDim Preserve lmGsfCode(0 To UBound(lmGsfCode) + 1) As Long
                            rst_Gsf.MoveNext
                        Loop
                        rst_Gsf.Close
                    End If
                End If
                For ilGsf = 0 To UBound(lmGsfCode) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    mSetPortsRequired
                    slVehGroupName = Trim$(vff_rst!VffGroupName)
                    slVehExportID = Trim$(vff_rst!VffWegenerExportID)
                    smVehicleGroupName = slVehGroupName
                    If lmGsfCode(ilGsf) = 0 Then
                        slXMLFileName = "PLSched_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & slRunLetter & ".XML"
                        slCSVFileName = "PLSched_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & slRunLetter & ".CSV"
                    Else
                        slXMLFileName = "PLSched_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & Trim$(Str$(imGameNo(ilGsf))) & "_" & slRunLetter & ".XML"
                        slCSVFileName = "PLSched_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & Trim$(Str$(imGameNo(ilGsf))) & "_" & slRunLetter & ".CSV"
                    End If
                    slOutputType = "F"
                    '7458
                    myEnt.fileName = slXMLFileName
                    '6808
                    'Handle multi-vehicles with the same Group name (append to file)
                    'If Not gDeleteFile(smExportPath & slXMLFileName) Then
                    '    mAddMsgToList "Could not delete file " & slXMLFileName & " before writing.  Appended."
                    'End If
                    blFound = False
                    For ilFile = 0 To UBound(slFileNamesExported) - 1 Step 1
                        If StrComp(slXMLFileName, Trim$(slFileNamesExported(ilFile)), vbBinaryCompare) = 0 Then
                            blFound = True
                            Exit For
                        End If
                    Next ilFile
                    If Not blFound Then
                        slFileNamesExported(UBound(slFileNamesExported)) = slXMLFileName
                        ReDim Preserve slFileNamesExported(0 To UBound(slFileNamesExported) + 1) As String
                        If Not gDeleteFile(smExportPath & slXMLFileName) Then
                            mAddMsgToList "Could not delete file " & slXMLFileName & " before writing.  Appended."
                        End If
                    End If
                    'User wamts CrLf.  The ulity is adding a Cr so only add the Lf in this code
                    ' Dan M 11/01/10 use smIniPathFileName that is created at formload
                    'ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "Wegener", slOutputType, smExportPath & slXMLFileName, sgCRLF)
                    '6807
                    'ilRet = csiXMLStart(smIniPathFileName, "Wegener", slOutputType, smExportPath & slXMLFileName, sgCRLF)
                    ilRet = csiXMLStart(smIniPathFileName, "Wegener", slOutputType, smExportPath & slXMLFileName, sgCRLF, "")
                    ilRet = csiXMLSetMethod("", "", "", "Playlist_Schedule")
                    If igExportSource = 2 Then DoEvents
                    csiXMLData "OT", "Interface_Header", ""
                    csiXMLData "CD", "SOURCE", "Traffic system"
                    csiXMLData "CD", "TARGET", "Compel"
                    slNowDT = Now
                    csiXMLData "CD", "CREATION_DATETIME", Format$(slNowDT, "yyyy-mm-ddThh:mm:ss")
                    csiXMLData "CD", "DESCRIPTION", "Playlist Schedule"
                    csiXMLData "CD", "INTERFACE_ID", "PLSched"
                    slTimeID = Timer
                    ilPos = InStr(1, slTimeID, ".", vbTextCompare)
                    If ilPos > 0 Then
                        slTimeID = Left(slTimeID, ilPos - 1) & Mid(slTimeID, ilPos + 1)
                    End If
                    slTimeID = Left$(slTimeID, 3)
                    csiXMLData "CD", "MESSAGE_ID", slWeekNo & slYearNo & slTimeID
                    csiXMLData "CT", "Interface_Header", ""
                    If udcCriteria.WGenCSV = vbChecked Then
                        If igExportSource = 2 Then DoEvents
                        'Handle multi-vehicles with the same Group name (append to file)
                        'gLogMsgWODT "ON", hmCSV, smExportDirectory & slCSVFileName
                        If Not blFound Then
                            gLogMsgWODT "ON", hmCSV, smExportDirectory & slCSVFileName
                        Else
                            gLogMsgWODT "OA", hmCSV, smExportDirectory & slCSVFileName
                        End If
                        slCSVRecord = "Vehicle " & gAddQuotes(Trim$(lbcVehicles.List(ilVef)))
                        gLogMsgWODT "W", hmCSV, slCSVRecord
                        slCSVRecord = "Start Date " & slSDate & " End Date " & slEDate
                        gLogMsgWODT "W", hmCSV, slCSVRecord
                        slCSVRecord = "Creation Date and Time " & Format$(slNowDT, "yyyy-mm-ddThh:mm:ss") & " Run Letter " & slRunLetter
                        gLogMsgWODT "W", hmCSV, slCSVRecord
                        slCSVRecord = "Message ID " & slWeekNo & slYearNo & slTimeID
                        gLogMsgWODT "W", hmCSV, slCSVRecord
                        slCSVRecord = "LogDate,LogTime,Break#,Position#,Contract#,Advertiser,Line#,Length,ISCI,Region Names..."
                        gLogMsgWODT "W", hmCSV, slCSVRecord
                    End If
                    slCSVRecord = ""
                    ReDim tmRegionBreakSpots(0 To 0) As REGIONBREAKSPOTS
                    ReDim tmTempRegionBreakSpots(0 To 0) As REGIONBREAKSPOTS
                    ReDim smGroupName(0 To 0) As String
                    ReDim tmSplitNetRegion(0 To 0) As SPLITNETREGION
                    lbcGroupDef.Clear
                    ilRegionWithInBreak = False
                    llODate = -1
                    llLogTime = -1
                    SQLQuery = "SELECT * FROM lst "
                    SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & imVefCode
                    If lmGsfCode(ilGsf) > 0 Then
                        SQLQuery = SQLQuery + " AND lstGsfCode = " & lmGsfCode(ilGsf)
                    End If
                    SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
                    SQLQuery = SQLQuery + " AND lstType = 0"    '0=Spot, 1=Avail
                    '3/9/16: Fix the filter
                    'SQLQuery = SQLQuery + " AND lstStatus < 20" 'Bypass MG/Bonus
                    SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
                    SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(llStartLstDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llEndLstDate, sgSQLDateForm) & "')" & ")"
                    SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                    Set lst_rst = gSQLSelectCall(SQLQuery)
                    If Not lst_rst.EOF Then
                        Do While Not lst_rst.EOF
                            If igExportSource = 2 Then DoEvents
                            blSplitNetworkSpot = False
                            blSpotOk = True
                            ilAnf = gBinarySearchAnf(lst_rst!lstAnfCode)
                            If ilAnf <> -1 Then
                                If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                                    blSpotOk = False
                                End If
                            End If
                            If igExportSource = 2 Then DoEvents
                            If (blSpotOk) And ((Left$(UCase(Trim$(lst_rst!lstZone)), 1) = slZone) Or (Trim$(lst_rst!lstZone) = "")) Then
                                'slDate = Format$(lst_rst!lstLogDate, sgShowDateForm)
                                'llDate = gdateValue(gAdjYear(slDate))
                                '7458
                                If Not myEnt.Add(lst_rst!lstLogDate, lst_rst!lstGsfCode, , , True) Then
                                    'gLogMsgWODT "W", hmCSV, myEnt.ErrorMessage
                                    gLogMsg myEnt.ErrorMessage, "WegenerExportLog.Txt", False
                                End If
                                llLstLogDate = gDateValue(gAdjYear(Format$(lst_rst!lstLogDate, sgShowDateForm)))
                                llLstLogTime = gTimeToLong(Format$(lst_rst!lstLogTime, sgShowTimeWSecForm), False)
                                llLstLogTime = llLstLogTime + llLocalAdj
                                If llLstLogTime < 0 Then
                                    llLstLogTime = llLstLogTime + 86400
                                    llLstLogDate = llLstLogDate - 1
                                ElseIf llLstLogTime > 86400 Then
                                    llLstLogTime = llLstLogTime - 86400
                                    llLstLogDate = llLstLogDate + 1
                                End If
                                If (llLstLogDate >= llSDate) And (llLstLogDate <= llEDate) Then
                                    If llODate <> llLstLogDate Then
                                        If igExportSource = 2 Then DoEvents
                                        If llODate <> -1 Then
                                            If llLogTime <> -1 Then
                                                csiXMLData "CT", "Playlist_Info", ""
                                            End If
                                            csiXMLData "CT", "Playlist_Set", ""
                                        End If
                                        llODate = llLstLogDate  'llDate
                                        csiXMLData "OT", "Playlist_Set", ""
                                        csiXMLData "CD", "Address", slVehGroupName
                                        csiXMLData "CD", "Vehicle_Name", gXMLNameFilter(slVehName)
                                        csiXMLData "CD", "Vehicle_Code", slVehExportID
                                        'csiXMLData "CD", "Day_of_Week", UCase$(Left$(Format$(llDate, "ddd"), 2))
                                        csiXMLData "CD", "Day_of_Week", UCase$(Left$(Format$(llLstLogDate, "ddd"), 2))
                                        llBreakNo = 0
                                        ilPositionNo = 0
                                        llLogTime = -1
                                    End If
                                    If llLogTime <> llLstLogTime Then
                                        If (UBound(tmTempRegionBreakSpots) > 0) And (ilRegionWithInBreak) Then
                                            For ilLoop = 0 To UBound(tmTempRegionBreakSpots) - 1 Step 1
                                                If igExportSource = 2 Then DoEvents
                                                tmRegionBreakSpots(UBound(tmRegionBreakSpots)) = tmTempRegionBreakSpots(ilLoop)
                                                ReDim Preserve tmRegionBreakSpots(0 To UBound(tmRegionBreakSpots) + 1) As REGIONBREAKSPOTS
                                            Next ilLoop
                                        End If
                                        If igExportSource = 2 Then DoEvents
                                        If llLogTime <> -1 Then
                                            csiXMLData "CT", "Playlist_Info", ""
                                        End If
                                        llLogTime = llLstLogTime
                                        llEventID = llEventID + 1
                                        slEventID = Trim$(Str$(llEventID))
                                        Do While Len(slEventID) < 5
                                            slEventID = "0" & slEventID
                                        Loop
                                        csiXMLData "OT", "Playlist_Info", "Id=" & """" & slTimeID & slEventID & """"
                                        llBreakNo = llBreakNo + 1
                                        ilPositionNo = 0
                                        csiXMLData "CD", "Break_Number", Trim$(Str$(llBreakNo))
                                        ReDim tmTempRegionBreakSpots(0 To 0) As REGIONBREAKSPOTS
                                        ilRegionWithInBreak = False
                                    End If
                                    If lst_rst!lstsplitnetwork = "P" Then
                                        'Loop until either Fill found or next primary or regular spot
                                        If igExportSource = 2 Then DoEvents
                                        blSplitNetworkSpot = True
                                        gCreateUDTforLST lst_rst, tmModelLST
                                        SQLQuery = "Select rafName from RAF_Region_Area"
                                        SQLQuery = SQLQuery & " Where (rafCode = " & lst_rst!lstRafCode & ")"
                                        Set raf_rst = gSQLSelectCall(SQLQuery)
                                        If Not raf_rst.EOF Then
                                            slSplitNetISCI = "," & gAddQuotes(Trim$(tmModelLST.sISCI) & " (" & Trim$(raf_rst!rafName) & ")")
                                        Else
                                            slSplitNetISCI = "," & gAddQuotes(Trim$(tmModelLST.sISCI) & " (" & "Primary" & ")")
                                        End If
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).sSource = "N"
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).iFirstSplitNetRegion = UBound(tmSplitNetRegion)
                                        tmSplitNetRegion(UBound(tmSplitNetRegion)).lLstCode = tmModelLST.lCode
                                        tmSplitNetRegion(UBound(tmSplitNetRegion)).iNext = -1
                                        ReDim Preserve tmSplitNetRegion(0 To UBound(tmSplitNetRegion) + 1) As SPLITNETREGION
                                        'Build array of split network regions
                                        blCreateFill = True
                                        lst_rst.MoveNext
                                        Do While Not lst_rst.EOF
                                            If igExportSource = 2 Then DoEvents
                                            If lst_rst!lstsplitnetwork = "F" Then
                                                blCreateFill = False
                                                gCreateUDTforLST lst_rst, tmLst
                                                lst_rst.MoveNext
                                                Exit Do
                                            ElseIf lst_rst!lstsplitnetwork <> "S" Then
                                                'Create Fill Lst spot
                                                Exit Do
                                            Else
                                                'Build array of regions network regions
                                                gCreateUDTforLST lst_rst, tmModelLST
                                                SQLQuery = "Select rafName from RAF_Region_Area"
                                                SQLQuery = SQLQuery & " Where (rafCode = " & lst_rst!lstRafCode & ")"
                                                Set raf_rst = gSQLSelectCall(SQLQuery)
                                                If Not raf_rst.EOF Then
                                                    slSplitNetISCI = slSplitNetISCI & "," & gAddQuotes(Trim$(tmModelLST.sISCI) & " (" & Trim$(raf_rst!rafName) & ")")
                                                Else
                                                    slSplitNetISCI = slSplitNetISCI & "," & gAddQuotes(Trim$(tmModelLST.sISCI) & " (" & "Seconadry" & ")")
                                                End If
                                                tmSplitNetRegion(UBound(tmSplitNetRegion) - 1).iNext = UBound(tmSplitNetRegion)
                                                tmSplitNetRegion(UBound(tmSplitNetRegion)).lLstCode = tmModelLST.lCode
                                                tmSplitNetRegion(UBound(tmSplitNetRegion)).iNext = -1
                                                ReDim Preserve tmSplitNetRegion(0 To UBound(tmSplitNetRegion) + 1) As SPLITNETREGION
                                            End If
                                            lst_rst.MoveNext
                                        Loop
                                        If blCreateFill Then
                                            ilRet = gCreateSplitFill(tmModelLST.iLen, 0, tmModelLST, tmLst)
                                        Else
                                            ilRet = True
                                        End If
                                        If ilRet Then
                                            If igExportSource = 2 Then DoEvents
                                            'Build tmTempRegionBreakSpot from Fill
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lBreakNo = llBreakNo
                                            ilPositionNo = ilPositionNo + 1
                                            slISCI = UCase$(Trim$(tmLst.sISCI))
                                            llAdf = gBinarySearchAdf(CLng(tmLst.iAdfCode))
                                            If llAdf <> -1 Then
                                                slAdvtName = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                                            Else
                                                slAdvtName = "Advertiser Name Missing"
                                            End If
                                            If igExportSource = 2 Then DoEvents
                                            'slAdvtName = mFilterCommas(slAdvtName)
                                            mTestISCI slISCI, slAdvtName, 0
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).iPositionNo = ilPositionNo
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLstCode = tmLst.lCode
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLogDate = llLstLogDate
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLogTime = llLstLogTime
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lSdfCode = 0
                                            tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).sISCI = slISCI
                                            ReDim Preserve tmTempRegionBreakSpots(0 To UBound(tmTempRegionBreakSpots) + 1) As REGIONBREAKSPOTS
                                            ilRegionWithInBreak = True
                                            'Output spot
                                            '7496
                                            'csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & ".MP2"
                                            csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & UCase(sgAudioExtension)
                                            slCSVRecord = Format$(llLstLogDate, "mm/dd/yy") & "," & Format$(gLongToTime(llLstLogTime), sgShowTimeWSecForm) & "," & llBreakNo & "," & lst_rst!lstPositionNo & "," & lst_rst!lstCntrNo & "," & gAddQuotes(slAdvtName) & "," & lst_rst!lstLineNo & "," & lst_rst!lstLen & "," & gAddQuotes(Trim$(lst_rst!lstISCI))
                                            'Add Primary and Seconday Info
                                            
                                            If udcCriteria.WGenCSV = vbChecked Then
                                                gLogMsgWODT "W", hmCSV, slCSVRecord & slSplitNetISCI
                                            End If
                                        Else
                                            'Output error message, Split Network ithout fills
                                            mAddMsgToList "Fill Not Found on " & slVehName & " in Break at " & Format$(llLstLogDate, "mm/dd/yy") & "," & Format$(llLstLogTime, sgShowTimeWSecForm)
                                        End If
                                    Else
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).sSource = "C"
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).iFirstSplitNetRegion = -1
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lBreakNo = llBreakNo
                                        ilPositionNo = ilPositionNo + 1
                                        slISCI = UCase$(Trim$(lst_rst!lstISCI))
                                        llAdf = gBinarySearchAdf(lst_rst!lstAdfCode)
                                        If llAdf <> -1 Then
                                            slAdvtName = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                                        Else
                                            slAdvtName = "Advertiser Name Missing"
                                        End If
                                        If igExportSource = 2 Then DoEvents
                                        'slAdvtName = mFilterCommas(slAdvtName)
                                        mTestISCI slISCI, slAdvtName, 0
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).iPositionNo = ilPositionNo
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLstCode = lst_rst!lstCode
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLogDate = llLstLogDate
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lLogTime = llLstLogTime
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).lSdfCode = lst_rst!lstSdfCode
                                        tmTempRegionBreakSpots(UBound(tmTempRegionBreakSpots)).sISCI = slISCI
                                        ReDim Preserve tmTempRegionBreakSpots(0 To UBound(tmTempRegionBreakSpots) + 1) As REGIONBREAKSPOTS
                                        'Output spot
                                        '7496
                                        'csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & ".MP2"
                                        csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & UCase(sgAudioExtension)
                                        '5/15/12: Show converted time
                                        'slCSVRecord = Format$(llLstLogDate, "mm/dd/yy") & "," & Format$(lst_rst!lstLogTime, sgShowTimeWSecForm) & "," & llBreakNo & "," & lst_rst!lstPositionNo & "," & lst_rst!lstCntrNo & "," & gAddQuotes(slAdvtName) & "," & lst_rst!lstLineNo & "," & lst_rst!lstLen & "," & gAddQuotes(Trim$(lst_rst!lstISCI))
                                        slCSVRecord = Format$(llLstLogDate, "mm/dd/yy") & "," & Format$(gLongToTime(llLstLogTime), sgShowTimeWSecForm) & "," & llBreakNo & "," & lst_rst!lstPositionNo & "," & lst_rst!lstCntrNo & "," & gAddQuotes(slAdvtName) & "," & lst_rst!lstLineNo & "," & lst_rst!lstLen & "," & gAddQuotes(Trim$(lst_rst!lstISCI))
                                        'Check if region defined for spot.  If so, retain lstLogDate, BreakNo and lstLogTime
                                        SQLQuery = "Select rsfCode, rstPtType, rsfCopyCode, rsfCrfCode, rafName from RSF_Region_Schd_Copy, RAF_Region_Area"
                                        SQLQuery = SQLQuery & " Where (rsfSdfCode = " & lst_rst!lstSdfCode
                                        SQLQuery = SQLQuery & " AND rsfType <> 'B'"     'Blackout
                                        SQLQuery = SQLQuery & " AND rsfType <> 'A'"     'Airing vehicle copy
                                        SQLQuery = SQLQuery & " AND rafType = 'C'"     'Split copy
                                        SQLQuery = SQLQuery & " AND rafCode = rsfRafCode" & ")"
                                        Set rsf_rst = gSQLSelectCall(SQLQuery)
                                        If Not rsf_rst.EOF Then
                                            ilRegionWithInBreak = True
                                            Do
                                                If igExportSource = 2 Then DoEvents
                                                ilRet = gGetCopy(rsf_rst!rstPtType, rsf_rst!rsfCopyCode, rsf_rst!rsfCrfCode, True, slRCartNo, slRProduct, slRISCI, slRCreativeTitle, llRCrfCsfCode, llRCpfCode, ilCifAdfCode)
                                                slCSVRecord = slCSVRecord & "," & gAddQuotes(slRISCI & " (" & Trim$(rsf_rst!rafName) & ")")
                                                rsf_rst.MoveNext
                                            Loop While Not rsf_rst.EOF
                                        End If
                                        If udcCriteria.WGenCSV = vbChecked Then
                                            gLogMsgWODT "W", hmCSV, slCSVRecord
                                        End If
                                    End If  'split network = P or not
                                End If  'date ok
                            End If  'spot ok , zone ok
                            If Not blSplitNetworkSpot Then
                                lst_rst.MoveNext
                            End If
                        Loop
                        If udcCriteria.WGenCSV = vbChecked Then
                            gLogMsgWODT "C", hmCSV, ""
                        End If
                        If llLogTime <> -1 Then
                            csiXMLData "CT", "Playlist_Info", ""
                        End If
                        csiXMLData "CT", "Playlist_Set", ""
                    Else
                        mAddMsgToList "No Spot Found on Affiliate for " & slVehName & " between " & slSDate & "-" & slEDate
                    End If
                    If (UBound(tmTempRegionBreakSpots) > 0) And (ilRegionWithInBreak) Then
                        For ilLoop = 0 To UBound(tmTempRegionBreakSpots) - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            tmRegionBreakSpots(UBound(tmRegionBreakSpots)) = tmTempRegionBreakSpots(ilLoop)
                            ReDim Preserve tmRegionBreakSpots(0 To UBound(tmRegionBreakSpots) + 1) As REGIONBREAKSPOTS
                        Next ilLoop
                    End If
                    'Output Breaks with region spots
                    ReDim tmCustomGroupNames(0 To 0) As CUSTOMGROUPNAMES
                    If UBound(tmRegionBreakSpots) > 0 Then
                        edcMsg.Text = "Generating Region Schedule for " & slVehName & "..."
                        If igExportSource = 2 Then DoEvents
                        'Output Region spots plus all generic spots within the break
                        mExportRegionSpot slVehGroupName, slVehName, slVehExportID, slTimeID, llEventID
                    End If
                    ilRet = csiXMLWrite(1)
                    ilRet = csiXMLEnd()
                    '7458 was file created?
                    '8886
                    'If Dir(smExportPath & slXMLFileName) > "" Then
                    If gFileExist(smExportPath & slXMLFileName) = FILEEXISTS Then
                        If Not myEnt.CreateEnts(Successful) Then
                            'gLogMsgWODT "W", hmCSV, myEnt.ErrorMessage
                            gLogMsg myEnt.ErrorMessage, "WegenerExportLog.Txt", False
                        End If
                    Else
                        If Not myEnt.CreateEnts(EntError) Then
                            'gLogMsgWODT "W", hmCSV, myEnt.ErrorMessage
                            gLogMsg myEnt.ErrorMessage, "WegenerExportLog.Txt", False
                        End If
                    End If
                    If igExportSource = 2 Then DoEvents
                    If imTerminate Then
                        mAddMsgToList "User Cancelled Export"
                        mExportSpots = False
                        Exit For
                    End If
                    edcMsg.Text = "Generating Custom Groups for " & slVehName & "..."
                    If igExportSource = 2 Then DoEvents
                    ilPos = InStr(1, slXMLFileName, "_", vbTextCompare)
                    If ilPos > 0 Then
                        slGrpName = "GrpMbr" & Mid$(slXMLFileName, ilPos)
                    Else
                        If lmGsfCode(ilGsf) = 0 Then
                            slGrpName = "GrpMbr_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & slRunLetter & ".XML"
                        Else
                            slGrpName = "GrpMbr_" & slVehExportID & "_" & slWeekNo & slYearNo & "_" & Trim$(Str$(imGameNo(ilGsf))) & "_" & slRunLetter & ".XML"
                        End If
                    End If
                    '6808
                    If Not gDeleteFile(smExportPath & slGrpName) Then
                        mAddMsgToList "Could not delete file " & slGrpName & " before writing.  Appended."
                    End If
                    '11/01/10 dan use smIniPathFileName
                    'ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "Wegener", slOutputType, smExportPath & slGrpName, sgCRLF)
                    '6807
                   ' ilRet = csiXMLStart(smIniPathFileName, "Wegener", slOutputType, smExportPath & slGrpName, sgCRLF)
                    ilRet = csiXMLStart(smIniPathFileName, "Wegener", slOutputType, smExportPath & slGrpName, sgCRLF, "")
                    ilRet = csiXMLSetMethod("", "", "", "Rx_Group_Membership")
                    If igExportSource = 2 Then DoEvents
                    csiXMLData "OT", "Interface_Header", ""
                    csiXMLData "CD", "SOURCE", "Traffic system"
                    csiXMLData "CD", "TARGET", "Compel"
                    csiXMLData "CD", "CREATION_DATETIME", Format$(slNowDT, "yyyy-mm-ddThh:mm:ss")
                    csiXMLData "CD", "DESCRIPTION", "Receiver Group Membership Assignment"
                    csiXMLData "CD", "INTERFACE_ID", "GrpMbr"
                    csiXMLData "CD", "MESSAGE_ID", slWeekNo & slYearNo & slTimeID
                    csiXMLData "CT", "Interface_Header", ""
                    If igExportSource = 2 Then DoEvents
                    ilRet = mExportGroup()
                    ilRet = csiXMLWrite(1)
                    ilRet = csiXMLEnd()
                Next ilGsf
                ilRet = gUpdateLastExportDate(imVefCode, slEDate)

'Jeff this is the end of the code to produce the second export
            Else
                mAddMsgToList "Vehicle not found " & slVehName
            End If
        End If
    Next ilVef
    'If ilAnyExports Then
    '    ilRet = csiXMLEnd()
    'End If
    'D.S. 11/17/17 - TTP #8687
    If bgTaskBlocked And igExportSource <> 2 Then
        gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_mmddyyyy.txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    
    On Error Resume Next
    lst_rst.Close
    vff_rst.Close
    rsf_rst.Close
    raf_rst.Close
    mClearAlerts llSDate, llEDate
    mExportSpots = True
    '7458
    Set myEnt = Nothing
    Exit Function
mExportSpotsErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerExportLog.txt", "Export Wegener-mExportSpots"
    mExportSpots = False
    Resume Next
    Exit Function
    
End Function



Function mFilterCommas(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    'Remove " and '
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, ",", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = " "
            ilFound = True
        End If
    Loop While ilFound
    mFilterCommas = slName
End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload FrmExportWegener
End Sub

Private Sub txtNumberDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtRunLetter_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub mExportRegionSpot(slVehGroupName As String, slVehName As String, slVehExportID As String, slTimeID As String, llEventID As Long)
    'Export any break that has region copy.  All spots with that break must be export with
    'each unique region defined for the spot
    'Spot two defined as:
    'Fmt1 and St1 and Not K111
    'Fmt1 and St2 and Not K111
    'The break has three spots so the following would be exported
    'assuming the region spot is the middle spot
    '  ISCI for spot1
    '  ISCI for Region Fmt1 and St1 and Not K111
    '  ISCI for spot3
    '
    '  ISCI for spot1
    '  ISCI for Region Fmt1 and St2 and Not K111
    '  ISCI for spot3
    '
    'The key to this module is forming all the possible combinations of regions
    'formed by the intersection of all the regions defined for each spot within a break
    '
    'Shown below is how a came up with the technique to form the combinations
    'Form all combination of regions by creating a matrix (spot # vs total # of combination regions)
    'To compute the total number of combination regions, three major formula's required
    'Formula 1:  Compute number of independent regions.
    '            This is the summ of the number of regions across all spots in the break
    'Formula 2:  Compute the total number of combination of regions
    '            This is multi-pass procedure
    '            Spot1#Regions x (Spot2#Regions+Spot3#regions+....+ Spotn#Regions)
    '            Spot2#Regions x (Spot3#Regions+Spot4#Regions+...+Spotn#Regions)
    '            Spot3#Regions x (Spot4#Regions+...+ Spotn#Regions)
    '            Etc
    'Formula 3:  Compute intersection of all regions
    '            Spot1#Region x Spot2#Region x Spot3#Region x ...... x Spotn#Region
    '
    'Example 3 Spots:  Spot 1 has 3 regions, spot 2 has 2 regions and spot 3 has 1 region
    '        Formula 1 yields (3 + 2+ 1):
    '        Spot 1:   R1  R2  R3  -   -
    '        Spot 2:   -   -   -   R4  R5
    '        Spot 3:   -   -   -   -   -  R6
    '        Formula 2 yields 3 * (2 + 1)
    '        Spot 1:   R1  R1  R1  R2  R2  R2  R3  R3  R3
    '        Spot 2:   R4  R5  -   R4  R5  -   R4  R5
    '        Spot 3:   -   -   R6  -   -   R6  -   -   R6
    '        Note that the first repeat is different then how the other row repeat
    '        Formula 2 yields 2 * (1)
    '        Spot 1:   -   -
    '        Spot 2:   R4  R5
    '        Spot 3:   R6  R6
    '        Formula 3 yields 3 * 2 * 1
    '        Spot 1:   R1  R2  R3  R1  R2  R3
    '        Spot 2:   R4  R5  R4  R5  R4  R5
    '        Spot 3:   R6  R6  R6  R6  R6  R6
    'Example 3 Spots:  Spot 1 has 1 region, Spot 2 has 2 regions and Spot 3 has 3 regions
    '        Spot 1:   R1  -   -   -   -
    '        Spot 2:   -   R2  R3
    '        Spot 3:   -   -   -   R4  R5  R6
    '        Formula 2 yields 1 * (2 + 3)
    '        Spot 1:   R1  R1  R1  R1  R1
    '        Spot 2:   R2  R3  -   -   -
    '        Spot 3:   -   -   R4  R5  R6
    '        Note that the first repeat is different then how the other row repeat
    '        Formula 2 yields 2 * (3)
    '        Spot 1:   -   -   -   -   -
    '        Spot 2:   R2  R2  R2  R3  R3  R3
    '        Spot 3:   R4  R5  R6  R4  R5  R6
    '        Formula 3 yields 3 * 2 * 1
    '        Spot 1:   R1  R1  R1  R1  R1  R1
    '        Spot 2:   R2  R3  R2  R3  R2  R3
    '        Spot 3:   R4  R5  R6  R4  R5  R6

    Dim ilIndex As Integer
    Dim ilStartBreakIndex As Integer
    Dim ilEndBreakIndex As Integer
    Dim ilFoundRegion As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilMove As Integer
    Dim ilAdj As Integer
    'Dim ilNoCells As Integer
    Dim llNoCells As Long
    Dim llNoCycles As Long
    Dim ilRow As Integer
    'Dim ilCol As Integer
    Dim llCol As Long
    Dim llCycle As Long
    Dim llCycleLoop As Long
    Dim llRepeat As Long
    Dim llRepeatFactor As Long
    'Dim ilSum As Integer
    Dim llSum As Long
    'Dim ilProduct As Integer
    Dim llProduct As Long
    Dim llRegionIndex As Long
    Dim ilRepeatStartRow As Integer
    Dim ilRepeatEndRow As Integer
    Dim ilPass As Integer
    Dim ilStationWithinBreak As Integer
    Dim slCartNo As String
    Dim slProduct As String
    Dim slISCI As String
    Dim slCreativeTitle As String
    Dim llCrfCsfCode As Long
    Dim llCrfCode As Long
    Dim llCpfCode As Long
    Dim ilStationFound As Integer
    Dim slEventID As String
    Dim ilNoRegions As Integer
    Dim ilNoProductRows As Integer
    Dim ilPositionNo As Integer
    Dim ilGroupNo As Integer
    Dim slGroup As String
    Dim slGroupA As String
    Dim slGroupB As String
    Dim slGroupC As String
    Dim slGroupD As String
    Dim slGroupLetter As String
    Dim ilCifAdfCode As Integer
    Dim ilSvRow As Integer
    Dim slRegionError As String
    '12/1/17: Test max region combinations in break
    Dim dlMaxRegionDef As Double
    '3/3/18
    Dim ilStnRow As Integer
    Dim llStnCol As Long
    Dim llStationCount As Long
    Dim llStationOtherFirst As Long
    Dim ilStnOuter As Integer
    Dim ilStnInner As Integer
    Dim llStnNext As Long
    Dim ilRowWithStation As Integer
    Dim blFound As Boolean
    Dim llSerialNo As Long
    On Error GoTo ErrHand
    
    '12/1/17: Test max region combinations in break
    dlMaxRegionDef = 10000000
    ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    '3/3/18
    ReDim tmStationSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim imRowWithStation(0 To 0) As Integer

    ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    '3/3/18
    mBuildAllowStationList

    ilIndex = 0
    Do
        If igExportSource = 2 Then DoEvents
        If imTerminate Then
            mAddMsgToList "User Cancelled Export"
            Exit Sub
        End If
        ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
        ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
        ReDim tmStationSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO

        ilStartBreakIndex = ilIndex
        ilEndBreakIndex = ilIndex
        Do
            If igExportSource = 2 Then DoEvents
            If (tmRegionBreakSpots(ilIndex).lLogDate = tmRegionBreakSpots(ilIndex + 1).lLogDate) And (tmRegionBreakSpots(ilIndex).lLogTime = tmRegionBreakSpots(ilIndex + 1).lLogTime) Then
                ilIndex = ilIndex + 1
                ilEndBreakIndex = ilIndex
            Else
                ilIndex = ilIndex + 1
                Exit Do
            End If
        Loop While ilIndex < UBound(tmRegionBreakSpots)
        ilFoundRegion = False
        ReDim tlRegionBreakSpotInfo(0 To ilEndBreakIndex - ilStartBreakIndex + 1) As REGIONBREAKSPOTINFO
        For ilLoop = ilStartBreakIndex To ilEndBreakIndex Step 1
            If igExportSource = 2 Then DoEvents
            tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lSdfCode = tmRegionBreakSpots(ilLoop).lSdfCode
            tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).sISCI = tmRegionBreakSpots(ilLoop).sISCI
            tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).iPositionNo = tmRegionBreakSpots(ilLoop).iPositionNo
            If tmRegionBreakSpots(ilLoop).sSource <> "N" Then
                ilRet = gBuildRegionDefinitions("W", tmRegionBreakSpots(ilLoop).lSdfCode, imVefCode, tlRegionDefinition(), tlSplitCategoryInfo())
            Else
                ilRet = gBuildSplitNetRegionDefinitions(tmRegionBreakSpots(ilLoop).iFirstSplitNetRegion, tmSplitNetRegion(), tlRegionDefinition(), tlSplitCategoryInfo())
            End If
            If ilRet Then
                'Form unique region for each OR combination
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lStartIndex = UBound(tmRegionDefinition)
                '3/3/18
                mFilterStations tlRegionDefinition(), tlSplitCategoryInfo()
                '3/3/18
                'gSeparateRegions tlRegionDefinition(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo()
                gSeparateRegions tlRegionDefinition(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo(), False
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lEndIndex = UBound(tmRegionDefinition) - 1
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).iNoRegions = tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lEndIndex - tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lStartIndex + 1
                '3/3/18: Save station categories
                mSaveStationCategory tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex), tlSplitCategoryInfo()
            Else
                'Output generic copy
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lStartIndex = -1
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).lEndIndex = -1
                tlRegionBreakSpotInfo(ilLoop - ilStartBreakIndex).iNoRegions = 0
            End If
        Next ilLoop
        
        ilRepeatStartRow = 0
        ilRepeatEndRow = -1
        ilNoRegions = 0
        'Remove rows without regions.  This is required to enable the formulas to work correctly
        ReDim tlSvRegionBreakSpotInfo(0 To UBound(tlRegionBreakSpotInfo)) As REGIONBREAKSPOTINFO
        ReDim ilMatrixPositionInfo(0 To UBound(tlRegionBreakSpotInfo)) As Integer
        For ilRow = 0 To UBound(tlSvRegionBreakSpotInfo) - 1 Step 1
            tlSvRegionBreakSpotInfo(ilRow) = tlRegionBreakSpotInfo(ilRow)
        Next ilRow
        For ilRow = 0 To UBound(tlSvRegionBreakSpotInfo) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If tlSvRegionBreakSpotInfo(ilRow).iNoRegions > 0 Then
                tlRegionBreakSpotInfo(ilNoRegions) = tlSvRegionBreakSpotInfo(ilRow)
                ilMatrixPositionInfo(ilNoRegions) = tlSvRegionBreakSpotInfo(ilRow).iPositionNo
                ilNoRegions = ilNoRegions + 1
            End If
        Next ilRow
        ReDim Preserve tlRegionBreakSpotInfo(0 To ilNoRegions) As REGIONBREAKSPOTINFO
        ReDim Preserve ilMatrixPositionInfo(0 To ilNoRegions) As Integer
        ilPass = 1
        Do
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                Exit Sub
            End If
            llSum = 0
            llProduct = 1
            ilNoProductRows = 0
            For ilRow = 0 To UBound(tlRegionBreakSpotInfo) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If ilRow > ilRepeatEndRow Then
                    llSum = llSum + tlRegionBreakSpotInfo(ilRow).iNoRegions
                ElseIf ilRow >= ilRepeatStartRow Then
                    '12/1/17: Test max region combinations in break
                    'llProduct = llProduct * tlRegionBreakSpotInfo(ilRow).iNoRegions
                    If CDbl(llProduct) < dlMaxRegionDef Then
                        llProduct = llProduct * tlRegionBreakSpotInfo(ilRow).iNoRegions
                    Else
                        llProduct = dlMaxRegionDef
                    End If
                    ilNoProductRows = ilNoProductRows + 1
                End If
            Next ilRow
            If (ilRepeatEndRow = -1) Then
                '12/1/17: Test max region combinations in break
                'llNoCells = llProduct * llSum
                If CDbl(llProduct) * CDbl(llSum) < dlMaxRegionDef Then
                    llNoCells = llProduct * llSum
                Else
                    'Add message
                    llNoCells = 0
                    ilNoRegions = 0
                    mGenRegionErrorMsg ilStartBreakIndex, ilEndBreakIndex
                End If
            ElseIf ilPass = ilNoProductRows Then
                '12/1/17: Test max region combinations in break
                'llNoCells = llProduct * llSum
                If CDbl(llProduct) * CDbl(llSum) < CDbl(dlMaxRegionDef) Then
                    llNoCells = llProduct * llSum
                Else
                    'Add message
                    llNoCells = 0
                    ilNoRegions = 0
                    mGenRegionErrorMsg ilStartBreakIndex, ilEndBreakIndex
                End If
            Else
                llNoCells = 0
            End If
            ReDim llMatrix(0 To UBound(tlRegionBreakSpotInfo), 0 To llNoCells) As Long
            If llNoCells > 0 Then
                If igExportSource = 2 Then DoEvents
                For ilRow = 0 To UBound(llMatrix, 1) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    For llCol = 0 To UBound(llMatrix, 2) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        llMatrix(ilRow, llCol) = -1   'No Region
                    Next llCol
                Next ilRow
                llNoCycles = 1
                For ilRow = ilRepeatStartRow To ilRepeatEndRow Step 1
                    If igExportSource = 2 Then DoEvents
                    llRepeatFactor = llNoCells / tlRegionBreakSpotInfo(ilRow).iNoRegions
                    llCol = 0
                    For llCycle = 1 To llNoCycles Step 1
                        If igExportSource = 2 Then DoEvents
                        For llRegionIndex = tlRegionBreakSpotInfo(ilRow).lStartIndex To tlRegionBreakSpotInfo(ilRow).lEndIndex Step 1
                            If igExportSource = 2 Then DoEvents
                            For llRepeat = 1 To llRepeatFactor Step 1
                                If igExportSource = 2 Then DoEvents
                                llMatrix(ilRow, llCol) = llRegionIndex
                                llCol = llCol + 1
                            Next llRepeat
                        Next llRegionIndex
                    Next llCycle
                    llNoCells = llRepeatFactor
                    llNoCycles = UBound(llMatrix, 2) / llNoCells
                Next ilRow
                llCol = 0
                For llCycleLoop = 0 To llNoCycles - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    For ilRow = ilRepeatEndRow + 1 To UBound(llMatrix, 1) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        For llRegionIndex = tlRegionBreakSpotInfo(ilRow).lStartIndex To tlRegionBreakSpotInfo(ilRow).lEndIndex Step 1
                            If igExportSource = 2 Then DoEvents
                            llMatrix(ilRow, llCol) = llRegionIndex
                            llCol = llCol + 1
                        Next llRegionIndex
                    Next ilRow
                Next llCycleLoop
                'Create Break Info
                ilStnRow = -1
                llStnCol = -1
                ilStationFound = False
                ReDim ilRegionWithStation(0 To UBound(llMatrix, 1)) As Integer
                For llCol = LBound(llMatrix, 2) To UBound(llMatrix, 2) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If imTerminate Then
                        mAddMsgToList "User Cancelled Export"
                        Exit Sub
                    End If
                    
                    
                    '3/3/18: Determine if stations not split out
                    ReDim imRowWithStation(0 To 0) As Integer
                    For ilRow = LBound(llMatrix, 1) To UBound(llMatrix, 1) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        If llMatrix(ilRow, llCol) >= 0 Then
                            If tmRegionDefinition(llMatrix(ilRow, llCol)).iStationCount > 0 Then
                                imRowWithStation(UBound(imRowWithStation)) = ilRow
                                ReDim Preserve imRowWithStation(0 To UBound(imRowWithStation) + 1) As Integer
                                'added to remove duplication of breaks with same definition
                                Exit For
                            End If
                        End If
                    Next ilRow
                    If UBound(imRowWithStation) = 0 Then
                        imRowWithStation(0) = -1
                        ReDim Preserve imRowWithStation(0 To 1) As Integer
                    End If
                    For ilRowWithStation = 0 To UBound(imRowWithStation) - 1 Step 1
                        llStationCount = 1
                        llStationOtherFirst = -1
                        ilRow = imRowWithStation(ilRowWithStation)
                        If ilRow >= 0 Then
                            'For ilRow = LBound(llMatrix, 1) To UBound(llMatrix, 1) - 1 Step 1
                                If igExportSource = 2 Then DoEvents
                            '    If llMatrix(ilRow, llCol) >= 0 Then
                            '        If tmRegionDefinition(llMatrix(ilRow, llCol)).iStationCount > 0 Then
                                        ilStnRow = ilRow
                                        llStnCol = llCol
                                        llStationCount = tmRegionDefinition(llMatrix(ilRow, llCol)).iStationCount
                                        llStationOtherFirst = tmRegionDefinition(llMatrix(ilRow, llCol)).lStationOtherFirst
                            '            Exit For
                            '        End If
                            '    End If
                            'Next ilRow
                        End If
                        For ilStnOuter = 0 To llStationCount - 1 Step 1
                            If llStationOtherFirst <> -1 Then
                                llStnNext = llStationOtherFirst
                                'For ilStnInner = 0 To ilStnOuter Step 1
                                    tmSplitCategoryInfo(tmRegionDefinition(llMatrix(ilStnRow, llStnCol)).lOtherFirst) = tmStationSplitCategoryInfo(llStnNext)   'tlSplitCategoryInfo(llStnNext)
                                    tmSplitCategoryInfo(tmRegionDefinition(llMatrix(ilStnRow, llStnCol)).lOtherFirst).lNext = -1
                                    llStnNext = tmStationSplitCategoryInfo(llStnNext).lNext   'tlSplitCategoryInfo(llStnNext).lNext
                                'Next ilStnInner
                                llStationOtherFirst = llStnNext
                            End If
                            
                            ilStationWithinBreak = True
                            ReDim tmMergeRegionDefinition(0 To 0) As REGIONDEFINITION
                            ReDim tmMergeSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
                            For ilRow = LBound(llMatrix, 1) To UBound(llMatrix, 1) - 1 Step 1
                                If igExportSource = 2 Then DoEvents
                                If llMatrix(ilRow, llCol) >= 0 Then
                                    ilSvRow = ilRow
                                    ilRet = mMergeCategory(llMatrix(ilRow, llCol), ilRepeatEndRow)
                                    If Not ilRet Then
                                        ilStationWithinBreak = False
                                        Exit For
                                    End If
                                End If
                            Next ilRow
                            If ilStationWithinBreak Then
                                ilStationWithinBreak = mAnyStations(slGroupA, slGroupB, slGroupC, slGroupD, llSerialNo)
                                If Not ilStationWithinBreak Then
                                    slRegionError = "Empty Region "
                                End If
                            Else
                                'Invalid region
                                slRegionError = "Region structure error "
                            End If
                            If Not ilStationWithinBreak Then
                               '3/3/18
                                If llStationCount <= 0 Then
                                    If (ilNoRegions <= 1) Or (ilRepeatEndRow = -1) Then
                                        llRegionIndex = llMatrix(ilSvRow, llCol)
                                        slRegionError = slRegionError & Trim$(tmRegionDefinition(llRegionIndex).sRegionName)
                                        'SQLQuery = "SELECT chfCntrNo, adfName "
                                        'SQLQuery = SQLQuery + " FROM chf_contract_Header, adf_Advertisers, sdf_Spot_Detail"
                                        'SQLQuery = SQLQuery + " WHERE (sdfCode = " & tmRegionDefinition(llRegionIndex).lSdfCode
                                        'SQLQuery = SQLQuery + " AND chfCode = sdfChfCode AND adfCode = sdfadfCode " & ")"
                                        'Set err_rst = gSQLSelectCall(SQLQuery)
                                        'If Not err_rst.EOF Then
                                        '    slRegionError = slRegionError & " Advertiser " & Trim$(err_rst!adfName) & " Contract # " & err_rst!chfCntrNo
                                        'End If
                                        'slRegionError = slRegionError & " " & slVehName & " Date " & Format$(tmRegionBreakSpots(ilStartBreakIndex).lLogDate, "m/d/yy") & " Break # " & Trim$(Str$(tmRegionBreakSpots(ilStartBreakIndex).lBreakNo)) & " Position # " & Trim$(Str$(tlRegionBreakSpotInfo(ilSvRow).iPositionNo))
                                        If igExportSource = 2 Then DoEvents
                                        SQLQuery = "SELECT chfCntrNo, adfName "
                                        SQLQuery = SQLQuery + " FROM chf_contract_Header, adf_Advertisers, crf_Copy_Rot_Header"
                                        SQLQuery = SQLQuery + " WHERE (crfCode = " & tmRegionDefinition(llRegionIndex).lCrfCode
                                        SQLQuery = SQLQuery + " AND chfCode = crfChfCode AND adfCode = crfadfCode " & ")"
                                        Set err_rst = gSQLSelectCall(SQLQuery)
                                        If Not err_rst.EOF Then
                                            slRegionError = slRegionError & " Advertiser " & Trim$(err_rst!adfName) & " Contract # " & err_rst!chfCntrNo
                                        End If
                                        mAddMsgToList slRegionError
                                        '3/3/18
                                        If ilStnOuter = llStationCount - 1 Then
                                            'Since this region has no stations, turn the region off
                                            tmRegionDefinition(llRegionIndex).lFormatFirst = -1
                                            tmRegionDefinition(llRegionIndex).lOtherFirst = -1
                                            tmRegionDefinition(llRegionIndex).lExcludeFirst = -1
                                        End If
                                    End If
                                End If
                            End If
                            If ilStationWithinBreak Then
                                If igExportSource = 2 Then DoEvents
                                If imTerminate Then
                                    mAddMsgToList "User Cancelled Export"
                                    Exit Sub
                                End If
                                For ilGroupNo = 0 To 3 Step 1
                                    If igExportSource = 2 Then DoEvents
                                    slGroupLetter = ""
                                    Select Case ilGroupNo
                                        Case 0
                                            If slGroupA <> "" Then
                                                slGroup = slGroupA
                                                slGroupLetter = "A"
                                            End If
                                        Case 1
                                            If slGroupB <> "" Then
                                                slGroup = slGroupB
                                                slGroupLetter = "B"
                                            End If
                                        Case 2
                                            If slGroupC <> "" Then
                                                slGroup = slGroupC
                                                slGroupLetter = "C"
                                            End If
                                        Case 3
                                            If slGroupD <> "" Then
                                                slGroup = slGroupD
                                                slGroupLetter = "D"
                                            End If
                                    End Select
                                    If slGroupLetter <> "" Then
                                        If igExportSource = 2 Then DoEvents
                                        ilStationFound = True
                                        csiXMLData "OT", "Playlist_Set", ""
                                        'Replace slVehGroupName with Region definition
                                        csiXMLData "CD", "Address", mFormRegionAddress(slVehGroupName, slGroup, slGroupLetter, llSerialNo)
                                        csiXMLData "CD", "Vehicle_Name", gXMLNameFilter(slVehName)
                                        csiXMLData "CD", "Vehicle_Code", slVehExportID
                                        csiXMLData "CD", "Day_of_Week", UCase$(Left$(Format$(tmRegionBreakSpots(ilStartBreakIndex).lLogDate, "ddd"), 2))
                                        llEventID = llEventID + 1
                                        slEventID = Trim$(Str$(llEventID))
                                        Do While Len(slEventID) < 5
                                            slEventID = "0" & slEventID
                                        Loop
                                        csiXMLData "OT", "Playlist_Info", "Id=" & """" & slTimeID & slEventID & """"
                                        csiXMLData "CD", "Break_Number", Trim$(Str$(tmRegionBreakSpots(ilStartBreakIndex).lBreakNo))
                                        For ilRow = 0 To UBound(tlSvRegionBreakSpotInfo) - 1 Step 1
                                            If igExportSource = 2 Then DoEvents
                                            ilPositionNo = tlSvRegionBreakSpotInfo(ilRow).iPositionNo
                                            slISCI = Trim$(tlSvRegionBreakSpotInfo(ilRow).sISCI)
                                            For ilLoop = 0 To UBound(ilMatrixPositionInfo) - 1 Step 1
                                                If igExportSource = 2 Then DoEvents
                                                If ilMatrixPositionInfo(ilLoop) = ilPositionNo Then
                                                    llRegionIndex = llMatrix(ilLoop, llCol)
                                                    If llMatrix(ilLoop, llCol) >= 0 Then
                                                        ilRet = gGetCopy(tmRegionDefinition(llRegionIndex).sPtType, tmRegionDefinition(llRegionIndex).lCopyCode, tmRegionDefinition(llRegionIndex).lCrfCode, True, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode, ilCifAdfCode)
                                                        If Not ilRet Then
                                                            slISCI = Trim$(tlSvRegionBreakSpotInfo(ilRow).sISCI)
                                                        End If
                                                        If igExportSource = 2 Then DoEvents
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            slISCI = UCase$(Trim$(slISCI))
                                            mTestISCI slISCI, "", tlSvRegionBreakSpotInfo(ilRow).lSdfCode
                                            '7496
                                            'csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & ".MP2"
                                            csiXMLData "CD", "File_Path", smMP2FilePath & slISCI & UCase(sgAudioExtension)
                                        Next ilRow
                                        csiXMLData "CT", "Playlist_Info", ""
                                        csiXMLData "CT", "Playlist_Set", ""
                                    End If
                                Next ilGroupNo
                            End If
                        '3/3/18
                        Next ilStnOuter
                    Next ilRowWithStation
                Next llCol
            End If
            If ilNoRegions <= 1 Then
                Exit Do
            End If
            If ilRepeatEndRow = -1 Then
                ilRepeatEndRow = 0
            Else
                ilRepeatEndRow = ilRepeatEndRow + 1
                If ilRepeatEndRow >= UBound(llMatrix, 1) - 1 Then
                    ilRepeatStartRow = 0
                    ilRepeatEndRow = ilRepeatStartRow + ilPass
                    ilPass = ilPass + 1
                Else
                    ilRepeatStartRow = ilRepeatStartRow + 1
                End If
            End If
        Loop While ilRepeatEndRow < UBound(llMatrix, 1) - 1
    
    Loop While ilIndex < UBound(tmRegionBreakSpots)
    Erase tmRegionDefinition
    Erase tmSplitCategoryInfo
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerExportLog.txt", "Export Wegener-mExportRegionSpot"
    Resume Next
    Exit Sub

End Sub

Private Function mAnyStations(slGroupInfoA As String, slGroupInfoB As String, slGroupInfoC As String, slGroupInfoD As String, llSerialNo As Long) As Integer
    'Add loop thru Wegener stations to determine if any station meets the region definition criteria
    'Region defined as Fmt1 and St1 and Not K111.  Look to see if any station will air within this region
    'This is used to avoid exporting regions that have no stations associated with it as Wegener will reject this region and all
    'other commands after this rejected command
    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim ilShttCode As Integer
    Dim ilMktCode As Integer
    Dim ilMSAMktCode As Integer
    Dim slState As String
    Dim ilFmtCode As Integer
    Dim ilTztCode As Integer
    Dim llRegion As Long
    Dim slGroupInfo As String
    Dim ilTest As Integer
    Dim ilPos As Integer
    Dim ilFind As Integer
    Dim ilImport As Integer
    Dim llVefIndex As Long
    Dim slPort As String
    Dim slCategory As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    
    'ReDim ilWegenerIndex(0 To 0) As Integer
    ReDim tmWegenerInfoIndex(0 To 0) As WEGENERINFOINDEX
    
    On Error GoTo ErrHandle
    
    slGroupInfoA = ""
    slGroupInfoB = ""
    slGroupInfoC = ""
    slGroupInfoD = ""
    'This code only works if only one Region definition exist
    If UBound(tmMergeRegionDefinition) <= 1 Then
        For llRegion = 0 To UBound(tmMergeRegionDefinition) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If tmMergeRegionDefinition(llRegion).lOtherFirst <> -1 Then
                llOtherIndex = tmMergeRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llOtherIndex).sCategory = "S" Then
                        For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            If tmMergeSplitCategoryInfo(llOtherIndex).iIntCode = tmWegenerImport(ilImport).iShttCode Then
                                If tmMergeSplitCategoryInfo(llOtherIndex).lLongCode = Val(tmWegenerImport(ilImport).sSerialNo1) Then
                                    'ilWegenerIndex(UBound(ilWegenerIndex)) = ilImport
                                    'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                    tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = ilImport
                                    tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(ilImport).sSerialNo1)
                                    ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
            
                                End If
                            End If
                        Next ilImport
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llOtherIndex = tmMergeSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
                    
                llOtherIndex = tmMergeRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llOtherIndex).sCategory = "M" Then
                        ilRet = mBinarySearchMarket(tmMergeSplitCategoryInfo(llOtherIndex).iIntCode)
                        If ilRet <> -1 Then
                            ilIndex = tmWegenerMarketSort(ilRet).iFirst
                            Do While ilIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                'ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
                                'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = tmWegenerIndex(ilIndex).iIndex
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(tmWegenerIndex(ilIndex).iIndex).sSerialNo1)
                                ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
                                ilIndex = tmWegenerIndex(ilIndex).iNext
                            Loop
                        End If
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llOtherIndex = tmMergeSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
                
                llOtherIndex = tmMergeRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llOtherIndex).sCategory = "A" Then
                        ilRet = mBinarySearchMSAMarket(tmMergeSplitCategoryInfo(llOtherIndex).iIntCode)
                        If ilRet <> -1 Then
                            ilIndex = tmWegenerMSAMarketSort(ilRet).iFirst
                            Do While ilIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                'ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
                                'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = tmWegenerIndex(ilIndex).iIndex
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(tmWegenerIndex(ilIndex).iIndex).sSerialNo1)
                                ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
                                ilIndex = tmWegenerIndex(ilIndex).iNext
                            Loop
                        End If
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llOtherIndex = tmMergeSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
                
                
                llOtherIndex = tmMergeRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llOtherIndex).sCategory = "N" Then
                        ilRet = mBinarySearchPostalName(Trim$(tmMergeSplitCategoryInfo(llOtherIndex).sName))
                        If ilRet <> -1 Then
                            ilIndex = tmWegenerPostalSort(ilRet).iFirst
                            Do While ilIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                'ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
                                'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = tmWegenerIndex(ilIndex).iIndex
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(tmWegenerIndex(ilIndex).iIndex).sSerialNo1)
                                ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
                                ilIndex = tmWegenerIndex(ilIndex).iNext
                            Loop
                        End If
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llOtherIndex = tmMergeSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
                
                llOtherIndex = tmMergeRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llOtherIndex).sCategory = "T" Then
                        ilRet = mBinarySearchTimeZone(tmMergeSplitCategoryInfo(llOtherIndex).iIntCode)
                        If ilRet <> -1 Then
                            ilIndex = tmWegenerTimeZoneSort(ilRet).iFirst
                            Do While ilIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                'ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
                                'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = tmWegenerIndex(ilIndex).iIndex
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(tmWegenerIndex(ilIndex).iIndex).sSerialNo1)
                                ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
                                ilIndex = tmWegenerIndex(ilIndex).iNext
                            Loop
                        End If
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llOtherIndex = tmMergeSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
            End If
            If tmMergeRegionDefinition(llRegion).lFormatFirst <> -1 Then
                llFormatIndex = tmMergeRegionDefinition(llRegion).lFormatFirst
                Do
                    If igExportSource = 2 Then DoEvents
                    If tmMergeSplitCategoryInfo(llFormatIndex).sCategory = "F" Then
                        ilRet = mBinarySearchFormat(tmMergeSplitCategoryInfo(llFormatIndex).iIntCode)
                        If ilRet <> -1 Then
                            ilIndex = tmWegenerFormatSort(ilRet).iFirst
                            Do While ilIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                'ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
                                'ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).iImport = tmWegenerIndex(ilIndex).iIndex
                                tmWegenerInfoIndex(UBound(tmWegenerInfoIndex)).lSerialNo = Val(tmWegenerImport(tmWegenerIndex(ilIndex).iIndex).sSerialNo1)
                                ReDim Preserve tmWegenerInfoIndex(0 To UBound(tmWegenerInfoIndex) + 1) As WEGENERINFOINDEX
                                ilIndex = tmWegenerIndex(ilIndex).iNext
                            Loop
                        End If
                        'If UBound(ilWegenerIndex) > 0 Then
                        If UBound(tmWegenerInfoIndex) > 0 Then
                            Exit For
                        End If
                    End If
                    llFormatIndex = tmMergeSplitCategoryInfo(llFormatIndex).lNext
                Loop While llFormatIndex <> -1
            End If
            If tmMergeRegionDefinition(llRegion).lOtherFirst = -1 Then
                If tmMergeRegionDefinition(llRegion).lFormatFirst = -1 Then
                    If tmMergeRegionDefinition(llRegion).lExcludeFirst <> -1 Then
'                        llExcludeIndex = tmMergeRegionDefinition(llRegion).lExcludeFirst
'                        Do
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "S" Then
'                                For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
'                                    If tmMergeSplitCategoryInfo(llExcludeIndex).iIntCode = tmWegenerImport(ilImport).iShttCode Then
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = ilImport
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                    End If
'                                Next ilImport
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "M" Then
'                                ilRet = mBinarySearchMarket(tmMergeSplitCategoryInfo(llExcludeIndex).iIntCode)
'                                If ilRet <> -1 Then
'                                    ilIndex = tmWegenerMarketSort(ilRet).iFirst
'                                    Do While ilIndex <> -1
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                        ilIndex = tmWegenerIndex(ilIndex).iNext
'                                    Loop
'                                End If
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "A" Then
'                                ilRet = mBinarySearchMSAMarket(tmMergeSplitCategoryInfo(llExcludeIndex).iIntCode)
'                                If ilRet <> -1 Then
'                                    ilIndex = tmWegenerMSAMarketSort(ilRet).iFirst
'                                    Do While ilIndex <> -1
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                        ilIndex = tmWegenerIndex(ilIndex).iNext
'                                    Loop
'                                End If
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "N" Then
'                                ilRet = mBinarySearchPostalName(Trim$(tmMergeSplitCategoryInfo(llExcludeIndex).sName))
'                                If ilRet <> -1 Then
'                                    ilIndex = tmWegenerPostalSort(ilRet).iFirst
'                                    Do While ilIndex <> -1
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                        ilIndex = tmWegenerIndex(ilIndex).iNext
'                                    Loop
'                                End If
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "T" Then
'                                ilRet = mBinarySearchTimeZone(tmMergeSplitCategoryInfo(llExcludeIndex).iIntCode)
'                                If ilRet <> -1 Then
'                                    ilIndex = tmWegenerTimeZoneSort(ilRet).iFirst
'                                    Do While ilIndex <> -1
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                        ilIndex = tmWegenerIndex(ilIndex).iNext
'                                    Loop
'                                End If
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            If tmMergeSplitCategoryInfo(llExcludeIndex).sCategory = "F" Then
'                                ilRet = mBinarySearchFormat(tmMergeSplitCategoryInfo(llExcludeIndex).iIntCode)
'                                If ilRet <> -1 Then
'                                    ilIndex = tmWegenerFormatSort(ilRet).iFirst
'                                    Do While ilIndex <> -1
'                                        ilWegenerIndex(UBound(ilWegenerIndex)) = tmWegenerIndex(ilIndex).iIndex
'                                        ReDim Preserve ilWegenerIndex(0 To UBound(ilWegenerIndex) + 1) As Integer
'                                        ilIndex = tmWegenerIndex(ilIndex).iNext
'                                    Loop
'                                End If
'                                If UBound(ilWegenerIndex) > 0 Then
'                                    Exit For
'                                End If
'                            End If
'                            llExcludeIndex = tmMergeSplitCategoryInfo(llExcludeIndex).lNext
'                        Loop While llExcludeIndex <> -1
                        'ReDim ilWegenerIndex(0 To UBound(tmWegenerImport)) As Integer
                        ReDim tmWegenerInfoIndex(0 To UBound(tmWegenerImport)) As WEGENERINFOINDEX
                        For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            'ilWegenerIndex(ilImport) = ilImport
                            tmWegenerInfoIndex(ilImport).iImport = ilImport
                            tmWegenerInfoIndex(ilImport).lSerialNo = Val(tmWegenerImport(ilImport).sSerialNo1)
                        Next ilImport
                    End If
                End If
            End If

        Next llRegion
    Else
        'ReDim ilWegenerIndex(0 To UBound(tmWegenerImport)) As Integer
        ReDim tmWegenerInfoIndex(0 To UBound(tmWegenerImport)) As WEGENERINFOINDEX
        For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
            'ilWegenerIndex(ilImport) = ilImport
            tmWegenerInfoIndex(ilImport).iImport = ilImport
            tmWegenerInfoIndex(ilImport).lSerialNo = Val(tmWegenerImport(ilImport).sSerialNo1)
        Next ilImport
    End If
    'For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
    'For ilIndex = 0 To UBound(ilWegenerIndex) - 1 Step 1
    For ilIndex = 0 To UBound(tmWegenerInfoIndex) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        'ilImport = ilWegenerIndex(ilIndex)
        ilImport = tmWegenerInfoIndex(ilIndex).iImport
        If Val(tmWegenerImport(ilImport).sSerialNo1) = tmWegenerInfoIndex(ilIndex).lSerialNo Then
        'If tmWegenerImport(ilImport).iShttCode = ilShttCode Then
            ilShttCode = tmWegenerImport(ilImport).iShttCode
            ilMktCode = tmWegenerImport(ilImport).iMktCode
            ilMSAMktCode = tmWegenerImport(ilImport).iMSAMktCode
            slState = tmWegenerImport(ilImport).sPostalName
            ilFmtCode = tmWegenerImport(ilImport).iFormatCode
            ilTztCode = tmWegenerImport(ilImport).iTztCode
            llVefIndex = tmWegenerImport(ilImport).lVefCodeFirst
            'llSerialNo = tmWegenerImport(ilImport).sSerialNo1
            Do While llVefIndex <> -1
                If igExportSource = 2 Then DoEvents
                If (tmWegenerVehInfo(llVefIndex).iVefCode = imVefCode) Then
                    llSerialNo = tmWegenerImport(ilImport).sSerialNo1
                    slPort = tmWegenerVehInfo(llVefIndex).sPort
                    ilTest = True
                    Select Case slPort
                        Case "A"
                            If slGroupInfoA <> "" Then
                                ilTest = False
                            End If
                        Case "B"
                            If slGroupInfoB <> "" Then
                                ilTest = False
                            End If
                        Case "C"
                            If slGroupInfoC <> "" Then
                                ilTest = False
                            End If
                        Case "D"
                            If slGroupInfoD <> "" Then
                                ilTest = False
                            End If
                    End Select
                    If ilTest Then
                        ilRet = gRegionTestDefinition(ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, tmMergeRegionDefinition(), tmMergeSplitCategoryInfo(), llRegion, slGroupInfo)
                        If ilRet Then
                            Select Case slPort
                                Case "A"
                                    slGroupInfoA = slGroupInfo
                                Case "B"
                                    slGroupInfoB = slGroupInfo
                                Case "C"
                                    slGroupInfoC = slGroupInfo
                                Case "D"
                                    slGroupInfoD = slGroupInfo
                            End Select
                            'Replace Call Letters if required
                            If ((slGroupInfoA <> "") Or (imPortDefined(0) = False)) And ((slGroupInfoB <> "") Or (imPortDefined(1) = False)) And ((slGroupInfoC <> "") Or (imPortDefined(2) = False)) And ((slGroupInfoD <> "") Or (imPortDefined(3) = False)) Then
                                mAnyStations = True
                                Exit Function
                            End If
                        End If
                    End If
'                    Exit For
                    Exit Do
                End If
                llVefIndex = tmWegenerVehInfo(llVefIndex).lVefCodeNext
            Loop
        'End If
        End If
    'Next ilImport
    Next ilIndex
    If (slGroupInfoA = "") And (slGroupInfoB = "") And (slGroupInfoC = "") And (slGroupInfoD = "") Then
        mAnyStations = False
    Else
        mAnyStations = True
    End If
    Exit Function
ErrHandle:
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadStationReceiverRecords     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File to get wegener       *
'*                      stations                       *
'*                                                     *
'*******************************************************
Private Function mReadStationReceiverRecords() As Integer
    Dim ilEof As Integer
    Dim slPath As String
    Dim slLine As String
    Dim slChar As String
    Dim slWord As String
    Dim ilRet As Integer
    Dim slCallLettersA As String
    Dim slCallLettersB As String
    Dim slCallLettersC As String
    Dim slSerialNo As String
    Dim slFromFile As String
    Dim ilImport As Integer
    Dim ilPosStart As Integer
    Dim ilPosEnd As Integer
    Dim ilVehPosStart As Integer
    Dim ilTRNPosStart As Integer
    Dim slGroup As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slMainGroup As String
    Dim slPort As String
    Dim ilFound As Integer
    Dim slTrueCallLetters As String
    Dim slCallLetters As String
    Dim il4600RX As Integer
    Dim ilPortFound As Integer
    Dim ilPass As Integer
    Dim ilStationFound As Integer
    Dim ilPort As Integer
    Dim blRx_Calls As Boolean
    Dim llVefIndex As Long
    Dim blFound As Boolean
    Dim llSerialNo As Long
    
    If igExportSource = 2 Then DoEvents
    mReadStationReceiverRecords = 0
    ReDim tmWegenerImport(0 To 0) As WEGENERIMPORT
    ReDim tmWegenerVehInfo(0 To 0) As WEGENERVEHINFO
    ReDim tmWegenerFormatSort(0 To 0) As WEGENERFORMATSORT
    ReDim tmWegenerTimeZoneSort(0 To 0) As WEGENERTIMEZONESORT
    ReDim tmWegenerMarketSort(0 To 0) As WEGENERMARKETSORT
    ReDim tmWegenerMSAMarketSort(0 To 0) As WEGENERMARKETSORT
    ReDim tmWegenerPostalSort(0 To 0) As WEGENERPOSTALSORT
    ReDim tmWegenerIndex(0 To 0) As WEGENERINDEX
    On Error GoTo mReadStationReceiverRecordsErr:
    slPath = udcCriteria.edcWStationInfo
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
    edcMsg.Text = "Reading Station Info from rx_calls.Csv...."
    'ilRet = 0
    slFromFile = slPath & "rx_calls.Csv"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet = 0 Then
        blRx_Calls = True
        Do While Not EOF(hmFrom)
            If igExportSource = 2 Then DoEvents
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                If igExportSource = 2 Then DoEvents
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slSerialNo = Trim$(smFields(1))
                    slSerialNo = Trim$(smFields(0))
                    If slSerialNo <> "SN" Then
                        'For ilPort = 2 To 5 Step 1
                        For ilPort = 1 To 4 Step 1
                            slCallLetters = Trim$(smFields(ilPort))
                            If slCallLetters <> "" Then
                                slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLetters)
                                ilRet = gBinarySearchStation(slTrueCallLetters)
                                If ilRet = -1 Then
                                    '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                                    If (InStr(1, UCase(slCallLetters), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-HD", vbBinaryCompare) > 0) Then
                                        mAddMsgToList slCallLetters & " not defined"
                                        mReadStationReceiverRecords = 3
                                    End If
                                Else
                                    tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLetters
                                    tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                                    tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                                    'tmWegenerImport(UBound(tmWegenerImport)).sPort = Chr(Asc("A") + ilPort - 2)
                                    tmWegenerImport(UBound(tmWegenerImport)).sPort = Chr(Asc("A") + ilPort - 1)
                                    tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                                    tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                                    ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                                End If
                            End If
                        Next ilPort
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    Else
        blRx_Calls = False
        edcMsg.Text = "Reading Station Info from JNS_RecSerialNum.Csv...."
        'ilRet = 0
        slFromFile = slPath & "JNS_RecSerialNum.Csv"
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
            mReadStationReceiverRecords = 1
            Exit Function
        End If
        Do While Not EOF(hmFrom)
            If igExportSource = 2 Then DoEvents
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                If igExportSource = 2 Then DoEvents
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slCallLettersA = Trim$(smFields(1))
                    slCallLettersA = Trim$(smFields(0))
                    'slSerialNo = smFields(2)
                    slSerialNo = smFields(1)
                    If (Left$(slSerialNo, 1) >= "A") And (Left$(slSerialNo, 1) <= "Z") Then
                        slSerialNo = Val(Mid$(slSerialNo, 2))
                    End If
                    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersA)
                    ilRet = gBinarySearchStation(slTrueCallLetters)
                    If ilRet = -1 Then
                        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                        If (InStr(1, UCase(slCallLettersA), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-HD", vbBinaryCompare) > 0) Then
                            mAddMsgToList slCallLettersA & " not defined"
                            mReadStationReceiverRecords = 3
                        End If
                    Else
                        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersA
                        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                        tmWegenerImport(UBound(tmWegenerImport)).sPort = "A"
                        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
        
        'ilRet = 0
        On Error GoTo mReadStationReceiverRecordsErr:
        edcMsg.Text = "Reading Station Info from PortB-C.Csv...."
        If igExportSource = 2 Then DoEvents
        'hmFrom = FreeFile
        slFromFile = slPath & "Port B-C.Csv"
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            'ilRet = 0
            slFromFile = slPath & "PortB-C.Csv"
            'Open slFromFile For Input Access Read As hmFrom
            ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
            If ilRet <> 0 Then
                'ilRet = 0
                slFromFile = slPath & "Port_B-C.Csv"
                'Open slFromFile For Input Access Read As hmFrom
                ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                If ilRet <> 0 Then
                    'ilRet = 0
                    slFromFile = slPath & "Port B-D.Csv"
                    'Open slFromFile For Input Access Read As hmFrom
                    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                    If ilRet <> 0 Then
                        'ilRet = 0
                        slFromFile = slPath & "PortB-D.Csv"
                        'Open slFromFile For Input Access Read As hmFrom
                        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                        If ilRet <> 0 Then
                            'ilRet = 0
                            slFromFile = slPath & "Port_B-D.Csv"
                            'Open slFromFile For Input Access Read As hmFrom
                            ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                            If ilRet <> 0 Then
                                mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
                                mReadStationReceiverRecords = 1
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Do While Not EOF(hmFrom)
            If igExportSource = 2 Then DoEvents
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                If igExportSource = 2 Then DoEvents
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slSerialNo = Trim$(smFields(1))
                    slSerialNo = Trim$(smFields(0))
                    If (Left$(slSerialNo, 1) >= "A") And (Left$(slSerialNo, 1) <= "Z") Then
                        slSerialNo = Val(Mid$(slSerialNo, 2))
                    End If
                    'slCallLettersB = Trim$(smFields(3))
                    'slCallLettersC = Trim$(smFields(5))
                    'If slCallLettersB <> "" Then
                    '    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersB)
                    '    ilRet = gBinarySearchStation(slTrueCallLetters)
                    '    If ilRet = -1 Then
                    '        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                    '        If (InStr(1, UCase(slCallLettersB), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-HD", vbBinaryCompare) > 0) Then
                    '            mAddMsgToList slCallLettersB & " not defined"
                    '            mReadStationReceiverRecords = 3
                    '        End If
                    '    Else
                    '        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersB
                    '        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).icode
                    '        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPort = "B"
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                    '        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                    '        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    '    End If
                    'End If
                    'If slCallLettersC <> "" Then
                    '    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersC)
                    '    ilRet = gBinarySearchStation(slTrueCallLetters)
                    '    If ilRet = -1 Then
                    '        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                    '        If (InStr(1, UCase(slCallLettersC), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-HD", vbBinaryCompare) > 0) Then
                    '            mAddMsgToList slCallLettersC & " not defined"
                    '            mReadStationReceiverRecords = 3
                    '        End If
                    '    Else
                    '        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersC
                    '        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).icode
                    '        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPort = "C"
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                    '        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                    '        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    '    End If
                    'End If
                    'For ilPort = 3 To 7 Step 2
                    For ilPort = 2 To 6 Step 2
                        slCallLetters = Trim$(smFields(ilPort))
                        If slCallLetters <> "" Then
                            slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLetters)
                           ilRet = gBinarySearchStation(slTrueCallLetters)
                            If ilRet = -1 Then
                                '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                                If (InStr(1, UCase(slCallLetters), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-HD", vbBinaryCompare) > 0) Then
                                    mAddMsgToList slCallLetters & " not defined"
                                    mReadStationReceiverRecords = 3
                                End If
                            Else
                                tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLetters
                                tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                                tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                                If ilPort = 3 Then
                                    tmWegenerImport(UBound(tmWegenerImport)).sPort = "B"
                                ElseIf ilPort = 5 Then
                                    tmWegenerImport(UBound(tmWegenerImport)).sPort = "C"
                                Else
                                    tmWegenerImport(UBound(tmWegenerImport)).sPort = "D"
                                End If
                                tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                                tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                                ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                            End If
                        End If
                    Next ilPort
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    End If
    
    'ilRet = 0
    On Error GoTo mReadStationReceiverRecordsErr:
    edcMsg.Text = "Reading Station Info from JNS_RecGroup.Csv...."
    If igExportSource = 2 Then DoEvents
    slFromFile = slPath & "JNS_RecGroup.Csv"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        mReadStationReceiverRecords = 1
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        If igExportSource = 2 Then DoEvents
        ilRet = 0
        On Error GoTo 0
        slLine = ""
        '2/10/12: Ignore fields that contain unless information
        'Do While Not EOF(hmFrom)
        '    slChar = Input(1, #hmFrom)
        '    If slChar = sgLF Then
        '        Exit Do
        '    ElseIf slChar <> sgCR Then
        '        slLine = slLine & slChar
        '    End If
        'Loop
        slWord = ""
        Do While Not EOF(hmFrom)
            If igExportSource = 2 Then DoEvents
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                If slLine = "" Then
                    slLine = slWord
                Else
                    If InStr(1, UCase(slWord), "FMT_", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), "ST_", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), "DMA_", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), "MSA_", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), "TZ_", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), smVehicleGroupPrefix, vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    ElseIf InStr(1, UCase(slWord), "4600RX", vbBinaryCompare) = 1 Then
                        slLine = slLine & slWord
                    End If
                End If
                Exit Do
            ElseIf slChar <> sgCR Then
                'slLine = slLine & slChar
                If slChar = "," Then
                    If slLine = "" Then
                        slLine = slWord & slChar
                    Else
                        If InStr(1, UCase(slWord), "FMT_", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), "ST_", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), "DMA_", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), "MSA_", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), "TZ_", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), smVehicleGroupPrefix, vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        ElseIf InStr(1, UCase(slWord), "4600RX", vbBinaryCompare) = 1 Then
                            slLine = slLine & slWord & slChar
                        End If
                    End If
                    slWord = ""
                Else
                    slWord = slWord & slChar
                End If
            End If
        Loop
        On Error GoTo mReadStationReceiverRecordsErr
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        If igExportSource = 2 Then DoEvents
        If imTerminate Then
            mAddMsgToList "User Cancelled Export"
            mReadStationReceiverRecords = 2
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If igExportSource = 2 Then DoEvents
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                'gParseCDFields slLine, True, smFields()
                ilRet = gParseItem(Left(slLine, 100), 1, ",", slCallLetters)
                il4600RX = InStr(1, slLine, "4600RX", vbTextCompare)
                If (slCallLetters <> "") And (il4600RX = 0) Then
                    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLetters)
                    ilRet = gBinarySearchStation(slTrueCallLetters)
                    If ilRet = -1 Then
                        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                        If (InStr(1, UCase(slCallLetters), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-HD", vbBinaryCompare) > 0) Then
                            mAddMsgToList slCallLetters & " not defined"
                            mReadStationReceiverRecords = 3
                        End If
                    Else
                        ilStationFound = False
                        slCallLettersA = slCallLetters
                        For ilPass = 0 To 3 Step 1
                            slCallLetters = slCallLettersA
                            If ilPass = 1 Then
                                slPort = "B"
                            ElseIf ilPass = 2 Then
                                slPort = "C"
                            ElseIf ilPass = 3 Then
                                slPort = "D"
                            Else
                                ilPortFound = True
                                llSerialNo = -1
                                slPort = "A"
                                For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                                    If (tmWegenerImport(ilImport).sPort = slPort) And (slCallLetters = Trim$(tmWegenerImport(ilImport).sCallLetters)) Then
                                        llSerialNo = Val(tmWegenerImport(ilImport).sSerialNo1)
                                        Exit For
                                    End If
                                Next ilImport
                            End If
                            If (ilPass = 1) Or (ilPass = 2) Or (ilPass = 3) Then
                                ilPortFound = False
                                For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                                    If igExportSource = 2 Then DoEvents
                                    If (tmWegenerImport(ilImport).sPort = "A") And (StrComp(Trim$(tmWegenerImport(ilImport).sCallLetters), slCallLettersA, vbTextCompare) = 0) And (llSerialNo = Val(tmWegenerImport(ilImport).sSerialNo1)) Then
                                        For ilLoop = 0 To UBound(tmWegenerImport) - 1 Step 1
                                            If igExportSource = 2 Then DoEvents
                                            If tmWegenerImport(ilLoop).sPort = slPort Then
                                                '3/7/19: stripped SN from serial number when stored into array
                                                'If Val(tmWegenerImport(ilLoop).sSerialNo1) = Val(Mid(tmWegenerImport(ilImport).sSerialNo1, 2)) Then
                                                If Val(tmWegenerImport(ilLoop).sSerialNo1) = llSerialNo Then
                                                    slCallLetters = Trim$(tmWegenerImport(ilLoop).sCallLetters)
                                                    ilPortFound = True
                                                    Exit For
                                                End If
                                            End If
                                        Next ilLoop
                                        Exit For
                                    End If
                                Next ilImport
                            End If
                            If ilPortFound Then
                                For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                                    If igExportSource = 2 Then DoEvents
                                    'If tmWegenerImport(ilImport).iShttCode = tgStationInfo(ilRet).iCode Then
                                    If (StrComp(Trim$(tmWegenerImport(ilImport).sCallLetters), Trim$(slCallLetters), vbTextCompare) = 0) And (Val(tmWegenerImport(ilImport).sSerialNo1) = llSerialNo) Then
                                        tmWegenerImport(ilImport).iRecGroupFd = True
                                        ilStationFound = True
                                        ilFound = 0
                                        ilPosStart = InStr(1, slLine, "DMA_", vbTextCompare)
                                        Do While ilPosStart > 0
                                            If igExportSource = 2 Then DoEvents
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgMarketInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        tmWegenerImport(ilImport).iMktCode = tgMarketInfo(ilLoop).lCode
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                Exit Do
                                            End If
                                            ilPosStart = InStr(ilPosEnd + 1, slLine, "DMA_", vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " DMA_ Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        ElseIf ilFound = 2 Then
                                            mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                        End If
                                        
                                        ilFound = 0
                                        ilPosStart = InStr(1, slLine, "MSA_", vbTextCompare)
                                        Do While ilPosStart > 0
                                            If igExportSource = 2 Then DoEvents
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgMSAMarketInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        tmWegenerImport(ilImport).iMSAMktCode = tgMSAMarketInfo(ilLoop).lCode
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                Exit Do
                                            End If
                                            ilPosStart = InStr(ilPosEnd + 1, slLine, "MSA_", vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " MSA_ Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        ElseIf ilFound = 2 Then
                                            mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                        End If
                                        ilFound = 0
                                        ilPosStart = InStr(1, slLine, "FMT_", vbTextCompare)
                                        Do While ilPosStart > 0
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgFormatInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        tmWegenerImport(ilImport).iFormatCode = tgFormatInfo(ilLoop).lCode
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                Exit Do
                                            End If
                                            ilPosStart = InStr(ilPosEnd + 1, slLine, "FMT_", vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " FMT_ Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        ElseIf ilFound = 2 Then
                                            mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                        End If
                                        ilFound = 0
                                        ilPosStart = InStr(1, slLine, "ST_", vbTextCompare)
                                        Do While ilPosStart > 0
                                            If igExportSource = 2 Then DoEvents
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgStateInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgStateInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        tmWegenerImport(ilImport).sPostalName = tgStateInfo(ilLoop).sPostalName
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                Exit Do
                                            End If
                                            ilPosStart = InStr(ilPosEnd + 1, slLine, "ST_", vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " ST_ Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        ElseIf ilFound = 2 Then
                                            mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                        End If
                                        ilFound = 0
                                        ilPosStart = InStr(1, slLine, "TZ_", vbTextCompare)
                                        Do While ilPosStart > 0
                                            If igExportSource = 2 Then DoEvents
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgTimeZoneInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        tmWegenerImport(ilImport).iTztCode = tgTimeZoneInfo(ilLoop).iCode
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                Exit Do
                                            End If
                                            ilPosStart = InStr(ilPosEnd + 1, slLine, "TZ_", vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " TZ_ Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        ElseIf ilFound = 2 Then
                                            mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                        End If
                                        ilFound = 0
                                        '6/17/11
                                        ''10/15/10:  Test for either veh_ or trn_
                                        ''ilPosStart = InStr(1, slLine, "VEH_", vbTextCompare)
                                        'ilVehPosStart = InStr(1, slLine, "VEH_", vbTextCompare)
                                        'ilTRNPosStart = InStr(1, slLine, "TRN_", vbTextCompare)
                                        'If ilTRNPosStart <= 0 Then
                                        '    ilPosStart = ilVehPosStart
                                        'ElseIf ilVehPosStart <= 0 Then
                                        '    ilPosStart = ilTRNPosStart
                                        'ElseIf (ilVehPosStart < ilTRNPosStart) Then
                                        '    ilPosStart = ilVehPosStart
                                        'Else
                                        '    ilPosStart = ilTRNPosStart
                                        'End If
                                        ilPosStart = InStr(1, UCase(slLine), smVehicleGroupPrefix, vbTextCompare)
                                        Do While ilPosStart > 0
                                            If igExportSource = 2 Then DoEvents
                                            If ilFound = 0 Then
                                                ilFound = 1
                                            End If
                                            ilPosEnd = InStr(ilPosStart, slLine, ",", vbTextCompare)
                                            If ilPosEnd = 0 Then
                                                ilPosEnd = Len(slLine) + 1
                                            End If
                                            slGroup = Mid$(slLine, ilPosStart, ilPosEnd - ilPosStart)
                                            If mFindWegenerIndex(slGroup, ilImport, slMainGroup, slPort) Then
                                                ilFound = 2
                                                'Loop for match
                                                For ilLoop = 0 To UBound(tgVffInfo) - 1 Step 1
                                                    If igExportSource = 2 Then DoEvents
                                                    If StrComp(Trim$(tgVffInfo(ilLoop).sGroupName), slMainGroup, vbTextCompare) = 0 Then
                                                        ilFound = 3
                                                        If tmWegenerImport(ilImport).lVefCodeFirst = -1 Then
                                                            tmWegenerVehInfo(UBound(tmWegenerVehInfo)).iVefCode = tgVffInfo(ilLoop).iVefCode
                                                            tmWegenerVehInfo(UBound(tmWegenerVehInfo)).sGroup = slMainGroup  'slGroup
                                                            tmWegenerVehInfo(UBound(tmWegenerVehInfo)).sPort = slPort
                                                            tmWegenerVehInfo(UBound(tmWegenerVehInfo)).lVefCodeNext = -1
                                                            tmWegenerImport(ilImport).lVefCodeFirst = UBound(tmWegenerVehInfo)
                                                            ReDim Preserve tmWegenerVehInfo(0 To UBound(tmWegenerVehInfo) + 1) As WEGENERVEHINFO
                                                        Else
                                                            '12/10/15: test for duplicates
                                                            blFound = False
                                                            llVefIndex = tmWegenerImport(ilImport).lVefCodeFirst
                                                            Do While llVefIndex <> -1
                                                                If igExportSource = 2 Then DoEvents
                                                                If (tmWegenerVehInfo(llVefIndex).iVefCode = tgVffInfo(ilLoop).iVefCode) Then
                                                                    blFound = True
                                                                    Exit Do
                                                                End If
                                                                llVefIndex = tmWegenerVehInfo(llVefIndex).lVefCodeNext
                                                            Loop
                                                            If Not blFound Then
                                                                tmWegenerVehInfo(UBound(tmWegenerVehInfo)).iVefCode = tgVffInfo(ilLoop).iVefCode
                                                                tmWegenerVehInfo(UBound(tmWegenerVehInfo)).sGroup = slMainGroup 'slGroup
                                                                tmWegenerVehInfo(UBound(tmWegenerVehInfo)).sPort = slPort
                                                                tmWegenerVehInfo(UBound(tmWegenerVehInfo)).lVefCodeNext = tmWegenerImport(ilImport).lVefCodeFirst
                                                                tmWegenerImport(ilImport).lVefCodeFirst = UBound(tmWegenerVehInfo)
                                                                ReDim Preserve tmWegenerVehInfo(0 To UBound(tmWegenerVehInfo) + 1) As WEGENERVEHINFO
                                                            End If
                                                        End If
                                                        '12/10/15: Remove for to get two or more vehicles with the same vehicle code.
                                                        '          The previous fix merged the vehicles into the same output file but did not handle region copy for both
                                                        'Exit For
                                                    End If
                                                Next ilLoop
                                                If ilFound = 2 Then
                                                    mAddMsgToList slCallLetters & " " & slMainGroup & " Root Group not found"
                                                End If
                                            End If
                                            '6/17/11
                                            ''10/15/10:  Test for either veh_ or trn_
                                            ''ilPosStart = InStr(ilPosEnd + 1, slLine, "Veh_", vbTextCompare)
                                            'ilVehPosStart = InStr(ilPosEnd + 1, slLine, "Veh_", vbTextCompare)
                                            'ilTRNPosStart = InStr(ilPosEnd + 1, slLine, "TRN_", vbTextCompare)
                                            'If ilTRNPosStart <= 0 Then
                                            '    ilPosStart = ilVehPosStart
                                            'ElseIf ilVehPosStart <= 0 Then
                                            '    ilPosStart = ilTRNPosStart
                                            'ElseIf (ilVehPosStart < ilTRNPosStart) Then
                                            '    ilPosStart = ilVehPosStart
                                            'Else
                                            '    ilPosStart = ilTRNPosStart
                                            'End If
                                            ilPosStart = InStr(ilPosEnd + 1, UCase(slLine), smVehicleGroupPrefix, vbTextCompare)
                                        Loop
                                        If ilFound = 1 Then
                                            mAddMsgToList slCallLetters & " Vehicle Group Port " & tmWegenerImport(ilImport).sPort & " not found"
                                        End If
                                        'Can't exit as station can be defined with a A and B Port
                                        'Exit For
                                    End If
                                Next ilImport
                            End If
                        Next ilPass
                        If Not ilStationFound Then
                            If blRx_Calls Then
                                mAddMsgToList slTrueCallLetters & " defined in JNS_RecGroup but not defined in RX_Calls"
                            Else
                                mAddMsgToList slTrueCallLetters & " defined in JNS_RecGroup but not defined in JNS_RecSerialNUM or Port B-C"
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If Not tmWegenerImport(ilImport).iRecGroupFd Then
            If blRx_Calls Then
                If Trim$(tmWegenerImport(ilImport).sPort) = "A" Then
                    mAddMsgToList Trim$(tmWegenerImport(ilImport).sCallLetters) & " on Port " & tmWegenerImport(ilImport).sPort & " defined in RX_Calls but not defined in JNS_RecGroup"
                Else
                    mAddMsgToList Trim$(tmWegenerImport(ilImport).sCallLetters) & " on Port " & tmWegenerImport(ilImport).sPort & " defined in Rx_Calls but not defined in JNS_RecGroup"
                End If
            Else
                If Trim$(tmWegenerImport(ilImport).sPort) = "A" Then
                    mAddMsgToList Trim$(tmWegenerImport(ilImport).sCallLetters) & " on Port " & tmWegenerImport(ilImport).sPort & " defined in JNS_RecSerialNUM but not defined in JNS_RecGroup"
                Else
                    mAddMsgToList Trim$(tmWegenerImport(ilImport).sCallLetters) & " on Port " & tmWegenerImport(ilImport).sPort & " defined in Port B-C but not defined in JNS_RecGroup"
                End If
            End If
        End If
    Next ilImport
    mSetWegenerSort
    
    Exit Function
mReadStationReceiverRecordsErr:
    ilRet = Err.Number
    Resume Next
End Function


Private Sub mSeparateRegions(tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO)
    'If a region is defined as:
    '(Fmt1 or Fmt2 or Fmt3) and (St1 or St2) and (Not K1111 and Not K222)
    'Convert to:
    'Region 1: Fmt1 and St1 And Not K111 and Not K222
    'Region 2: Fmt1 and St2 And Not K111 and Not K222
    'Region 3: Fmt2 and St1 And Not K111 and Not K222
    'Region 4: Fmt2 and St2 And Not K111 and Not K222
    'Region 5: Fmt3 and St1 And Not K111 and Not K222
    'Region 6: Fmt3 and St2 And Not K111 and Not K222
    Dim llFormatIndex As Long
    Dim llRegion As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    
    On Error GoTo ErrHandle
    
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            
        If tlRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            Do
                If igExportSource = 2 Then DoEvents
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = UBound(tmSplitCategoryInfo)
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llFormatIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                If tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
                    llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
                    Do
                        If igExportSource = 2 Then DoEvents
                        tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = UBound(tmSplitCategoryInfo)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llOtherIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                        If llExcludeIndex <> -1 Then
                            tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                            Do While llExcludeIndex <> -1
                                If igExportSource = 2 Then DoEvents
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                            Loop
                        End If
                        ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                        llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
                        If llOtherIndex <> -1 Then
                            tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                            tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = UBound(tmSplitCategoryInfo)
                            tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                            tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llFormatIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        End If
                    Loop While llOtherIndex <> -1
                Else
                    llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                    If llExcludeIndex <> -1 Then
                        tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Do While llExcludeIndex <> -1
                            If igExportSource = 2 Then DoEvents
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                    End If
                End If
                llFormatIndex = tlSplitCategoryInfo(llFormatIndex).lNext
            Loop While llFormatIndex <> -1
        ElseIf tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
            llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
            Do
                If igExportSource = 2 Then DoEvents
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = UBound(tmSplitCategoryInfo)
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llOtherIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                If llExcludeIndex <> -1 Then
                    tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                    ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                    Do While llExcludeIndex <> -1
                        If igExportSource = 2 Then DoEvents
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                    Loop
                End If
                ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
        Else
            'Exclude only
            llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
            If llExcludeIndex <> -1 Then
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Do While llExcludeIndex <> -1
                    If igExportSource = 2 Then DoEvents
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                    ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
            End If
        End If
    Next llRegion
    Exit Sub
ErrHandle:
    Resume Next

End Sub

Private Function mFormRegionAddress(slVehGroup As String, slInGroupInfo As String, slGroupLetter As String, llSerialNo As Long) As String
    'Translate the User defined region into region definition that Wegener understands
    'User enters:
    'Urban and California
    'Wegener wants
    'Fmt_123 ^ St_CA
    'Wegener names are call Group Names and are defined as menu items for each category
    '(Format, State, Time zone and Market.  For stations, the call letters are used)
    'Symbols:  ^ = And; ~ = Not
    Dim ilPos As Integer
    Dim slGroupInfo As String
    Dim slStr As String
    Dim slInclExcl As String
    Dim slCategory As String
    Dim slvalue As String
    Dim ilValue As Integer
    Dim ilSnt As Integer
    Dim slAddress As String
    Dim ilRet As Integer
    Dim slGroupName As String
    Dim ilShtt As Integer
    Dim ilFind As Integer
    Dim slChar As String
    Dim llSerialNo1 As Long
    Dim llTestSerialNo1 As Long
    Dim ilImport As Integer
    Dim ilLoop As Integer
    
    On Error GoTo mFormRegionAddressErr:
    slAddress = ""
    slGroupInfo = slInGroupInfo
    ilPos = 1
    Do
        If igExportSource = 2 Then DoEvents
        ilPos = InStr(1, slGroupInfo, "|", vbTextCompare)
        If ilPos = 0 Then
            If Len(slGroupInfo) = 0 Then
                Exit Do
            Else
                ilPos = Len(slGroupInfo) + 1
            End If
        End If
        slStr = Left(slGroupInfo, ilPos - 1)
        slGroupInfo = Mid$(slGroupInfo, ilPos + 1)
        slInclExcl = Left$(slStr, 1)
        slCategory = Mid$(slStr, 2, 1)
        slvalue = Trim$(Mid$(slStr, 3))
        If slCategory <> "N" Then
            ilValue = Val(slvalue)
        End If
        Select Case slCategory
            Case "M"    'DMA Market
                ilRet = gBinarySearchMkt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgMarketInfo(ilRet).sGroupName)
                End If
            Case "A"    'MSA Market
                ilRet = gBinarySearchMSAMkt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgMSAMarketInfo(ilRet).sGroupName)
                End If
            Case "N"    'State Name
                For ilSnt = 0 To UBound(tgStateInfo) - 1 Step 1
                    If StrComp(Trim$(tgStateInfo(ilSnt).sPostalName), slvalue, vbTextCompare) = 0 Then
                        slGroupName = Trim$(tgStateInfo(ilSnt).sGroupName)
                        Exit For
                    End If
                Next ilSnt
            Case "F"    'Format
                ilRet = gBinarySearchFmt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgFormatInfo(ilRet).sGroupName)
                End If
            Case "T"    'Time zone
                ilRet = gBinarySearchTzt(ilValue)
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgTimeZoneInfo(ilRet).sGroupName)
                End If
            Case "S"    'Station
                'ilShtt = gBinarySearchStationInfoByCode(ilValue)
                'If ilShtt <> -1 Then
                '    slGroupName = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                '    If tgStationInfoByCode(ilShtt).sPort <> "A" Then
                '        slChar = Left$(tgStationInfoByCode(ilShtt).sSerialNo1, 1)
                '        If (slChar >= "A") And (slChar <= "Z") Then
                '            llSerialNo1 = Val(Mid$(tgStationInfoByCode(ilShtt).sSerialNo1, 2))
                '        Else
                '            llSerialNo1 = Val(tgStationInfoByCode(ilShtt).sSerialNo1)
                '        End If
                '        'Find match Port A station
                '        For ilFind = 0 To UBound(tgStationInfoByCode) - 1 Step 1
                '            If (tgStationInfoByCode(ilFind).sUsedForWegener = "Y") And (tgStationInfoByCode(ilFind).sPort = "A") Then
                '                slChar = Left$(tgStationInfoByCode(ilFind).sSerialNo1, 1)
                '                If (slChar >= "A") And (slChar <= "Z") Then
                '                    llTestSerialNo1 = Val(Mid$(tgStationInfoByCode(ilFind).sSerialNo1, 2))
                '                Else
                '                    llTestSerialNo1 = Val(tgStationInfoByCode(ilFind).sSerialNo1)
                '                End If
                '                If llTestSerialNo1 = llSerialNo1 Then
                '                    slGroupName = Trim$(tgStationInfoByCode(ilFind).sCallLetters)
                '                    Exit For
                '                End If
                '            End If
                '        Next ilFind
                '    End If
                'End If
                For ilLoop = 0 To UBound(tmWegenerImport) - 1 Step 1
                    If tmWegenerImport(ilLoop).iShttCode = ilValue And (Val(tmWegenerImport(ilLoop).sSerialNo1) = llSerialNo) Then
                        slGroupName = Trim$(tmWegenerImport(ilLoop).sCallLetters)
                        If tmWegenerImport(ilLoop).sPort <> "A" Then
                            For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                                If (ilLoop <> ilImport) And (tmWegenerImport(ilImport).sPort = "A") Then
                                    '3/7/19: stripped SN from serial number when stored into array
                                    'If Val(tmWegenerImport(ilLoop).sSerialNo1) = Val(Mid(tmWegenerImport(ilImport).sSerialNo1, 2)) Then
                                    If Val(tmWegenerImport(ilLoop).sSerialNo1) = Val(tmWegenerImport(ilImport).sSerialNo1) Then
                                        slGroupName = Trim$(tmWegenerImport(ilImport).sCallLetters)
                                        Exit For
                                    End If
                                End If
                            Next ilImport
                        End If
                        Exit For
                    End If
                Next ilLoop
        End Select
        If slInclExcl = "E" Then
            slGroupName = "~" & slGroupName
        End If
        If slCategory <> "S" Then
            If slAddress = "" Then
                slAddress = slGroupName & "_" & slGroupLetter
            Else
                slAddress = slAddress & "^" & slGroupName & "_" & slGroupLetter
            End If
        Else
            If slAddress = "" Then
                slAddress = slGroupName
            Else
                slAddress = slAddress & "^" & slGroupName
            End If
        End If
    Loop While Len(slGroupInfo) > 0
    slAddress = slVehGroup & "_" & slGroupLetter & "^" & slAddress
    'ilPos = 1
    'Do
    '    ilPos = InStr(ilPos, slAddress, "^", vbTextCompare)
    '    If ilPos > 0 Then
    '        slAddress = Left(slAddress, ilPos - 1) & "_" & slGroupLetter & Mid(slAddress, ilPos)
    '    Else
    '        slAddress = slAddress & "_" & slGroupLetter
    '        Exit Do
    '    End If
    '    ilPos = ilPos + 3
    'Loop
    
    mFormRegionAddress = mCreateGroupNames(slAddress)
    Exit Function
mFormRegionAddressErr:
    Resume Next
End Function

Private Function mMergeCategory(llRegionIndex As Long, ilRepeatEndRow As Integer) As Integer
    'Combine Region category definition togather.
    'This is used to form region from each spot with a break (The intersection of regions acrodd spots)
    'The result is used to determine if any station will receive the region copy defined by the merge.
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim llMergeOtherIndex As Long
    Dim llLastMergeOtherIndex As Long
    Dim llMergeExcludeIndex As Long
    Dim llLastMergeExcludeIndex As Long
    Dim ilImport As Integer
    Dim ilAllowMerge As Integer
    Dim ilShttCode As Integer
    '3/3/18
    Dim llIndex As Long
    Dim llNext As Long
    Dim llStnNext As Long
    
    On Error GoTo ErrHandle
    
    If (tmRegionDefinition(llRegionIndex).lFormatFirst = -1) And (tmRegionDefinition(llRegionIndex).lOtherFirst = -1) And (tmRegionDefinition(llRegionIndex).lExcludeFirst = -1) Then
        mMergeCategory = False
        Exit Function
    End If
    If UBound(tmMergeRegionDefinition) = 0 Then
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lFormatFirst = -1
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lOtherFirst = -1
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lExcludeFirst = -1
        ReDim Preserve tmMergeRegionDefinition(0 To UBound(tmMergeRegionDefinition) + 1) As REGIONDEFINITION
    End If

    If tmRegionDefinition(llRegionIndex).lOtherFirst <> -1 Then
        llOtherIndex = tmRegionDefinition(llRegionIndex).lOtherFirst
        If tmSplitCategoryInfo(llOtherIndex).sCategory = "S" Then
            ilAllowMerge = False
            For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                ilShttCode = tmWegenerImport(ilImport).iShttCode
                If tmSplitCategoryInfo(llOtherIndex).iIntCode = ilShttCode Then
                    ilAllowMerge = True
                    Exit For
                End If
            Next ilImport
            If ilAllowMerge = False Then
                mMergeCategory = False
                Exit Function
            End If
        End If
    End If
    
    
    If tmRegionDefinition(llRegionIndex).lExcludeFirst <> -1 Then
        llExcludeIndex = tmRegionDefinition(llRegionIndex).lExcludeFirst
        Do While llExcludeIndex <> -1
            If igExportSource = 2 Then DoEvents
            If tmSplitCategoryInfo(llExcludeIndex).sCategory = "S" Then
                ilAllowMerge = False
                For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    ilShttCode = tmWegenerImport(ilImport).iShttCode
                    If tmSplitCategoryInfo(llExcludeIndex).iIntCode = ilShttCode Then
                        ilAllowMerge = True
                        Exit For
                    End If
                Next ilImport
                If ilAllowMerge = False Then
                    mMergeCategory = False
                    Exit Function
                End If
            End If
            llExcludeIndex = tmSplitCategoryInfo(llExcludeIndex).lNext
        Loop
    End If
    
    If tmRegionDefinition(llRegionIndex).lFormatFirst <> -1 Then
        If tmMergeRegionDefinition(0).lFormatFirst <> -1 Then
            mMergeCategory = False
            Exit Function
        End If
        llFormatIndex = tmRegionDefinition(llRegionIndex).lFormatFirst
        tmMergeRegionDefinition(0).lFormatFirst = UBound(tmMergeSplitCategoryInfo)
        tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llFormatIndex)
        ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
    End If
    If tmRegionDefinition(llRegionIndex).lOtherFirst <> -1 Then
        llOtherIndex = tmRegionDefinition(llRegionIndex).lOtherFirst
        '3/3/18
        llIndex = -1
        'If tmSplitCategoryInfo(llOtherIndex).sCategory = "S" Then
        '    ilAllowMerge = False
        '    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
        '    'If tmWegenerImport(ilImport).iShttCode = ilShttCode Then
        '        ilShttCode = tmWegenerImport(ilImport).iShttCode
        '        If tmSplitCategoryInfo(llOtherIndex).iIntCode = ilShttCode Then
        '            ilAllowMerge = True
        '            Exit For
        '        End If
        '    Next ilImport
        'Else
        '    ilAllowMerge = True
        'End If
        'If ilAllowMerge Then
            If tmMergeRegionDefinition(0).lOtherFirst <> -1 Then
                llMergeOtherIndex = tmMergeRegionDefinition(0).lOtherFirst
                Do While llMergeOtherIndex <> -1
                    If igExportSource = 2 Then DoEvents
                    If tmSplitCategoryInfo(llOtherIndex).sCategory = tmMergeSplitCategoryInfo(llMergeOtherIndex).sCategory Then
                        'Handle the case where two spots have matching values like same DMA value
                        'mMergeCategory = False
                        'Exit Function
                        If tmSplitCategoryInfo(llOtherIndex).sCategory = "M" Then
                            If tmSplitCategoryInfo(llOtherIndex).iIntCode <> tmMergeSplitCategoryInfo(llMergeOtherIndex).iIntCode Then
                                mMergeCategory = False
                                Exit Function
                            End If
                        ElseIf tmSplitCategoryInfo(llOtherIndex).sCategory = "A" Then
                            If tmSplitCategoryInfo(llOtherIndex).iIntCode <> tmMergeSplitCategoryInfo(llMergeOtherIndex).iIntCode Then
                                mMergeCategory = False
                                Exit Function
                            End If
                        ElseIf tmSplitCategoryInfo(llOtherIndex).sCategory = "N" Then
                            If tmSplitCategoryInfo(llOtherIndex).sName <> tmMergeSplitCategoryInfo(llMergeOtherIndex).sName Then
                                mMergeCategory = False
                                Exit Function
                            End If
                        ElseIf tmSplitCategoryInfo(llOtherIndex).sCategory = "T" Then
                            If tmSplitCategoryInfo(llOtherIndex).iIntCode <> tmMergeSplitCategoryInfo(llMergeOtherIndex).iIntCode Then
                                mMergeCategory = False
                                Exit Function
                            End If
                        ElseIf tmSplitCategoryInfo(llOtherIndex).sCategory = "F" Then
                            If tmSplitCategoryInfo(llOtherIndex).iIntCode <> tmMergeSplitCategoryInfo(llMergeOtherIndex).iIntCode Then
                                mMergeCategory = False
                                Exit Function
                            End If
                        '3/3/18
                        ElseIf tmSplitCategoryInfo(llOtherIndex).sCategory = "S" Then
                            If (ilRepeatEndRow <> -1) Then
                                If tmRegionDefinition(llRegionIndex).iStationCount > 0 Then
                                    llStnNext = tmRegionDefinition(llRegionIndex).lStationOtherFirst
                                    Do
                                        If tmStationSplitCategoryInfo(llStnNext).iIntCode = tmMergeSplitCategoryInfo(llMergeOtherIndex).iIntCode Then
                                            llIndex = llStnNext
                                            Exit Do
                                        End If
                                        llStnNext = tmStationSplitCategoryInfo(llStnNext).lNext
                                    Loop While llStnNext <> -1
                                '****************** Not sure
                                Else
                                    mMergeCategory = False
                                    Exit Function
                                End If
                            End If
                            
                        End If
                    End If
                    llLastMergeOtherIndex = llMergeOtherIndex
                    llMergeOtherIndex = tmMergeSplitCategoryInfo(llMergeOtherIndex).lNext
                Loop
                tmMergeSplitCategoryInfo(llLastMergeOtherIndex).lNext = UBound(tmMergeSplitCategoryInfo)
            Else
                tmMergeRegionDefinition(0).lOtherFirst = UBound(tmMergeSplitCategoryInfo)
            End If
            '3/3/18
            'tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llOtherIndex)
            If llIndex = -1 Then
                tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llOtherIndex)
            Else
                tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmStationSplitCategoryInfo(llIndex)
            End If
            tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)).lNext = -1
            ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
        'End If
    End If
    If tmRegionDefinition(llRegionIndex).lExcludeFirst <> -1 Then
        llExcludeIndex = tmRegionDefinition(llRegionIndex).lExcludeFirst
        Do While llExcludeIndex <> -1
            If igExportSource = 2 Then DoEvents
            '3-12-09 if station, verify as wegener station
            'If tmSplitCategoryInfo(llExcludeIndex).sCategory = "S" Then
            '    ilAllowMerge = False
            '    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
            '    'If tmWegenerImport(ilImport).iShttCode = ilShttCode Then
            '        ilShttCode = tmWegenerImport(ilImport).iShttCode
            '        If tmSplitCategoryInfo(llExcludeIndex).iIntCode = ilShttCode Then
            '            ilAllowMerge = True
            '            Exit For
            '        End If
            '    Next ilImport
            'Else
            '    ilAllowMerge = True
            'End If
            'If ilAllowMerge Then
                If tmMergeRegionDefinition(0).lExcludeFirst <> -1 Then
                    llMergeExcludeIndex = tmMergeRegionDefinition(0).lExcludeFirst
                    Do While llMergeExcludeIndex <> -1
                        llLastMergeExcludeIndex = llMergeExcludeIndex
                        llMergeExcludeIndex = tmMergeSplitCategoryInfo(llMergeExcludeIndex).lNext
                    Loop
                    tmMergeSplitCategoryInfo(llLastMergeExcludeIndex).lNext = UBound(tmMergeSplitCategoryInfo)
                Else
                    tmMergeRegionDefinition(0).lExcludeFirst = UBound(tmMergeSplitCategoryInfo)
                End If
                tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llExcludeIndex)
                tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
            'End If
            llExcludeIndex = tmSplitCategoryInfo(llExcludeIndex).lNext
        Loop
    End If
    mMergeCategory = True
    Exit Function
    Exit Function
ErrHandle:
    Resume Next
End Function

Private Sub mAddMsgToList(slMsg As String)
    'Add horizontal scroll if required and add message to list box
    'The control pbcArial is used to get the approximate width of the text as the list box does not has a TextWidth command
    Dim llValue As Long
    Dim llRg As Long
    Dim llMaxWidth
    Dim llRet As Long
    Dim llRow As Long
    
    llMaxWidth = (pbcArial.TextWidth(slMsg))
    If llMaxWidth > lmMaxWidth Then
        lmMaxWidth = llMaxWidth
    End If
    If lmMaxWidth > lbcMsg.Width Then
        'Divide by 15 to convert units and add 120 for little extra room
        'Scale Mode is in Twips
        llValue = lmMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slMsg)
    If llRow < 0 Then
        lbcMsg.AddItem slMsg
        gLogMsg slMsg, "WegenerExportLog.Txt", False
    End If
End Sub

Private Function mExportGroup() As Integer
    Dim slGrpName As String
    Dim ilGroup As Integer
    
    mExportGroup = True
    For ilGroup = 0 To UBound(tmCustomGroupNames) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        csiXMLData "OT", "Group_Assignment", ""
        csiXMLData "CD", "Group_Name", Trim$(tmCustomGroupNames(ilGroup).sName)
        csiXMLData "CD", "Assignment_Method", "Complete"
        csiXMLData "CD", "Compel_Addr", Trim$(tmCustomGroupNames(ilGroup).sCategoryGroup)
        csiXMLData "CT", "Group_Assignment", ""
    Next ilGroup
End Function

Private Function mCreateGroupNames(slInAddress As String) As String
    Dim slAddress As String
    Dim ilPos As Integer
    Dim slNewAddress As String
    Dim slCustomNo As String
    Dim ilCustomNameLen As Integer
    Dim slStr As String
    Dim llRow As Long
    
    On Error GoTo mCreateGroupNamesErr:
    'Max length allowed is 58
    If Len(slInAddress) <= 58 Then
        mCreateGroupNames = slInAddress
        Exit Function
    End If
    'Test if group name previously defined
    llRow = SendMessageByString(lbcGroupDef.hwnd, LB_FINDSTRING, -1, slInAddress)
    If llRow >= 0 Then
        mCreateGroupNames = smGroupName(lbcGroupDef.ItemData(llRow))
        Exit Function
    End If
    ilCustomNameLen = 0
    slNewAddress = ""
    slAddress = slInAddress
    'Split up into 58 character group names
    ilPos = InStrRev(slAddress, "^", -1, vbTextCompare)
    Do While ilPos > 0
        If igExportSource = 2 Then DoEvents
        slStr = Mid$(slAddress, ilPos)
        If Len(slStr) + Len(slNewAddress) > 58 Then
            imCustomGroupNo = imCustomGroupNo + 1
            slCustomNo = Trim$(Str$(imCustomGroupNo))
            Do While Len(slCustomNo) < 4
                slCustomNo = "0" & slCustomNo
            Loop
            tmCustomGroupNames(UBound(tmCustomGroupNames)).sName = smCustomGroupName & slCustomNo
            tmCustomGroupNames(UBound(tmCustomGroupNames)).sCategoryGroup = Mid(slNewAddress, 2) 'Remove AND (^) symbol
            ReDim Preserve tmCustomGroupNames(0 To UBound(tmCustomGroupNames) + 1) As CUSTOMGROUPNAMES
            slNewAddress = "^" & smCustomGroupName & slCustomNo
            ilCustomNameLen = Len(slNewAddress) 'All custom groups will be of the same length
        End If
        slNewAddress = slStr & slNewAddress
        slAddress = Left$(slAddress, ilPos - 1)
        If Len(slAddress) + Len(smVehicleGroupName) + 1 + ilCustomNameLen <= 58 Then
            imCustomGroupNo = imCustomGroupNo + 1
            slCustomNo = Trim$(Str$(imCustomGroupNo))
            Do While Len(slCustomNo) < 4
                slCustomNo = "0" & slCustomNo
            Loop
            tmCustomGroupNames(UBound(tmCustomGroupNames)).sName = smCustomGroupName & slCustomNo
            tmCustomGroupNames(UBound(tmCustomGroupNames)).sCategoryGroup = Mid$(slNewAddress, 2) 'Remove AND (^) symbol
            ReDim Preserve tmCustomGroupNames(0 To UBound(tmCustomGroupNames) + 1) As CUSTOMGROUPNAMES
            slAddress = slAddress & "^" & smCustomGroupName & slCustomNo
            Exit Do
        End If
        ilPos = InStrRev(slAddress, "^", -1, vbTextCompare)
    Loop
    If igExportSource = 2 Then DoEvents
    lbcGroupDef.AddItem slInAddress
    lbcGroupDef.ItemData(lbcGroupDef.NewIndex) = UBound(smGroupName)
    smGroupName(UBound(smGroupName)) = slAddress
    ReDim Preserve smGroupName(0 To UBound(smGroupName) + 1) As String
    mCreateGroupNames = slAddress
    Exit Function
mCreateGroupNamesErr:
    Resume Next
End Function

Private Function mUpdateShttUsedForWegener() As Integer
    Dim ilShtt As Integer
    Dim ilRet As Integer
    '11/26/17
    Dim ilIndex As Integer
    Dim blRepopRequired As Boolean
    
    On Error GoTo ErrHand
    'SQLQuery = "Update SHTT Set shttUsedForWegener = '" & "N" & "', "
    'SQLQuery = SQLQuery & "shttSerialNo1 = '', shttSerialNo2 = '', shttPort = ''"
    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '    GoSub ErrHand:
    'End If
    '11/26/17
    blRepopRequired = False
    For ilShtt = 0 To UBound(tmWegenerImport) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        '12/13/08:  Because stations can have multi-ports and multi-serial #, don't update fields
        'SQLQuery = "Update SHTT Set shttUsedForWegener = '" & "Y" & "', "
        'SQLQuery = SQLQuery & "shttSerialNo1 = '" & tmWegenerImport(ilShtt).sSerialNo1 & "', shttSerialNo2 = '', shttPort = '" & tmWegenerImport(ilShtt).sPort & "'"
        SQLQuery = "Update SHTT Set shttUsedForWegener = '" & "Y" & "'"
        SQLQuery = SQLQuery & " Where shttCode = " & tmWegenerImport(ilShtt).iShttCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "WegenerExportLog.txt", "Export Wegener-mUpdateShttUsedForWegener"
            mUpdateShttUsedForWegener = False
            Exit Function
        End If
        ilIndex = gBinarySearchStationInfoByCode(tmWegenerImport(ilShtt).iShttCode)
        If ilIndex <> -1 Then
            tgStationInfoByCode(ilIndex).sUsedForWegener = "Y"
            ilIndex = gBinarySearchStation(Trim$(tgStationInfoByCode(ilShtt).sCallLetters))
            If ilIndex <> -1 Then
                tgStationInfo(ilIndex).sUsedForWegener = "Y"
            Else
                blRepopRequired = True
            End If
        Else
            blRepopRequired = True
        End If
    Next ilShtt
    '11/26/17
    gFileChgdUpdate "shtt.mkd", blRepopRequired
    ilRet = gPopStations()
    mUpdateShttUsedForWegener = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerExportLog.txt", "Export Wegener-mUpdateShttUsedForWegener"
    mUpdateShttUsedForWegener = False
    Exit Function
End Function

Private Function mFindWegenerIndex(slInGroup As String, ilImport As Integer, slGroup As String, slPort As String) As Integer
    Dim ilPort As Integer
    Dim slChar As String
    Dim llSerialNo1A As Long
    Dim llSerialNo1Port As Long
    
    mFindWegenerIndex = True
    slGroup = slInGroup
    If Mid(slGroup, Len(slGroup) - 1, 1) = "_" Then
        slPort = right$(slGroup, 1)
        slGroup = Left$(slGroup, Len(slGroup) - 2)
        If tmWegenerImport(ilImport).sPort <> slPort Then
            mFindWegenerIndex = False
        End If
    Else
        'Ignore names without port letter
        mFindWegenerIndex = False
    End If

End Function

Private Function mRemoveExtraFromCallLetters(slCallLetters) As String
    Dim ilPos As Integer
    Dim slTrueCallLetters As String
    Dim slTestBand As String
    
    
    slTrueCallLetters = Trim$(slCallLetters)
    slTestBand = "-AM"
    ilPos = 1
    Do
        If igExportSource = 2 Then DoEvents
        ilPos = InStr(ilPos, slTrueCallLetters, slTestBand, vbTextCompare)
        If ilPos <= 0 Then
            If slTestBand = "-HD" Then
                Exit Do
            ElseIf slTestBand = "-FM" Then
                slTestBand = "-HD"
            Else
                slTestBand = "-FM"
            End If
            ilPos = 1
        Else
            If ilPos + 2 < Len(slTrueCallLetters) Then
                If Mid$(slTrueCallLetters, ilPos + 3, 1) <> "/" Then
                    slTrueCallLetters = Left$(slTrueCallLetters, ilPos + 2) & Mid(slTrueCallLetters, ilPos + 4)
                Else
                    If slTestBand = "-HD" Then
                        Exit Do
                    ElseIf slTestBand = "-FM" Then
                        slTestBand = "-HD"
                    Else
                        slTestBand = "-FM"
                    End If
                End If
                ilPos = ilPos + 1
            Else
                Exit Do
            End If
        End If
    Loop
    mRemoveExtraFromCallLetters = slTrueCallLetters
End Function

Private Sub mTestISCI(slInISCI As String, slInAdvtName As String, llSdfCode As Long)
    Dim slAdvtName As String
    Dim ilChar As Integer
    Dim ilError As Integer
    Dim slChar As String
    Dim llAdf As Long
    
    On Error GoTo ErrHand
    
    ilError = False
    '12/27/08:  Allowed characters A-Z 0-9 -(Dash) _(Underscore)
    For ilChar = 1 To Len(slInISCI) Step 1
        If igExportSource = 2 Then DoEvents
        slChar = Mid(slInISCI, ilChar, 1)
        If (Asc(slChar) < Asc("A")) Or (Asc(slChar) > Asc("Z")) Then
            'If (Asc(slChar) < Asc("a")) Or (Asc(slChar) > Asc("z")) Then
                If (Asc(slChar) < Asc("0")) Or (Asc(slChar) > Asc("9")) Then
                    If (Asc(slChar) <> Asc("-")) And (Asc(slChar) <> Asc("_")) Then
                        ilError = True
                        Exit For
                    End If
                End If
            'End If
        End If
    Next ilChar
    
    'If InStr(1, slInISCI, " ", vbTextCompare) > 0 Then
    If ilError Then
        If Trim$(slInAdvtName) = "" Then
            If llSdfCode > 0 Then
                SQLQuery = "SELECT adfName "
                SQLQuery = SQLQuery + " FROM adf_Advertisers, sdf_Spot_Detail"
                SQLQuery = SQLQuery + " WHERE (sdfCode = " & llSdfCode
                SQLQuery = SQLQuery + " AND adfCode = sdfadfCode " & ")"
                Set err_rst = gSQLSelectCall(SQLQuery)
                If Not err_rst.EOF Then
                    slAdvtName = Trim$(err_rst!adfName)
                Else
                    slAdvtName = "Advertiser Name Missing"
                End If
            Else
                slAdvtName = "Network Fill"
            End If
            mAddMsgToList "ISCI Name " & slInISCI & " with " & slAdvtName & " has illegal characters"
        Else
            mAddMsgToList "ISCI Name " & slInISCI & " has illegal characters"
        End If
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerExportLog.txt", "Export Wegener-mTestISCI"
    Exit Sub
End Sub

Private Sub mSetPortsRequired()
    Dim ilImport As Integer
    Dim llVefIndex As Long
    Dim slPort As String
    
    On Error GoTo ErrHandle
    imPortDefined(0) = False
    imPortDefined(1) = False
    imPortDefined(2) = False
    imPortDefined(3) = False
    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llVefIndex = tmWegenerImport(ilImport).lVefCodeFirst
        Do While llVefIndex <> -1
            If igExportSource = 2 Then DoEvents
            If (tmWegenerVehInfo(llVefIndex).iVefCode = imVefCode) Then
                slPort = tmWegenerVehInfo(llVefIndex).sPort
                If slPort = "A" Then
                    imPortDefined(0) = True
                ElseIf slPort = "B" Then
                    imPortDefined(1) = True
                ElseIf slPort = "C" Then
                    imPortDefined(2) = True
                ElseIf slPort = "D" Then
                    imPortDefined(3) = True
                End If
                Exit Do
            End If
            llVefIndex = tmWegenerVehInfo(llVefIndex).lVefCodeNext
        Loop
    
    Next ilImport
    Exit Sub
ErrHandle:
    Resume Next

End Sub

Private Sub mSetWegenerSort()
    Dim ilLoop As Integer
    Dim ilCheck As Integer
    Dim ilFound As Integer
    
    For ilLoop = LBound(tmWegenerImport) To UBound(tmWegenerImport) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tmWegenerImport(ilLoop).iFormatCode > 0 Then
            ilFound = False
            For ilCheck = 0 To UBound(tmWegenerFormatSort) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmWegenerFormatSort(ilCheck).iFmtCode = tmWegenerImport(ilLoop).iFormatCode Then
                    ilFound = True
                    tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                    tmWegenerIndex(UBound(tmWegenerIndex)).iNext = tmWegenerFormatSort(ilCheck).iFirst
                    tmWegenerFormatSort(ilCheck).iFirst = UBound(tmWegenerIndex)
                    ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                    Exit For
                End If
            Next ilCheck
            If Not ilFound Then
                tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                tmWegenerIndex(UBound(tmWegenerIndex)).iNext = -1
                tmWegenerFormatSort(UBound(tmWegenerFormatSort)).iFmtCode = tmWegenerImport(ilLoop).iFormatCode
                tmWegenerFormatSort(UBound(tmWegenerFormatSort)).iFirst = UBound(tmWegenerIndex)
                ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                ReDim Preserve tmWegenerFormatSort(0 To UBound(tmWegenerFormatSort) + 1) As WEGENERFORMATSORT
            End If
        End If
        If tmWegenerImport(ilLoop).iTztCode > 0 Then
            ilFound = False
            For ilCheck = 0 To UBound(tmWegenerTimeZoneSort) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmWegenerTimeZoneSort(ilCheck).iTztCode = tmWegenerImport(ilLoop).iTztCode Then
                    ilFound = True
                    tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                    tmWegenerIndex(UBound(tmWegenerIndex)).iNext = tmWegenerTimeZoneSort(ilCheck).iFirst
                    tmWegenerTimeZoneSort(ilCheck).iFirst = UBound(tmWegenerIndex)
                    ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                    Exit For
                End If
            Next ilCheck
            If Not ilFound Then
                tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                tmWegenerIndex(UBound(tmWegenerIndex)).iNext = -1
                tmWegenerTimeZoneSort(UBound(tmWegenerTimeZoneSort)).iTztCode = tmWegenerImport(ilLoop).iTztCode
                tmWegenerTimeZoneSort(UBound(tmWegenerTimeZoneSort)).iFirst = UBound(tmWegenerIndex)
                ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                ReDim Preserve tmWegenerTimeZoneSort(0 To UBound(tmWegenerTimeZoneSort) + 1) As WEGENERTIMEZONESORT
            End If
        End If
        If Trim$(tmWegenerImport(ilLoop).sPostalName) <> "" Then
            ilFound = False
            For ilCheck = 0 To UBound(tmWegenerPostalSort) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmWegenerPostalSort(ilCheck).sPostalName = tmWegenerImport(ilLoop).sPostalName Then
                    ilFound = True
                    tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                    tmWegenerIndex(UBound(tmWegenerIndex)).iNext = tmWegenerPostalSort(ilCheck).iFirst
                    tmWegenerPostalSort(ilCheck).iFirst = UBound(tmWegenerIndex)
                    ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                    Exit For
                End If
            Next ilCheck
            If Not ilFound Then
                tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                tmWegenerIndex(UBound(tmWegenerIndex)).iNext = -1
                tmWegenerPostalSort(UBound(tmWegenerPostalSort)).sPostalName = tmWegenerImport(ilLoop).sPostalName
                tmWegenerPostalSort(UBound(tmWegenerPostalSort)).iFirst = UBound(tmWegenerIndex)
                ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                ReDim Preserve tmWegenerPostalSort(0 To UBound(tmWegenerPostalSort) + 1) As WEGENERPOSTALSORT
            End If
        End If
        If tmWegenerImport(ilLoop).iMktCode > 0 Then
            ilFound = False
            For ilCheck = 0 To UBound(tmWegenerMarketSort) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmWegenerMarketSort(ilCheck).iMktCode = tmWegenerImport(ilLoop).iMktCode Then
                    ilFound = True
                    tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                    tmWegenerIndex(UBound(tmWegenerIndex)).iNext = tmWegenerMarketSort(ilCheck).iFirst
                    tmWegenerMarketSort(ilCheck).iFirst = UBound(tmWegenerIndex)
                    ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                    Exit For
                End If
            Next ilCheck
            If Not ilFound Then
                tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                tmWegenerIndex(UBound(tmWegenerIndex)).iNext = -1
                tmWegenerMarketSort(UBound(tmWegenerMarketSort)).iMktCode = tmWegenerImport(ilLoop).iMktCode
                tmWegenerMarketSort(UBound(tmWegenerMarketSort)).iFirst = UBound(tmWegenerIndex)
                ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                ReDim Preserve tmWegenerMarketSort(0 To UBound(tmWegenerMarketSort) + 1) As WEGENERMARKETSORT
            End If
        End If
        If tmWegenerImport(ilLoop).iMSAMktCode > 0 Then
            ilFound = False
            For ilCheck = 0 To UBound(tmWegenerMSAMarketSort) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmWegenerMSAMarketSort(ilCheck).iMktCode = tmWegenerImport(ilLoop).iMSAMktCode Then
                    ilFound = True
                    tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                    tmWegenerIndex(UBound(tmWegenerIndex)).iNext = tmWegenerMSAMarketSort(ilCheck).iFirst
                    tmWegenerMSAMarketSort(ilCheck).iFirst = UBound(tmWegenerIndex)
                    ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                    Exit For
                End If
            Next ilCheck
            If Not ilFound Then
                tmWegenerIndex(UBound(tmWegenerIndex)).iIndex = ilLoop
                tmWegenerIndex(UBound(tmWegenerIndex)).iNext = -1
                tmWegenerMSAMarketSort(UBound(tmWegenerMSAMarketSort)).iMktCode = tmWegenerImport(ilLoop).iMSAMktCode
                tmWegenerMSAMarketSort(UBound(tmWegenerMSAMarketSort)).iFirst = UBound(tmWegenerIndex)
                ReDim Preserve tmWegenerIndex(0 To UBound(tmWegenerIndex) + 1) As WEGENERINDEX
                ReDim Preserve tmWegenerMSAMarketSort(0 To UBound(tmWegenerMSAMarketSort) + 1) As WEGENERMARKETSORT
            End If
        End If
    Next ilLoop
    If igExportSource = 2 Then DoEvents
    If UBound(tmWegenerFormatSort) - 1 > 1 Then
        ArraySortTyp fnAV(tmWegenerFormatSort(), 0), UBound(tmWegenerFormatSort), 0, LenB(tmWegenerFormatSort(0)), 0, -1, 0
    End If
    If igExportSource = 2 Then DoEvents
    If UBound(tmWegenerTimeZoneSort) - 1 > 1 Then
        ArraySortTyp fnAV(tmWegenerTimeZoneSort(), 0), UBound(tmWegenerTimeZoneSort), 0, LenB(tmWegenerTimeZoneSort(0)), 0, -1, 0
    End If
    If igExportSource = 2 Then DoEvents
    If UBound(tmWegenerPostalSort) - 1 > 1 Then
        ArraySortTyp fnAV(tmWegenerPostalSort(), 0), UBound(tmWegenerPostalSort), 0, LenB(tmWegenerPostalSort(0)), 0, LenB(tmWegenerPostalSort(0).sPostalName), 0
    End If
    If igExportSource = 2 Then DoEvents
    If UBound(tmWegenerMarketSort) - 1 > 1 Then
        ArraySortTyp fnAV(tmWegenerMarketSort(), 0), UBound(tmWegenerMarketSort), 0, LenB(tmWegenerMarketSort(0)), 0, -1, 0
    End If
    If igExportSource = 2 Then DoEvents
    If UBound(tmWegenerMSAMarketSort) - 1 > 1 Then
        ArraySortTyp fnAV(tmWegenerMSAMarketSort(), 0), UBound(tmWegenerMSAMarketSort), 0, LenB(tmWegenerMSAMarketSort(0)), 0, -1, 0
    End If
End Sub

Public Function mBinarySearchFormat(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    ilMin = LBound(tmWegenerFormatSort)
    ilMax = UBound(tmWegenerFormatSort) - 1
    Do While ilMin <= ilMax
        If igExportSource = 2 Then DoEvents
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmWegenerFormatSort(ilMiddle).iFmtCode Then
            'found the match
            mBinarySearchFormat = ilMiddle
            Exit Function
        ElseIf ilCode < tmWegenerFormatSort(ilMiddle).iFmtCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchFormat = -1
    Exit Function
    
End Function

Public Function mBinarySearchTimeZone(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    ilMin = LBound(tmWegenerTimeZoneSort)
    ilMax = UBound(tmWegenerTimeZoneSort) - 1
    Do While ilMin <= ilMax
        If igExportSource = 2 Then DoEvents
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmWegenerTimeZoneSort(ilMiddle).iTztCode Then
            'found the match
            mBinarySearchTimeZone = ilMiddle
            Exit Function
        ElseIf ilCode < tmWegenerTimeZoneSort(ilMiddle).iTztCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchTimeZone = -1
    Exit Function
    
End Function

Public Function mBinarySearchPostalName(slCode As String) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilResult As Integer
    
    ilMin = LBound(tmWegenerPostalSort)
    ilMax = UBound(tmWegenerPostalSort) - 1
    Do While ilMin <= ilMax
        If igExportSource = 2 Then DoEvents
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        ilResult = StrComp(Trim(tmWegenerPostalSort(ilMiddle).sPostalName), slCode, vbTextCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchPostalName = ilMiddle  ' Found it !
                Exit Function
            Case 1:
                ilMax = ilMiddle - 1
            Case -1:
                ilMin = ilMiddle + 1
        End Select
    Loop
    mBinarySearchPostalName = -1
    Exit Function
    
End Function

Public Function mBinarySearchMarket(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    ilMin = LBound(tmWegenerMarketSort)
    ilMax = UBound(tmWegenerMarketSort) - 1
    Do While ilMin <= ilMax
        If igExportSource = 2 Then DoEvents
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmWegenerMarketSort(ilMiddle).iMktCode Then
            'found the match
            mBinarySearchMarket = ilMiddle
            Exit Function
        ElseIf ilCode < tmWegenerMarketSort(ilMiddle).iMktCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMarket = -1
    Exit Function
    
End Function

Public Function mBinarySearchMSAMarket(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    
    ilMin = LBound(tmWegenerMSAMarketSort)
    ilMax = UBound(tmWegenerMSAMarketSort) - 1
    Do While ilMin <= ilMax
        If igExportSource = 2 Then DoEvents
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmWegenerMSAMarketSort(ilMiddle).iMktCode Then
            'found the match
            mBinarySearchMSAMarket = ilMiddle
            Exit Function
        ElseIf ilCode < tmWegenerMSAMarketSort(ilMiddle).iMktCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMSAMarket = -1
    Exit Function
    
End Function


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
                If igExportSource = 2 Then DoEvents
                slDate = Format$(llDate, "m/d/yy")
                ilRet = gAlertClear("A", "F", "S", imVefCode, slDate)
                ilRet = gAlertClear("A", "R", "S", imVefCode, slDate)
            Next llDate
        End If
    Next ilVef
    ilRet = gAlertForceCheck()
End Sub

Private Sub udcCriteria_WGenerate(ilIndex As Integer, ilValue As Integer)
    If ilIndex = 1 Then
        If ilValue = vbChecked Then
            edcStartDate.Enabled = True
            'edcDate.Enabled = True
            edcDate.SetEnabled True
            edcDays.Enabled = True
            txtNumberDays.Enabled = True
            'ckcGenCSV.Enabled = True
            lbcVehicles.Enabled = True
        Else
            edcStartDate.Enabled = False
            'edcDate.Enabled = False
            edcDate.SetEnabled False
            edcDays.Enabled = False
            txtNumberDays.Enabled = False
            'ckcGenCSV.Enabled = False
            lbcVehicles.Enabled = False
        End If
    End If
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
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("W", "Wegener", "W", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function
Private Sub mGenRegionErrorMsg(ilStartBreakIndex As Integer, ilEndBreakIndex As Integer)
    Dim slBreakDate As String
    Dim slBreakTime As String
    Dim slRegionError As String
    Dim ilIndex As Integer
    bmRegionMaxExceeded = True
    For ilIndex = ilStartBreakIndex To ilEndBreakIndex Step 1
        If lmPrevSdfCode <> tmRegionBreakSpots(ilIndex).lSdfCode Then
            lmPrevSdfCode = tmRegionBreakSpots(ilIndex).lSdfCode
            slBreakDate = Format(tmRegionBreakSpots(ilIndex).lLogDate, sgShowDateForm)
            slBreakTime = gLongToTime(tmRegionBreakSpots(ilIndex).lLogTime)
            SQLQuery = "SELECT vefName, adfName "
            SQLQuery = SQLQuery + " FROM sdf_Spot_Detail Left Outer Join vef_Vehicles On sdfVefCode = vefCode"
            SQLQuery = SQLQuery + " Left Outer Join adf_Advertisers on sdfAdfCode = adfCode"
            SQLQuery = SQLQuery + " WHERE (sdfCode = " & tmRegionBreakSpots(ilIndex).lSdfCode & ")"
            Set err_rst = cnn.Execute(SQLQuery)
            If Not err_rst.EOF Then
                slRegionError = "Region Break Definitions exceeded: " & Trim$(err_rst!vefName) & " for " & Trim$(err_rst!adfName) & " on " & slBreakDate & " " & slBreakTime
                gLogMsg slRegionError, "WegenerExportLog.Txt", False
            End If
        End If
    Next ilIndex
End Sub

Private Sub mSaveStationCategory(tlRegionBreakSpotInfo As REGIONBREAKSPOTINFO, tlSplitCategoryInfo() As SPLITCATEGORYINFO)
    Dim llIndex As Long
    Dim llStnNext As Long
    Dim blAddStation As Boolean
    Dim ilImport As Integer
    Dim llVefIndex As Long
    Dim blStnFd As Boolean
    Dim blFirst As Boolean
    Dim ilMethod As Integer
    Dim ilShtt As Integer
    
    
    For llIndex = tlRegionBreakSpotInfo.lStartIndex To tlRegionBreakSpotInfo.lEndIndex Step 1
        If tmRegionDefinition(llIndex).iStationCount > 0 Then
            blFirst = True
            tmRegionDefinition(llIndex).iStationCount = 0
            llStnNext = tmRegionDefinition(llIndex).lStationOtherFirst
            Do
                blAddStation = True
                If blFirst Then
                    tmRegionDefinition(llIndex).lStationOtherFirst = UBound(tmStationSplitCategoryInfo)
                    blFirst = False
                Else
                    blAddStation = False
                    blStnFd = False
                    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        If tlSplitCategoryInfo(llStnNext).iIntCode = tmWegenerImport(ilImport).iShttCode Then
                            blAddStation = True
                            tmStationSplitCategoryInfo(UBound(tmStationSplitCategoryInfo) - 1).lNext = UBound(tmStationSplitCategoryInfo)
                            Exit For
                        End If
                    Next ilImport
                End If
                If blAddStation Then
                    tmRegionDefinition(llIndex).iStationCount = tmRegionDefinition(llIndex).iStationCount + 1
                    tmStationSplitCategoryInfo(UBound(tmStationSplitCategoryInfo)) = tlSplitCategoryInfo(llStnNext)
                    tmStationSplitCategoryInfo(UBound(tmStationSplitCategoryInfo)).lNext = -1
                    ReDim Preserve tmStationSplitCategoryInfo(0 To UBound(tmStationSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                End If
                llStnNext = tlSplitCategoryInfo(llStnNext).lNext
            Loop While llStnNext <> -1
            
        End If
    Next llIndex
End Sub

Private Sub mFilterStations(tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO)
    Dim llRegion As Long
    Dim ilStationCount As Integer
    Dim llOtherIndex As Integer
    Dim blAddStation As Boolean
    Dim blStnFd As Boolean
    Dim llStnNext As Long
    Dim llLastStnIndex As Long
    Dim ilImport As Integer
    Dim blFirst As Boolean
    Dim llVefIndex As Long
    Dim llVefIndex1 As Long
    Dim llVefIndex2 As Long
    Dim ilShtt As Integer
    Dim ilMethod As Integer
    Dim llStnIndex As Long
    Dim llLoop As Long
    Dim blFd As Boolean
    
    '2/7/20
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        If (tlRegionDefinition(llRegion).lOtherFirst <> -1) And (tlRegionDefinition(llRegion).lFormatFirst = -1) And (tlRegionDefinition(llRegion).lExcludeFirst = -1) Then
            llStnNext = tlRegionDefinition(llRegion).lOtherFirst
            Do
                If tlSplitCategoryInfo(llStnNext).sCategory = "S" Then
                    If tlSplitCategoryInfo(llStnNext).lLongCode = 0 Then
                        blStnFd = False
                        For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            If tlSplitCategoryInfo(llStnNext).iIntCode = tmWegenerImport(ilImport).iShttCode Then
'                                blStnFd = True
                                llVefIndex1 = tmWegenerImport(ilImport).lVefCodeFirst
                                Do While llVefIndex1 <> -1
                                    If (tmWegenerVehInfo(llVefIndex1).iVefCode = imVefCode) Then
                                    'If (tmWegenerVehInfo(llVefIndex1).iVefCode = imVefCode) And (tmWegenerVehInfo(llVefIndex1).sPort = tmWegenerImport(ilImport).sPort) Then
                                        blStnFd = True
                                        llStnIndex = llStnNext
                                        tlSplitCategoryInfo(llStnIndex).lLongCode = Val(tmWegenerImport(ilImport).sSerialNo1)
                                        For ilShtt = ilImport + 1 To UBound(tmWegenerImport) - 1 Step 1
                                            If tlSplitCategoryInfo(llStnIndex).iIntCode = tmWegenerImport(ilShtt).iShttCode Then
                                                'llVefIndex2 = tmWegenerImport(ilImport).lVefCodeFirst
                                                llVefIndex2 = tmWegenerImport(ilShtt).lVefCodeFirst
                                                Do While llVefIndex2 <> -1
                                                    If (tmWegenerVehInfo(llVefIndex2).iVefCode = imVefCode) Then
                                                    'If (tmWegenerVehInfo(llVefIndex2).iVefCode = imVefCode) And (tmWegenerVehInfo(llVefIndex2).sPort = tmWegenerImport(ilShtt).sPort) Then
                                                        If tlSplitCategoryInfo(llStnIndex).lLongCode <> Val(tmWegenerImport(ilShtt).sSerialNo1) Then
                                                            'Create another tlSplitCategory
                                                            blFd = False
                                                            For llLoop = LBound(tlSplitCategoryInfo) To UBound(tlSplitCategoryInfo) - 1 Step 1
                                                                If Val(tmWegenerImport(ilShtt).sSerialNo1) = tlSplitCategoryInfo(llLoop).lLongCode Then
                                                                    blFd = True
                                                                    Exit For
                                                                End If
                                                            Next llLoop
                                                            If blFd = False Then
                                                                tlSplitCategoryInfo(UBound(tlSplitCategoryInfo)) = tlSplitCategoryInfo(llStnIndex)
                                                                tlSplitCategoryInfo(UBound(tlSplitCategoryInfo)).lLongCode = Val(tmWegenerImport(ilShtt).sSerialNo1)
                                                                tlSplitCategoryInfo(UBound(tlSplitCategoryInfo)).lNext = tlSplitCategoryInfo(llStnIndex).lNext
                                                                tlSplitCategoryInfo(llStnNext).lNext = UBound(tlSplitCategoryInfo)
                                                                ReDim Preserve tlSplitCategoryInfo(LBound(tlSplitCategoryInfo) To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                                                            End If
                                                            'Exit For
                                                        End If
                                                        Exit Do
                                                    End If
                                                    llVefIndex2 = tmWegenerVehInfo(llVefIndex2).lVefCodeNext
                                                Loop
                                            End If
                                        Next ilShtt
                                        Exit Do
                                    End If
                                    llVefIndex1 = tmWegenerVehInfo(llVefIndex1).lVefCodeNext
                                Loop
                                If blStnFd Then
                                    Exit For
                                End If
                           End If
                        Next ilImport
                    End If
                End If
                llStnNext = tlSplitCategoryInfo(llStnNext).lNext
            Loop While llStnNext <> -1
        End If
    Next llRegion
    
    
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        If (tlRegionDefinition(llRegion).lOtherFirst <> -1) And (tlRegionDefinition(llRegion).lFormatFirst = -1) And (tlRegionDefinition(llRegion).lExcludeFirst = -1) Then
            '3/3/18: test if only stations.  If so don't expand
            ilStationCount = 0
            llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
            Do
                If tlSplitCategoryInfo(llOtherIndex).sCategory <> "S" Then
                    ilStationCount = -1
                    Exit Do
                End If
                ilStationCount = ilStationCount + 1
                llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
            If ilStationCount > 0 Then
                blFirst = True
                tlRegionDefinition(llRegion).iStationCount = 0
                llStnNext = tlRegionDefinition(llRegion).lOtherFirst
                tlRegionDefinition(llRegion).lOtherFirst = -1
                Do
                    blAddStation = False
                    blStnFd = False
                    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        If tlSplitCategoryInfo(llStnNext).iIntCode = tmWegenerImport(ilImport).iShttCode Then
                            If tlSplitCategoryInfo(llStnNext).lLongCode = Val(tmWegenerImport(ilImport).sSerialNo1) Then
                                blStnFd = True
                                llVefIndex = tmWegenerImport(ilImport).lVefCodeFirst
                                Do While llVefIndex <> -1
                                    If (tmWegenerVehInfo(llVefIndex).iVefCode = imVefCode) Then
                                        blAddStation = True
                                        Exit Do
                                    End If
                                    llVefIndex = tmWegenerVehInfo(llVefIndex).lVefCodeNext
                                Loop
                                If blAddStation Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilImport
                    If blAddStation Then
                        If blFirst Then
                            blFirst = False
                            tlRegionDefinition(llRegion).lOtherFirst = llStnNext
                        Else
                            tlSplitCategoryInfo(llLastStnIndex).lNext = llStnNext
                        End If
                        tlRegionDefinition(llRegion).iStationCount = tlRegionDefinition(llRegion).iStationCount + 1
                        llLastStnIndex = llStnNext
                        llStnNext = tlSplitCategoryInfo(llStnNext).lNext
                        tlSplitCategoryInfo(llLastStnIndex).lNext = -1
                    Else
                        llStnNext = tlSplitCategoryInfo(llStnNext).lNext
                    End If
                Loop While llStnNext <> -1
            End If
        End If
    Next llRegion
End Sub

Private Sub mBuildAllowStationList()
    Dim ilImport As Integer
    Dim llVefIndex As Long
    ReDim imAllowedShttCode(0 To 0) As Integer
    
    For ilImport = 0 To UBound(tmWegenerImport) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llVefIndex = tmWegenerImport(ilImport).lVefCodeFirst
        Do While llVefIndex <> -1
            If (tmWegenerVehInfo(llVefIndex).iVefCode = imVefCode) Then
                imAllowedShttCode(UBound(imAllowedShttCode)) = tmWegenerImport(ilImport).iShttCode
                ReDim Preserve imAllowedShttCode(0 To UBound(imAllowedShttCode) + 1) As Integer
                Exit Do
            End If
            llVefIndex = tmWegenerVehInfo(llVefIndex).lVefCodeNext
        Loop
    Next ilImport
    
End Sub

' JD 01-25-24 Added support for the new Wegener Check Utility
'
' Run the utility to download the RX_Calls and JNSGroup files regardless of the utility check box.
Private Function RunCheckUtility()
    Dim slCheckUtilityProgram As String
    Dim slCheckUtilityResults As String
    Dim hlFrom As Integer
    Dim slLine As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    RunCheckUtility = True
    Screen.MousePointer = vbHourglass
    
    slCheckUtilityProgram = sgExeDirectory & "WegenerUtility.exe"
    If gFileExist(slCheckUtilityProgram) = 1 Then
        chkRunCheckUtility.Value = vbUnchecked
        Screen.MousePointer = vbDefault
        gMsgBox "The program " & slCheckUtilityProgram & " does not exist. The Wegener files may need to be downloaded manually."
        Exit Function
    End If
    
    slCheckUtilityResults = smExportDirectory & "WegenerCheckResults.txt"
    If chkRunCheckUtility.Value <> vbChecked Then
        ' When not checked, we still want to download the files.
        mAddMsgToList "Downloading Wegener files"
        gShellAndWait slCheckUtilityProgram & " DOWNLOADFILES " & slCheckUtilityResults
    Else
        ' When it is checked the download will be run first as part of the overall process.
        mAddMsgToList "Running Check Wegener Utility"
        gShellAndWait slCheckUtilityProgram & " AUTO " & slCheckUtilityResults
    End If
    
    ilRet = gFileOpen(slCheckUtilityResults, "Input Access Read", hlFrom)
    If ilRet = 1 Then
        mAddMsgToList "The utility did not return any results. Continuing export."
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Not EOF(hlFrom) Then
        Line Input #hlFrom, slLine
    End If
    Close hlFrom
    If slLine = "Success" Then
        mAddMsgToList slLine
    Else
        If slLine = "Cancel" Then
            mAddMsgToList "Export was canceled."
            imExporting = False
            RunCheckUtility = False
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            mAddMsgToList "Ignoring utility check errors."
        End If
    End If
    gDeleteFile (slCheckUtilityResults)
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    ' ignore all errors.
End Function
