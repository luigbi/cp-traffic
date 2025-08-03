VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form FrmExportIDC 
   Caption         =   "Export IDC"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "AffExportIDC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdSites 
      Caption         =   "Check &Sites"
      Height          =   375
      Left            =   1170
      TabIndex        =   17
      Top             =   5445
      Width           =   1665
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
      Left            =   9105
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   4185
      Visible         =   0   'False
      Width           =   1035
   End
   Begin V81Affiliate.CSI_Calendar edcStartDate 
      Height          =   285
      Left            =   1515
      TabIndex        =   1
      Top             =   150
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      Text            =   "11/8/2010"
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
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1560
      Width           =   3810
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   1560
      Width           =   1635
   End
   Begin VB.ListBox lbcSort 
      Height          =   255
      Left            =   8205
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   5145
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7710
      Top             =   5040
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   3990
      TabIndex        =   3
      Text            =   "1"
      Top             =   165
      Width           =   405
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4245
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lbcStation 
      Height          =   2595
      ItemData        =   "AffExportIDC.frx":08CA
      Left            =   4230
      List            =   "AffExportIDC.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1815
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   4560
      Width           =   570
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3960
      ItemData        =   "AffExportIDC.frx":08CE
      Left            =   6615
      List            =   "AffExportIDC.frx":08D0
      TabIndex        =   10
      Top             =   435
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2595
      ItemData        =   "AffExportIDC.frx":08D2
      Left            =   120
      List            =   "AffExportIDC.frx":08D4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1815
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   885
      Top             =   5610
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5865
      FormDesignWidth =   9975
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   3015
      TabIndex        =   11
      Top             =   5460
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4980
      TabIndex        =   12
      Top             =   5460
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3105
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   360
      Left            =   45
      TabIndex        =   13
      Top             =   5115
      Width           =   7530
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7065
      TabIndex        =   9
      Top             =   105
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
      Caption         =   "Guide"
      Visible         =   0   'False
      Begin VB.Menu mnuCsiTest 
         Caption         =   "Csi-Test"
      End
   End
End
Attribute VB_Name = "FrmExportIDC"
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
'Dan M 10/8/10 replaced dateValue with gDateValue throughout form
Private imGenerating As Integer
Private smDate As String     'Export Date
Private smEndDate As String
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
Private raf_rst As ADODB.Recordset
Private rsf_rst As ADODB.Recordset
Private ief_rst As ADODB.Recordset
Private crf_rst As ADODB.Recordset
Private cnf_rst As ADODB.Recordset
Private cif_rst As ADODB.Recordset
Private cpf_rst As ADODB.Recordset
Private eht_rst As ADODB.Recordset
Private evt_rst As ADODB.Recordset
Private ect_rst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmIDCGeneric() As IDCGENERIC
Private tmIDCSplit() As IDCSPLIT
Private tmIDCReceiver() As IDCRECEIVER
Private bmMgsPrevExisted As Boolean
Private hmCSV As Integer
Private hmRegionISCI As Integer
Private hmCsf As Integer
Dim smNowDate As String
'Dan M for writing messages in list box
Private lmMaxWidth As Long
Private Type ISCIBYPERCENT
    iPercent As Integer
    sFilterRISCI As String * 20
    sUnfilterRISCI As String * 20
    lNextRISCI As Long
End Type
Private tmISCIByPercent() As ISCIBYPERCENT
Private Type ISCICOUNT
    lCifCode As Long
    iCount As Integer
    iPercent As Integer
End Type
Private tmISCICount() As ISCICOUNT
Private Type GENERICCIF
    lCifCode As Long
    sISCI As String * 20
End Type
Private tmCifCode() As GENERICCIF
Private Const FILEFACTS As String = "IDCFacts"
Private Const FILEERROR As String = "IDCExport"
Private Const FILEDEBUG As String = "IDCDebug"
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const XMLDATE As String = "yyyy-mm-dd"
Private lmEqtCode As Long
'5882
Private mRsRotations As ADODB.Recordset
Private Const XMLTIME As String = "hh:mm:ss"
'logging:
Private myFacts As CLogger
Private smPathForgLogMsg As String
Private myErrors As CLogger
'internal guide chose to run 'test' send to test server, so change values as needed
Dim bmCsiTest As Boolean
'5206 clean log
Dim bmIsExportClean As Boolean
Dim bmIsMessageClean As Boolean
'show adv in facts log
Dim myAdvDictionary As Dictionary
'6419 grouping--really blackout
'Dim rsBlackout As ADODB.Recordset
'6514
Private mRsCrfSplit As ADODB.Recordset

Private Sub mFillVehicle()
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim llVef As Long
    
    On Error GoTo ErrHand
    lbcVehicles.Clear
    
    chkAll.Value = vbUnchecked
    slNowDate = Format(gNow(), sgSQLDateForm)
    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND RTrim(attIDCReceiverID) <> ''"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        llVef = gBinarySearchVef(CLng(rst!attvefCode))
        If llVef <> -1 Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = rst!attvefCode
        End If
        rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mFileVehicle"
    Resume Next
    Exit Sub
IndexErr:
    ilRet = 1
    Resume Next
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
    mCleanFolders
    mExport
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcStartDate.Text = ""
    Unload FrmExportIDC
End Sub
Private Sub mCleanFolders()
    '5206 clean exports and messages, but only once
    If Not myErrors Is Nothing Then
        With myErrors
            If Not bmIsMessageClean Then
                .CleanThisFolder = messages
                '8886
'                .CleanFolder "IDC"
                .CleanFolder
                If Len(.ErrorMessage) > 0 Then
                    .WriteWarning "Couldn't delete old files from 'messages': " & .ErrorMessage
                Else
                    bmIsMessageClean = True
                End If
            End If
            If Not bmIsExportClean Then
                .CleanThisFolder = exports
                '8886
'                .CleanFolder "IDC"
                .CleanFolder
                If Len(.ErrorMessage) > 0 Then
                    .WriteWarning "Couldn't delete old files from 'exports': " & .ErrorMessage
                Else
                    bmIsExportClean = True
                End If
            End If
        End With
    End If
End Sub
Private Function mGetURL() As String
    Dim slIniPath As String
    Dim slRet As String
    
    slRet = ""
    slIniPath = gXmlIniPath(True)
    If LenB(slIniPath) > 0 Then
        mLoadFromIni "IDC", "URL", slIniPath, slRet
        If slRet = "Not Found" Then
            slRet = ""
        End If
    End If
    mGetURL = slRet
End Function
Private Function mGetBackupUrl() As String
    Dim slIniPath As String
    Dim slRet As String
    
    slRet = ""
    slIniPath = gXmlIniPath(True)
    If LenB(slIniPath) > 0 Then
        mLoadFromIni "IDC", "Backup", slIniPath, slRet
        If slRet = "Not Found" Then
            slRet = ""
        End If
    End If
    mGetBackupUrl = slRet
End Function

Private Sub cmdSites_Click()
    mTestSites
End Sub

Private Sub edcStartDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Sub edcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        udcCriteria.Left = lacStartDate.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        'udcCriteria.Top = edcStartDate.Top + (3 * edcStartDate.Height) / 4
        udcCriteria.Top = txtNumberDays.Top + (3 * txtNumberDays.Height / 2)
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
        If igExportSource = 2 Then
            slNowStart = gNow()
            edcStartDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
            sgExportResultName = "IDCResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "IDC Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "IDC Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.2
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts FrmExportIDC
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub
Private Sub mInit()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    cmdSites.Visible = False
    Screen.MousePointer = vbHourglass
    lmMaxWidth = lbcMsg.Width
    imTerminate = False
    imFirstTime = True
    bmIsMessageClean = False
    bmIsExportClean = False
    FrmExportIDC.Caption = "Export IDC - " & sgClientName
    'csi internal guide-for testing help
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuGuide.Visible = True
    End If
    Set myErrors = New CLogger
    myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEERROR)
    myErrors.CleanThisFolder = messages
    smPathForgLogMsg = FILEERROR & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcStartDate.Text = smDate
    txtNumberDays.Text = 7
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    If Not ilRet Then
        imTerminate = True
    End If
    ilRet = mOpenCSF()
    If Not ilRet Then
        imTerminate = True
    End If
    lbcStation.Clear
    mFillVehicle
    chkAll.Value = vbChecked
    If lbcVehicles.ListCount = 1 Then
        imVefCode = lbcVehicles.ItemData(0)
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    End If
    ilRet = gPopAvailNames()
    If Not ilRet Then
        imTerminate = True
    End If
    If imTerminate Then
        tmcTerminate.Enabled = True
    End If
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
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    ilRet = mCloseCSF()
    
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmIDCGeneric
    Erase tmIDCSplit
    Erase tmIDCReceiver
    Erase tmISCIByPercent
    Erase tmISCICount
    Erase tmCifCode
    
    cprst.Close
    raf_rst.Close
    rsf_rst.Close
    ief_rst.Close
    crf_rst.Close
    cnf_rst.Close
    cif_rst.Close
    cpf_rst.Close
    eht_rst.Close
    evt_rst.Close
    ect_rst.Close
    Set myFacts = Nothing
    Set myErrors = Nothing
    Set FrmExportIDC = Nothing
End Sub


Private Sub lbcStation_Click()
    If imAllStationClick Then
        Exit Sub
    End If
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
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



Private Sub mnuCsiTest_Click()
    mnuCsiTest.Checked = Not mnuCsiTest.Checked
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload FrmExportIDC
End Sub



Private Function mGatherIDC() As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim slStr As String
    Dim ilOkStation As Integer
    Dim ilOkVehicle As Integer
    Dim ilVef As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim ilIncludeSpot As Integer
    Dim ilIndex As Integer
    Dim slISCI As String
    Dim slRISCI As String
    Dim llRCrfCsfCode As Long
    Dim llRCrfCode As Long
    Dim llCrfCode As Long
    Dim llODate As Long
    Dim ilVpf As Integer
    Dim ilRegionExist As Integer
    Dim slVehicleName As String
    Dim slStationName As String
    Dim llTotalExport As Long
    Dim slIDCReceiverID As String
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
    
    Dim llLoopGeneric As Long
    Dim llIndexGeneric As Long
    Dim llIndexSplit As Long
    Dim llIndexReceiver As Long
    Dim blSplitFound As Boolean
    Dim blReceiverFound As Boolean
    Dim llAdf As Long
    Dim llCif As Long
    Dim llRRafCode As Long
    Dim slMissingCopy As String
    Dim ilLastVefCodeExported As Integer
    
    On Error GoTo ErrHand
    llTotalExport = 0
    imExporting = True
    imGenerating = 1
    ilLastVefCodeExported = -1
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    smEndDate = sEndDate
    slSDate = smDate
    slEDate = gObtainNextSunday(slSDate)
    If gDateValue(gAdjYear(sEndDate)) < gDateValue(gAdjYear(slEDate)) Then
        slEDate = sEndDate
    End If
    imVefCode = 0
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(ilVef) Then
            If imVefCode = 0 Then
                imVefCode = lbcVehicles.ItemData(ilVef)
            Else
                imVefCode = -1
                Exit For
            End If
        End If
    Next ilVef
    ReDim tmIDCGeneric(0 To 0) As IDCGENERIC
    ReDim tmIDCSplit(0 To 0) As IDCSPLIT
    ReDim tmIDCReceiver(0 To 0) As IDCRECEIVER
    ReDim tmISCIByPercent(0 To 0) As ISCIBYPERCENT
    Do
           If igExportSource = 2 Then DoEvents
        'Dan M 9/20/13 added vehicle ordering
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attVoiceTracked, attIDCReceiverID"
        '7701
        SQLQuery = SQLQuery & " FROM shtt, cptt, vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
        SQLQuery = SQLQuery & " WHERE ( RTrim(attIDCReceiverID) <> '' AND ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode AND vefcode = cpttVefCode"
        SQLQuery = SQLQuery & " AND vatWvtVendorId = " & Vendors.iDc
        'SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'I'"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        If imVefCode > 0 Then
            SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
        End If
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If igExportSource = 2 Then DoEvents
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
                ilOkVehicle = False
                For ilVef = 0 To lbcVehicles.ListCount - 1
                    If igExportSource = 2 Then DoEvents
                    If lbcVehicles.Selected(ilVef) Then
                        If lbcVehicles.ItemData(ilVef) = cprst!cpttvefcode Then
                            imVefCode = lbcVehicles.ItemData(ilVef)
                            ilOkVehicle = True
                            Exit For
                        End If
                    End If
                Next ilVef
            End If
            slIDCReceiverID = Trim$(cprst!attIDCReceiverID)
            If ilOkStation And ilOkVehicle And (slIDCReceiverID <> "") Then
                If (ilLastVefCodeExported > 0) And (ilLastVefCodeExported <> imVefCode) Then
                    If lbcStation.ListCount <= 0 Or chkAllStation.Value = vbChecked Then
                        'gClearAbf imVefCode, 0, sMoDate, gObtainNextSunday(sMoDate)
                    End If
                    ilLastVefCodeExported = imVefCode
                End If
                '9/6/11: Moved vehicle setting here because it can change when looking thru CPTT
                'Jeff
                If igExportSource = 2 Then DoEvents
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
                    tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
                    tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                    'Create AST records
                    imAdfCode = -1
                    igTimes = 1 'By Week
                    If igExportSource = 2 Then DoEvents
                    llODate = -1
                    '6442 was false,false,true
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                    gFilterAstExtendedTypes tmAstInfo
'                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                    ilIndex = LBound(tmAstInfo)
                    slVehicleName = mGetVehicleName(tmAstInfo(ilIndex).iVefCode)
                    slStationName = mGetStationName(tmAstInfo(ilIndex).iShttCode)
                    lacResult.Caption = "Generating Regional copy for " & slStationName & ", " & slVehicleName
                    Do While ilIndex < UBound(tmAstInfo)
                        If igExportSource = 2 Then DoEvents
                        blSpotOk = True
                        ilAnf = gBinarySearchAnf(tmAstInfo(ilIndex).iAnfCode)
                        If ilAnf <> -1 Then
                            If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                                blSpotOk = False
                            End If
                        End If
                        If (blSpotOk) And ((gDateValue(gAdjYear(tmAstInfo(ilIndex).sFeedDate)) >= gDateValue(gAdjYear(slSDate))) And (gDateValue(gAdjYear(tmAstInfo(ilIndex).sFeedDate)) <= gDateValue(gAdjYear(sEndDate)))) Then
                            llODate = gDateValue(gAdjYear(tmAstInfo(ilIndex).sFeedDate))
                            If (tgStatusTypes(gGetAirStatus(tmAstInfo(ilIndex).iStatus)).iPledged <> 2) Then
                                If tmAstInfo(ilIndex).iRegionType > 0 Then
'                                'test
'                                If tmAstInfo(ilIndex).lLstBkoutLstCode > 0 Then
'                                    llODate = llODate
'                                End If
                                'end test
                                    mGetGenericISCIForBlackouts tmAstInfo(ilIndex)
                                    'Loop on all Generic defined for the spot
                                    If tmAstInfo(ilIndex).lCifCode <> 0 Then
                                        ReDim tmCifCode(0 To 0) As GENERICCIF
                                        llCrfCode = gGetSdfCrfCode(tmAstInfo(ilIndex).lSdfCode)
                                        If llCrfCode = 0 Then
                                            tmCifCode(0).lCifCode = tmAstInfo(ilIndex).lCifCode
                                            tmCifCode(0).sISCI = tmAstInfo(ilIndex).sISCI
                                            ReDim Preserve tmCifCode(0 To 1) As GENERICCIF
                                        Else
                                            'Get all the generic copy
                                            mBuildGenericCif llCrfCode
                                        End If
                                        
                                        For llCif = 0 To UBound(tmCifCode) - 1 Step 1
                                            'Build arrays if Required
                                            llIndexGeneric = -1
                                            For llLoopGeneric = 0 To UBound(tmIDCGeneric) - 1 Step 1
                                                If igExportSource = 2 Then DoEvents
                                                If tmIDCGeneric(llLoopGeneric).lCifCode = tmCifCode(llCif).lCifCode Then
                                                    'If tmIDCGeneric(llLoopGeneric).lFeedDate = llODate Then
                                                        llIndexGeneric = llLoopGeneric
                                                        Exit For
                                                    'End If
                                                End If
                                            Next llLoopGeneric
                                            If llIndexGeneric = -1 Then
                                                llIndexGeneric = UBound(tmIDCGeneric)
                                                tmIDCGeneric(llIndexGeneric).lCifCode = tmCifCode(llCif).lCifCode
                                                tmIDCGeneric(llIndexGeneric).lCrfCsfCode = tmAstInfo(ilIndex).lCrfCsfCode
                                                tmIDCGeneric(llIndexGeneric).sISCI = tmCifCode(llCif).sISCI
                                                'tmIDCGeneric(llIndexGeneric).lFeedDate = llODate
                                                tmIDCGeneric(llIndexGeneric).lFirstSplit = -1
                                                tmIDCGeneric(llIndexGeneric).lTriggerId = 0
                                                ReDim Preserve tmIDCGeneric(0 To llIndexGeneric + 1) As IDCGENERIC
                                                'add advertiser to fact log
                                                If Not myAdvDictionary Is Nothing Then
                                                    If Not myAdvDictionary.Exists(gXMLNameFilter(tmCifCode(llCif).sISCI)) Then
                                                        myAdvDictionary.Add gXMLNameFilter(tmCifCode(llCif).sISCI), tmAstInfo(ilIndex).iAdfCode
                                                    End If
                                                End If
                                            End If
                                            blSplitFound = False
                                            llRRafCode = mGetRafCode(tmAstInfo(ilIndex).lRCrfCode)
                                            llIndexSplit = tmIDCGeneric(llIndexGeneric).lFirstSplit
                                            Do While llIndexSplit <> -1
                                                If igExportSource = 2 Then DoEvents
                                                'If tmIDCSplit(llIndexSplit).lRCifCode = tmAstInfo(ilIndex).lRCifCode Then
                                                'Dan--don't match on rotation, match on info within rotations
                                                '5882 2/22 go back to matching on rotation.
                                                If tmIDCSplit(llIndexSplit).lRCrfCode = tmAstInfo(ilIndex).lRCrfCode Then
                                                'If tmIDCSplit(llIndexSplit).lRafCode = llRRafCode Then
                                                    blSplitFound = True
                                                    Exit Do
                                                End If
                                                llIndexSplit = tmIDCSplit(llIndexSplit).lNextSplit
                                            Loop
                                            If Not blSplitFound Then
                                                If igExportSource = 2 Then DoEvents
                                                llIndexSplit = UBound(tmIDCSplit)
                                                tmIDCSplit(llIndexSplit).lRCrfCode = tmAstInfo(ilIndex).lRCrfCode
                                                'tmIDCSplit(llIndexSplit).lRCifCode = tmAstInfo(ilIndex).lRCifCode
                                                'tmIDCSplit(llIndexSplit).lRRsfCode = tmAstInfo(ilIndex).lRRsfCode
                                                'tmIDCSplit(llIndexSplit).lRafCode = mGetRafCode(tmAstInfo(ilIndex).lRCrfCode)
                                                tmIDCSplit(llIndexSplit).lRafCode = llRRafCode
                                                'tmIDCSplit(llIndexSplit).sRISCI = tmAstInfo(ilIndex).sRISCI
                                                tmIDCSplit(llIndexSplit).lNextSplit = tmIDCGeneric(llIndexGeneric).lFirstSplit
                                                llAdf = gBinarySearchAdf(CLng(tmAstInfo(ilIndex).iAdfCode))
                                                If llAdf <> -1 Then
                                                    tmIDCSplit(llIndexSplit).sAdvName = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                                                Else
                                                    tmIDCSplit(llIndexSplit).sAdvName = "Advertiser Name Missing"
                                                End If
                                                tmIDCSplit(llIndexSplit).lFirstReceiver = -1
                                                tmIDCSplit(llIndexSplit).lFirstRISCI = -1
                                                tmIDCGeneric(llIndexGeneric).lFirstSplit = llIndexSplit
                                                ReDim Preserve tmIDCSplit(0 To llIndexSplit + 1) As IDCSPLIT
                                            End If
                                            
                                            blReceiverFound = False
                                            llIndexReceiver = tmIDCSplit(llIndexSplit).lFirstReceiver
                                            Do While llIndexReceiver <> -1
                                                If igExportSource = 2 Then DoEvents
                                                If Trim$(tmIDCReceiver(llIndexReceiver).sReceiverID) = Trim$(slIDCReceiverID) Then
                                                    blReceiverFound = True
                                                    Exit Do
                                                End If
                                                llIndexReceiver = tmIDCReceiver(llIndexReceiver).lNextReceiver
                                            Loop
                                            If Not blReceiverFound Then
                                                If igExportSource = 2 Then DoEvents
                                                llIndexReceiver = UBound(tmIDCReceiver)
                                                tmIDCReceiver(llIndexReceiver).sReceiverID = slIDCReceiverID
                                                tmIDCReceiver(llIndexReceiver).sCallLetters = cprst!shttCallLetters
                                                tmIDCReceiver(llIndexReceiver).lAttCode = cprst!cpttatfCode
                                                tmIDCReceiver(llIndexReceiver).lNextReceiver = tmIDCSplit(llIndexSplit).lFirstReceiver
                                                tmIDCSplit(llIndexSplit).lFirstReceiver = llIndexReceiver
                                                ReDim Preserve tmIDCReceiver(0 To llIndexReceiver + 1) As IDCRECEIVER
                                            End If
                                            ' one record for each rotation/generic/siteID
'                                            If tmAstInfo(ilIndex).lLstBkoutLstCode > 0 Then
'                                                rsBlackout.Filter = "Index = " & llIndexSplit & " AND SiteId = '" & slIDCReceiverID & "' AND Generic = '" & tmIDCGeneric(llIndexGeneric).sISCI & "'"
'                                                If rsBlackout.EOF Then
'                                                    rsBlackout.AddNew Array("Index", "SiteID", "Generic"), Array(llIndexSplit, slIDCReceiverID, tmIDCGeneric(llIndexGeneric).sISCI)
'                                                End If
'                                            End If
                                        Next llCif
                                    Else 'cif = 0
                                        'error message: missing copy can't send to idc. Write out at end of method
                                        slMissingCopy = mMissingCopy(ilIndex, slMissingCopy)
                                    End If
                                End If
                            End If
                        End If
                        ilIndex = ilIndex + 1
                        If imTerminate Then
                            imExporting = False
                            Exit Function
                        End If
                    Loop
                End If
                If igExportSource = 2 Then DoEvents
            End If
            cprst.MoveNext
        Wend
        If imTerminate Then
            imExporting = False
            Exit Function
        End If
        If igExportSource = 2 Then DoEvents
        llODate = -1
        
        
        If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
             gClearASTInfo True
             '12/11/17: Clear abf
             'gClearAbf imVefCode, 0, sMoDate, gObtainNextSunday(sMoDate)
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
    If Len(slMissingCopy) > 0 Then
        mSetResults "Missing copy: cannot send some spot info to IDC: " & slMissingCopy, MESSAGERED
        myErrors.WriteWarning "Missing copy: cannot sent some spot info to IDC: " & slMissingCopy
    End If
    mGatherIDC = True
    Exit Function
mExportSpotInsertionsErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mGatherIDC"
    mGatherIDC = False
    Exit Function
    
End Function
Private Function mMissingCopy(ilIndex As Integer, slOldMessage As String) As String
    Dim slNewMessage As String
    Dim slTempMessage As String
    Dim ilFound As Integer
    Dim slAdv As String
    Dim llContract As Long
    Dim ilVefCode As Integer
    Dim slVehicle As String
    If igExportSource = 2 Then DoEvents
    ilFound = gBinarySearchAdf(CLng(tmAstInfo(ilIndex).iAdfCode))
    If ilFound <> -1 Then
        slAdv = Trim$(tgAdvtInfo(ilFound).sAdvtName)
    End If
    If igExportSource = 2 Then DoEvents
    ilFound = gBinarySearchVef(CLng(tmAstInfo(ilIndex).iVefCode))
    If ilFound <> -1 Then
        slVehicle = Trim$(tgVehicleInfo(ilFound).sVehicleName)
    End If
    If igExportSource = 2 Then DoEvents
    llContract = tmAstInfo(ilIndex).lCntrNo
    slTempMessage = slAdv & "-" & slVehicle & "-Contract#" & llContract
    If InStr(1, slOldMessage, slTempMessage) = 0 Then
        slNewMessage = slOldMessage & vbCrLf & slTempMessage
    Else
        slNewMessage = slOldMessage
    End If
    mMissingCopy = slNewMessage
    
End Function
Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode, attIDCReceiverID"
    SQLQuery = SQLQuery & " FROM shtt, att"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If Trim$(rst!attIDCReceiverID) <> "" Then
            lbcStation.AddItem Trim$(rst!shttCallLetters)
            lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
        End If
        rst.MoveNext
    Wend
    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mFileStations"

End Sub



Private Sub mSetResults(slMsg As String, llFGC As Long)
    Dim llLoop As Long
    bmMgsPrevExisted = False
    For llLoop = 0 To lbcMsg.ListCount - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If slMsg = lbcMsg.List(llLoop) Then
            bmMgsPrevExisted = True
            Exit Sub
        End If
    Next llLoop
    gAddMsgToListBox FrmExportIDC, lmMaxWidth, slMsg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
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

Private Function mGetStationName(iShttCode As Integer) As String
    Dim llLoop As Integer
    mGetStationName = ""
    For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgStationInfo(llLoop).iCode = iShttCode Then
            mGetStationName = Trim(tgStationInfo(llLoop).sCallLetters)
            Exit For
        End If
    Next
End Function

Private Function mFileNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    'Remove " and '
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    mFileNameFilter = slName
End Function

Private Sub mExport()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVehicleSelected As Integer
    Dim blIDC As Boolean
    
    On Error GoTo ErrHand
    
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    lacResult.Caption = ""
    imTerminate = False
    If (udcCriteria.DGenType(0) = vbUnchecked) And (udcCriteria.DGenType(2) = vbUnchecked) And (udcCriteria.DGenType(3) = vbUnchecked) Then
        gMsgBox "Either 'Send To...' or 'Generate...' must be specifed", vbOKOnly
        Exit Sub
    End If
    If edcStartDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        'edcStartDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcStartDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        'edcStartDate.SetFocus
        Exit Sub
    Else
        smDate = Format(edcStartDate.Text, sgShowDateForm)
    End If
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        'txtNumberDays.SetFocus
        Exit Sub
    End If
    Select Case Weekday(gAdjYear(smDate))
        Case vbMonday
            If imNumberDays > 7 Then
                gMsgBox "Number of days can not exceed 7.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbTuesday
            If imNumberDays > 6 Then
                gMsgBox "Number of days can not exceed 6.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbWednesday
            If imNumberDays > 5 Then
                gMsgBox "Number of days can not exceed 5.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbThursday
            If imNumberDays > 4 Then
                gMsgBox "Number of days can not exceed 4.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbFriday
            If imNumberDays > 3 Then
                gMsgBox "Number of days can not exceed 3.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
           End If
        Case vbSaturday
            If imNumberDays > 2 Then
                gMsgBox "Number of days can not exceed 2.", vbOKOnly
               ' txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbSunday
            If imNumberDays > 1 Then
                gMsgBox "Number of days can not exceed 1.", vbOKOnly
                'txtNumberDays.SetFocus
                Exit Sub
            End If
    End Select
    smNowDate = Format$(gNow(), "m/d/yy")
    If (udcCriteria.DGenType(0) = vbChecked) Then
        If gDateValue(gAdjYear(smDate)) <= gDateValue(gAdjYear(smNowDate)) Then
            Beep
            gMsgBox "Date must be after today's date " & smNowDate, vbCritical
           ' edcStartDate.SetFocus
            Exit Sub
        End If
    End If
  '  Set rsBlackout = mPrepRecordsetBlackout()
    'adv in facts log
    If udcCriteria.DGenType(2) = vbChecked Then
        Set myAdvDictionary = New Dictionary
    End If
    smExportDirectory = udcCriteria.DExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
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
    If Not gPopCopy(smDate, "Export IDC") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    On Error GoTo 0
    lacResult.Caption = ""
    If Not mCleanIef() Then
        mSetResults "Couldn't delete old records from table IEF. Please contact Counterpoint.", MESSAGERED
    End If
    Set mRsCrfSplit = mPrepRecordset()
    '8688
    bgTaskBlocked = False
    sgTaskBlockedName = "IDC Export"
    ilRet = mGatherIDC()
    gCloseRegionSQLRst
    If imTerminate Then
        Call mSetResults("** User Terminated **", RGB(255, 0, 0))
        myErrors.WriteFacts "*** User Terminated **", True
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
       ' cmdCancel.SetFocus
        Exit Sub
    End If
    If (ilRet = False) Then
        Call mSetResults("Export Failed", RGB(255, 0, 0))
        myErrors.WriteWarning "** Terminated - mExportSpotInsertions returned False **", True
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
       ' cmdCancel.SetFocus
        Exit Sub
    Else
       ' mFilterGeneric
        ilRet = mGetRegionISCI()
        If imTerminate Then
            Call mSetResults("** User Terminated **", RGB(255, 0, 0))
            myErrors.WriteFacts "*** User Terminated **", True
            ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
            imExporting = False
            Screen.MousePointer = vbDefault
          '  cmdCancel.SetFocus
            Exit Sub
        End If
        blIDC = True
        lacResult.Caption = ""
        If udcCriteria.DGenType(0) = vbChecked Or udcCriteria.DGenType(2) = vbChecked Then
            'Send to IDC or Generate file
            blIDC = mSendtoIdc()
            If blIDC Then
                '8688
                If bgTaskBlocked And igExportSource <> 2 Then
                     mSetResults "Some spots were blocked during export.", MESSAGERED
                     gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
                     myErrors.WriteWarning "Some spots were blocked during export.", True
                     lacResult.Caption = "Please refer to the Messages folder for file TaskBlocked_" & sgTaskBlockedDate & ".txt."
                End If
                mClearAlerts
                gCustomEndStatus lmEqtCode, 1, ""
            Else
                gCustomEndStatus lmEqtCode, 2, ""
            End If
        End If
        If udcCriteria.DGenType(3) = vbChecked Then
            'generate xref
            mExportRegionISCI
        End If
    End If
Cleanup:
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    If Not mRsRotations Is Nothing Then
        If (mRsRotations.State And adStateOpen) <> 0 Then
            mRsRotations.Close
        End If
        Set mRsRotations = Nothing
    End If
    If Not mRsCrfSplit Is Nothing Then
        If (mRsCrfSplit.State And adStateOpen) <> 0 Then
            mRsCrfSplit.Close
        End If
        Set mRsCrfSplit = Nothing
    End If
'    If Not rsBlackout Is Nothing Then
'        If (rsBlackout.State And adStateOpen) <> 0 Then
'            rsBlackout.Close
'        End If
'        Set rsBlackout = Nothing
'    End If
    'adv in facts log
    If udcCriteria.DGenType(2) = vbChecked Then
        Set myAdvDictionary = New Dictionary
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mExport"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    GoTo Cleanup
End Sub
Private Function mBuildRecordset(llGenericIndex As Long) As Boolean
    Dim llSplitIndex As Long
    
On Error GoTo ERRORBOX
    mRsRotations.Filter = adFilterNone
    If mRsRotations.RecordCount > 0 Then
        Do While Not mRsRotations.EOF
            mRsRotations.Delete
            mRsRotations.MoveNext
        Loop
    End If
    llSplitIndex = tmIDCGeneric(llGenericIndex).lFirstSplit
    Do While llSplitIndex <> -1
        With tmIDCSplit(llSplitIndex)
'        'test only!
'        If .iRotation = 53 Then
'            .iRotation = 40
'        End If
'          mRsRotations.AddNew Array("SplitIndex", "Start", "End", "Region", "Rotation", "Found", "StartTime", "EndTime"), Array(llSplitIndex, .sCrfStartDate, .sCrfEndDate, .lRafCode, .iRotation, False, .sStartTime, .sEndTime)
'           llSplitIndex = .lNextSplit
            mRsRotations.AddNew Array("SplitIndex", "Start", "End", "Region", "Rotation", "Found", "StartTime", "EndTime"), Array(llSplitIndex, .sCrfStartDate, .sCrfEndDate, .lRafCode, .iRotation, False, .sStartTime, .sEndTime)
            '6514
            mRsCrfSplit.Filter = "SplitIndex = " & llSplitIndex
            Do While Not mRsCrfSplit.EOF
                mRsRotations.AddNew Array("SplitIndex", "Start", "End", "Region", "Rotation", "Found", "StartTime", "EndTime"), Array(mRsCrfSplit!SPLITINDEX, mRsCrfSplit!Start, mRsCrfSplit!End, mRsCrfSplit!Region, mRsCrfSplit!Rotation, False, mRsCrfSplit!startTime, mRsCrfSplit!endTime)
                mRsCrfSplit.MoveNext
            Loop
            llSplitIndex = .lNextSplit
        End With
    Loop
    mBuildRecordset = True
    Exit Function
ERRORBOX:
    myErrors.WriteError "Error in mBuildRecordset: " & Err.Description
    mBuildRecordset = False
End Function
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Start", adDate
            .Append "End", adDate
            .Append "StartTime", adChar, 10
            .Append "EndTime", adChar, 10
            .Append "SplitIndex", adInteger
            .Append "Region", adInteger
            .Append "Rotation", adInteger
            .Append "Found", adBoolean
        End With
    myRs.Open
    myRs!Rotation.Properties("optimize") = True
    myRs.Sort = "Rotation desc"
    Set mPrepRecordset = myRs
End Function
Private Function mPrepRecordsetGroup() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "AttCode", adInteger
            .Append "SiteId", adChar, 5
            .Append "Call", adChar, 10
        End With
    myRs.Open
    myRs!attCode.Properties("optimize") = True
    Set mPrepRecordsetGroup = myRs
End Function
Private Function mPrepRecordsetBlackout() As ADODB.Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Index", adInteger
            .Append "SiteId", adChar, 5
            .Append "StartDate", adDate
            .Append "EndDate", adDate
            .Append "Generic", adChar, 15
            .Append "Rotation", adInteger
        End With
    myRs.Open
    myRs!Rotation.Properties("optimize") = True
    Set mPrepRecordsetBlackout = myRs
End Function
Private Function mCleanIef() As Boolean
    Dim slPastDate As String
    Dim llPastDate As Long
    
    If igExportSource = 2 Then DoEvents
    slPastDate = DateAdd("d", -2, smNowDate)
    llPastDate = gDateValue(slPastDate)
    '10/24/13 ddf changes only v60 Dan
    SQLQuery = "DELETE ief_IDC_Enforced WHERE  iefScheduleDate < " & llPastDate
   ' SQLQuery = "DELETE ief_IDC_Enforced WHERE  iefGenericCifCode < " & llPastDate
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, "Export IDC-mCleanIef"
        mCleanIef = False
        Exit Function
    End If
    mCleanIef = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mCleanIef"
    mCleanIef = False
End Function

Private Function mOpenCSVFile(slCSVFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenCSVFileErr:
    ilRet = 0
    slToFile = smExportDirectory & "IDC_" & Format(smDate, "YYYYMMDD") & "-" & Format(smEndDate, "YYYYMMDD") & ".csv"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo mOpenCSVFileErr:
    'hmCSV = FreeFile
    'Open slToFile For Output As hmCSV
    ilRet = gFileOpen(slToFile, "Output", hmCSV)
    If ilRet <> 0 Then
        Close hmCSV
        hmCSV = -1
        gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenCSVFile = False
        Exit Function
    End If
    On Error GoTo 0
    slCSVFileName = slToFile
    mOpenCSVFile = True
    Exit Function
'mOpenCSVFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Function mSendtoIdc() As Boolean
    Dim slErrorMessage As String
    Dim llLoop As Long
    Dim ilCounter As Integer
    Dim myIdc As CIdc
    Dim slUrl As String
    Dim blIsDeleteNeeded As Boolean
    Dim blKeepErrorInFunction As Boolean
    Dim slLogInfo As String
    ' for gUpdateLastExportDate- doesn't exist in v58. also used for messages to user
    Dim blAtLeastOne As Boolean
   'testing
    Dim slTestReceivers As String
    'scheduling
    Dim dlScheduleDate As Date
    'for messages
    Dim blIsConnectionError As Boolean
    Dim blIsDeletionError As Boolean
    Dim blIsTriggerError As Boolean
    Dim blIsStationError As Boolean
    Dim llTrigger As Long
    Dim llSchedule As Long
    Dim slDate As String
    Dim blOk As Boolean
    Dim ilMidnight As Integer
    Dim slTrigger As String
    'for trigger by week
    Dim dlStart As Date
    Dim dlEnd As Date
    Dim dlScheduleStart As Date
    Dim dlScheduleEnd As Date
    Dim ilDays As Integer
    Dim c As Integer
    Dim lRsDate As ADODB.Recordset
    '6419
    Dim lRsGroup As ADODB.Recordset
    If igExportSource = 2 Then DoEvents
    'reset if they previously pressed to cancel
    imTerminate = False
    mSendtoIdc = True
    blIsConnectionError = False
    blIsDeletionError = False
    blIsTriggerError = False
    blIsStationError = False
    slErrorMessage = ""
    lbcSort.Clear
    slUrl = mGetURL()
    If Len(slUrl) = 0 Then
        mSendtoIdc = False
        blIsConnectionError = True
        slErrorMessage = "  Could not read values from xml.ini."
        myErrors.WriteWarning slErrorMessage, True
        GoTo Cleanup
    End If 'url exists
    slLogInfo = "Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days."
    'send to test server, so only some values allowed. so Jason can test.
    If mnuCsiTest.Checked = True Then
        bmCsiTest = True
    Else
        bmCsiTest = False
    End If
    Set myIdc = New CIdc
    'send to idc. If not sending, don't create xml log file and put into test mode to stop the sending
    If udcCriteria.DGenType(0) = vbChecked Then
        myIdc.LogPath = myIdc.CreateLogName(sgMsgDirectory & FILEDEBUG)
    Else
        myIdc.isTest = True
    End If
    ' debug mode
    If UCase(slUrl) = "TEST" Then
        'isTest: for testing only.  don't delete, don't write to idc, don't update ief.
        myIdc.isTest = True
        mSetResults "Running in 'Test Mode'", MESSAGEBLACK
    End If
    'sending to idc
    If udcCriteria.DGenType(0) = vbChecked Then
        myErrors.WriteFacts "IDC Export." & slLogInfo, True
    End If
    lacResult.Caption = "Attempting to connect to IDC"
    DoEvents
    myIdc.SoapUrl = slUrl
    myIdc.LogStart
    If Len(myIdc.SoapUrl) = 0 Then
        mSendtoIdc = False
        blIsConnectionError = True
        slErrorMessage = myIdc.ErrorMessage
        myErrors.WriteWarning slErrorMessage
        GoTo Cleanup
    End If
    'url didn't work, or IDC not operating. Log start is here.
    If Not myIdc.IsConnected() Then
        mSendtoIdc = False
        blIsConnectionError = True
        slErrorMessage = "Could not connect to IDC."
        myErrors.WriteWarning slErrorMessage
        GoTo Cleanup
    End If 'connected
    '4/25/13 add redundancy state
    If Not myIdc.IsMasterState Then
        myErrors.WriteWarning "IDC is in Backup Redundancy State."
        mSetResults "IDC is in Backup Redundancy State", MESSAGEBLACK
        myIdc.SoapUrl = mGetBackupUrl()
        If Len(myIdc.SoapUrl) = 0 Then
            mSendtoIdc = False
            blIsConnectionError = True
            slErrorMessage = myIdc.ErrorMessage
            myErrors.WriteWarning slErrorMessage
            GoTo Cleanup
        ElseIf Not myIdc.IsConnected() Then
            mSendtoIdc = False
            blIsConnectionError = True
            slErrorMessage = "Could not connect to IDC."
            myErrors.WriteWarning slErrorMessage
            GoTo Cleanup
        End If
    End If
    Set myFacts = New CLogger
    If udcCriteria.DGenType(2) = vbChecked Then
        With myFacts
            .LogPath = .CreateLogName(smExportDirectory & FILEFACTS)
            If myIdc.isTest Then
                .WriteFacts "Testing only " & slLogInfo, True
            Else
                .WriteFacts "Sending to Idc " & slLogInfo, True
            End If
        End With
    End If
On Error GoTo 0
    If UBound(tmIDCGeneric) = 0 Then
        mSendtoIdc = True
        'sending to idc
        If udcCriteria.DGenType(0) = vbChecked Then
            myErrors.WriteFacts "There is no regional copy to send."
        End If
        myFacts.WriteFacts "There is no regional copy to send."
        mSetResults "There is no regional copy to send.", MESSAGEBLACK
        GoTo Cleanup
    End If
    Set mRsRotations = mPrepRecordset()
    Set lRsDate = mPrepRecordsetDate()
    Set lRsGroup = mPrepRecordsetGroup()
    For llLoop = 0 To UBound(tmIDCGeneric) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        myIdc.GenericISCI = gXMLNameFilter(tmIDCGeneric(llLoop).sISCI)
        myIdc.Notes = gXMLNameFilter(mGetCSFComment(tmIDCGeneric(llLoop).lCrfCsfCode)) 'slComment
       'get all splits for this generic
       ' made into function after corrupt database caused crash 9/12/13 Dan
        If mBuildRecordset(llLoop) Then
             mBuildRsDate lRsDate  ', lRsGroup
            If lRsDate.RecordCount > 0 Then
                lRsDate.MoveFirst
                dlStart = lRsDate!Date
                lRsDate.MoveNext
                'each pair of days gets a trigger for generic.  Find all regional that are for these dates
                Do While Not lRsDate.EOF
                    dlEnd = lRsDate!Date
                    myIdc.TriggerDate = dlStart
                    myIdc.NumberDays = DateDiff("d", dlStart, dlEnd)
                    mRsRotations.Filter = "End >=" & dlStart & " AND Start <" & dlEnd & " AND Found = False "
                    Do While Not mRsRotations.EOF
                        If igExportSource = 2 Then DoEvents
                        ilMidnight = mIsMidnight(mRsRotations!startTime, mRsRotations!endTime)
                        ' = midnight; regular rotation
                        If ilMidnight = 1 Then
                            '6175
                            If tmIDCSplit(mRsRotations!SPLITINDEX).iRotation > 0 Then
                                 myIdc.RegionalInfo = mGetRegionInfo(mRsRotations!SPLITINDEX, myIdc, lRsGroup)
                            End If
                        ElseIf ilMidnight = 0 Then
                            '6175
                            If tmIDCSplit(mRsRotations!SPLITINDEX).iRotation > 0 Then
                                myIdc.RegionalInfoTime = mGetRegionInfo(mRsRotations!SPLITINDEX, myIdc, lRsGroup)
                            End If
                        Else
                            'error!
                        End If
                        mRsRotations.MoveNext
                    Loop    ' set of regions for trigger
                    'gathered rotations.  Now write them out.
                    If myIdc.isAtLeastOneRegional() Or myIdc.isAtLeastOneRegionalTime() Then
                        If mCreateTriggers(myIdc) Then
                            blAtLeastOne = True
                        Else
                            blIsTriggerError = True
                        End If
                        myIdc.Clear True
                    End If
                    dlStart = dlEnd
                    lRsDate.MoveNext
                Loop
                End If 'set splits ok
           ' Next ilCounter 'for day
            myIdc.Clear False
        Else
            mSendtoIdc = False
        End If
    Next llLoop 'generic
    'Deletes moved here
On Error GoTo ERRCANTDELETE
    blIsDeleteNeeded = mIsDeleteNeeded()
    If blIsDeleteNeeded Then
        If mDeleteSchedules(myIdc) Then
            If Not myIdc.isTest Then
                lacResult.Caption = "Previous schedules removed from IDC"
            End If
        Else
            mSendtoIdc = False
            blIsDeletionError = True
        End If
    End If
        '6508 rare case of no triggers being sent because of missing files.
    If myIdc.IsNoTriggers() Then
            blIsTriggerError = True
            myErrors.WriteError "No Triggers were created."
            myFacts.WriteWarning "No Triggers were created."
    Else
    '    '6247. now write out schedules. One per day.
        If Not myIdc.CreateSchedule() Then
            blIsTriggerError = True
            myErrors.WriteError "Couldn't create schedule: " & myIdc.ErrorMessage
            myFacts.WriteWarning "Couldn't create schedule: " & myIdc.ErrorMessage
        ElseIf Not myIdc.isTest Then
            blOk = myIdc.GetScheduleInfo(True, llTrigger, llSchedule, slDate)
            myFacts.WriteFacts "List of triggers and schedules created"
            Do While blOk
                mUpdateIEF llTrigger, slDate, llSchedule
                myFacts.WriteFacts "        " & slDate & ": Trigger #: " & llTrigger & "   Schedule #: " & llSchedule
                blOk = myIdc.GetScheduleInfo(False, llTrigger, llSchedule, slDate)
            Loop
            myFacts.WriteFacts myIdc.ScheduleInfoMore
        Else
            myFacts.WriteFacts myIdc.ScheduleInfoMore
        End If
    End If
    'Dan M 7/18/12 add UpdateLastExportDate when transmitting. Must do by each vehicle, so simply get list of vehicles that could've been sent.
    If blAtLeastOne Then
        mLogEndDate
    End If
Cleanup:
    lacResult.Caption = ""
    If Not myIdc Is Nothing Then
        myIdc.LogEnd
    End If
    If udcCriteria.DGenType(2) = vbChecked Then
        If Not myFacts Is Nothing Then
            If myFacts.isLog Then
                If UBound(tmIDCGeneric) > 0 Then
                    mSetResults "Generated File", MESSAGEGREEN
                    lacResult.Caption = "Exports placed into: " & smExportDirectory
                End If
            Else
                mSetResults "Could not generate text file", MESSAGERED
            End If
        Else
                mSetResults "Could not generate text file", MESSAGERED
        End If
    End If
    Set myFacts = Nothing
    If blIsConnectionError Or blIsDeletionError Or blIsStationError Or blIsTriggerError Then
        'warnings written as they happen.  Letting user know there was an issue
        If blIsConnectionError Then
            mSetResults "Problems with Send to IDC: " & slErrorMessage, MESSAGERED
            mSetResults "Export Failed. See '" & FILEERROR & "' for issues.", MESSAGERED
        Else
            If blIsDeletionError Then
                mSetResults "Problems with deleting previous triggers.", MESSAGERED
            End If
            If blIsStationError Then
                mSetResults "Problems with stations: Some stations overlap in triggers.", MESSAGERED
            End If
            If blIsTriggerError And blAtLeastOne Then
                mSetResults "Some generics could not be sent. See '" & FILEERROR & "' for issues.", MESSAGERED
            ElseIf blIsTriggerError Then
                mSetResults "Export Failed. See '" & FILEERROR & "' for issues.", MESSAGERED
            ElseIf udcCriteria.DGenType(0) = vbChecked Then
                mSetResults "Sent to IDC. See '" & FILEERROR & "' for issues.", MESSAGERED
            End If
        End If
    'error in code
    ElseIf mSendtoIdc = False Then
        mSetResults "Problems with Send to IDC: " & slErrorMessage, MESSAGERED
        mSetResults "Export Failed. See '" & FILEERROR & "' for issues.", MESSAGERED
    ElseIf UBound(tmIDCGeneric) > 0 And udcCriteria.DGenType(0) = vbChecked Then
        mSetResults "Sent to IDC.", MESSAGEGREEN
    End If
    If Not myIdc Is Nothing And udcCriteria.DGenType(0) = vbChecked Then
        myErrors.WriteFacts "End send to IDC"
    End If
    Set myIdc = Nothing
    If Not lRsDate Is Nothing Then
        If (lRsDate.State And adStateOpen) <> 0 Then
            lRsDate.Close
        End If
        Set lRsDate = Nothing
    End If
    If Not lRsGroup Is Nothing Then
        If (lRsGroup.State And adStateOpen) <> 0 Then
            lRsGroup.Close
        End If
        Set lRsGroup = Nothing
    End If

    Exit Function
ERRCANTWRITE:
    mSendtoIdc = False
    blIsConnectionError = True
    slErrorMessage = "  Could not write to IDC server." & Err.Description
    GoTo Cleanup
    Exit Function
ERRCANTDELETE:
    mSendtoIdc = False
    blIsDeletionError = True
    Resume Next
End Function
Public Function mCreateTriggers(myIdc As CIdc) As Boolean
    Dim blRet As Boolean
    Dim slTriggers As String
    Dim llTrigger As Long
    Dim c As Integer
    Dim slTriggersArray() As String
On Error GoTo ERRORBOX

    blRet = True
    If myIdc.isAtLeastOneRegional() Then 'Or myIdc.isAtLeastOneRegionalTime Then
        llTrigger = myIdc.CreateTrigger()
        If llTrigger = 0 Then ' And Not myIdc.isTest Then
            blRet = False
            myErrors.WriteWarning "Couldn't create trigger for " & myIdc.TriggerDate & ":" & myIdc.ErrorMessage
        Else
            mWriteFacts llTrigger, myIdc
        End If
    End If
    'timed must always come after non timed!
    If myIdc.isAtLeastOneRegionalTime() Then
        slTriggers = myIdc.CreateTriggerTime()
        If Len(slTriggers) = 0 Then 'Or slTriggers = "0" Then  'And Not myIdc.isTest
            blRet = False
            myErrors.WriteWarning "Couldn't create time sensitive trigger :" & myIdc.ErrorMessage
        Else
            slTriggersArray = Split(slTriggers, myIdc.SplitMajor)
            For c = 0 To UBound(slTriggersArray)
                mWriteFacts CLng(slTriggersArray(c)), myIdc
            Next c
        End If
    End If
    mCreateTriggers = blRet
    Exit Function
ERRORBOX:
    mCreateTriggers = False
    myErrors.WriteError " mCreateTriggers: " & Err.Description
End Function
Public Function mGetRegionInfo(llSplitIndex As Long, myIdc As CIdc, lRsGroup As ADODB.Recordset) As CcsiToIdcRegionalInfo

    Dim myRegion As CcsiToIdcRegionalInfo
    Dim slRegionName As String
    Dim slRISCI As String
    Dim llISCIIndex As Long
    Dim slReceivers As String
    Dim llReceiverIndex As Long
    Dim slSafe As String
    Dim slTempIsci As String
    Dim slTempPercent As String
    Dim slRegionUnsafeStations As String
    Dim slStartDate As String
    Dim slEndDate As String
    
    slRegionUnsafeStations = ""
    Set myRegion = New CcsiToIdcRegionalInfo
    slRegionName = mGetRegionName(tmIDCSplit(llSplitIndex).lRafCode)
    slRISCI = ""
    llISCIIndex = tmIDCSplit(llSplitIndex).lFirstRISCI
'                'ttp 5287.
    Do While llISCIIndex <> -1 'each region's filenames
        If igExportSource = 2 Then DoEvents
        slTempIsci = gXMLNameFilter(tmISCIByPercent(llISCIIndex).sFilterRISCI)
        'dan 7/11/13 isci is uppercase; extension is lower case.
        slTempIsci = UCase(slTempIsci)
        slTempPercent = tmISCIByPercent(llISCIIndex).iPercent
        'in testing, got a slTempIsci that was blank.  Let it fall through so user can see the issue.
            If slRISCI = "" Then
                slRISCI = slTempIsci & ".mp3" & myRegion.SplitMinor & slTempPercent
            Else
                slRISCI = slRISCI & myRegion.SplitMajor & slTempIsci & ".mp3" & myRegion.SplitMinor & slTempPercent
            End If
        llISCIIndex = tmISCIByPercent(llISCIIndex).lNextRISCI
    Loop    'set of mp3 for region
    slReceivers = ""
    llReceiverIndex = tmIDCSplit(llSplitIndex).lFirstReceiver
    Do While llReceiverIndex <> -1
      '  If mBlackoutStationSafe(tmIDCReceiver(llReceiverIndex).sReceiverID, myIdc.GenericISCI, llSplitIndex) Then
            slReceivers = myIdc.BuildLine(slReceivers, tmIDCReceiver(llReceiverIndex).sCallLetters, tmIDCReceiver(llReceiverIndex).sReceiverID)
              '6419 group site ids
           ' slReceivers = mGroupSiteIds(llReceiverIndex, lRsGroup, myIdc.SplitMajor, myIdc.SplitMinor, slReceivers)
       ' End If
        llReceiverIndex = tmIDCReceiver(llReceiverIndex).lNextReceiver
    Loop    'set of stations for region
    'slSafe returns the edited sites.
    slRegionUnsafeStations = myIdc.IsSafeStation(slReceivers, slSafe)
   ' slSafe = mBuildSafeBlackoutStation(slSafe, llSplitIndex, myIdc.SplitMajor)
    'none safe! get out
    If Len(slSafe) > 2 Then
        slSafe = mLoseLastLetter(slSafe)
        slStartDate = Trim$(tmIDCSplit(llSplitIndex).sCrfStartDate)
        If DateDiff("d", slStartDate, smDate) > 0 Then
            slStartDate = smDate
        End If
        slEndDate = Trim$(tmIDCSplit(llSplitIndex).sCrfEndDate)
        If DateDiff("d", smEndDate, slEndDate) > 0 Then
            slEndDate = CDate(smEndDate)
        End If
        'sending to test server
        If bmCsiTest Then
            slRISCI = mTestReplaceFilename(slRISCI)
'            slStartDate = DateAdd("d", 7, gObtainNextMonday(Now))
'            slEndDate = DateAdd("d", 6, slStartDate)
        End If
        With myRegion
            .Definition = gXMLNameFilter(slRegionName)
            .startDate = slStartDate
            .endDate = slEndDate
            .startTime = Trim$(tmIDCSplit(llSplitIndex).sStartTime)
            .endTime = Trim$(tmIDCSplit(llSplitIndex).sEndTime)
            '"3S018B16.MP3" & .SplitMinor & "60" & .SplitMajor & "YQZN0273.MP3" & .SplitMinor & "40"
            .RegionalISCIs = slRISCI
            '"41" & .splitmajor & "42"
            .Sites = slSafe ' slReceivers
            'for facts only
            .Rotation = tmIDCSplit(llSplitIndex).iRotation
        End With
    End If
    Set mGetRegionInfo = myRegion
    'block message.  this is now common
    slRegionUnsafeStations = ""
    
End Function

Private Sub mWriteFacts(llTrigger As Long, myIdc As CIdc)
    Dim slTrigger As String
    Dim slLine As String
    Dim slAdv As String
    Dim ilAdv As Integer
    Dim ilAdfIndex As Integer
    
    ilAdv = 0
    slAdv = ""
    slLine = ""
    slTrigger = ""
    slTrigger = "Trigger #" & CStr(llTrigger)
    If Not myAdvDictionary Is Nothing Then
        If myAdvDictionary.Exists(myIdc.GenericISCI) Then
            ilAdv = myAdvDictionary.Item(myIdc.GenericISCI)
            For ilAdfIndex = 0 To UBound(tgAdvtInfo) - 1 Step 1
                If tgAdvtInfo(ilAdfIndex).iCode = ilAdv Then
                    slAdv = tgAdvtInfo(ilAdfIndex).sAdvtName
                    Exit For
                End If
            Next ilAdfIndex
            If Len(slAdv) > 0 Then
                slAdv = " Advertiser: " & slAdv
            End If
        End If
    End If
    slLine = " Trigger for " & myIdc.TriggerDate & ", " & myIdc.NumberDays & " days " & slTrigger & vbCrLf & "Generic:" & myIdc.GenericISCI & slAdv & vbCrLf & myIdc.WriteRegionalInfo(llTrigger)
    myFacts.WriteFacts (slLine)
End Sub
Private Function mIsMidnight(slTime As String, slTime2 As String) As Integer
    ' -1 is error, 0 is not midnight, 1 is midnight
    Dim ilRet As Integer
    
On Error GoTo ERRORBOX
'    slTime = Trim$(slTime)
'    slTime2 = Trim$(slTime2)
    If IsDate(slTime) And IsDate(slTime2) Then
        If DateDiff("s", "00:00:00", slTime) <> 0 Then
            ilRet = 0
        ElseIf DateDiff("s", "00:00:00", slTime2) <> 0 Then
            ilRet = 0
        Else
            ilRet = 1
        End If
    Else
        ilRet = -1
    End If
    mIsMidnight = ilRet
    Exit Function
ERRORBOX:
    mIsMidnight = -1
End Function
Private Sub mLogEndDate()
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim slEDate As String

    For ilVef = 0 To lbcVehicles.ListCount - 1
        ilVefCode = lbcVehicles.ItemData(ilVef)
        slEDate = mGetEarliestEndDate(smDate, imNumberDays)
        gUpdateLastExportDate ilVefCode, slEDate
    Next ilVef

End Sub
Private Function mGetEarliestEndDate(ByVal slStartDate As String, ilNumberDays As Integer) As String
    Dim slEndChosen As String
    Dim slEndOfWeek As String
    
    slEndChosen = DateAdd("d", ilNumberDays - 1, slStartDate)
    slEndOfWeek = gObtainNextSunday(slStartDate)
    If gDateValue(gAdjYear(slEndChosen)) < gDateValue(gAdjYear(slEndOfWeek)) Then
         mGetEarliestEndDate = slEndChosen
    Else
        mGetEarliestEndDate = slEndOfWeek
    End If
End Function
Private Function mTestReplaceFilename(slFileString As String) As String
    Dim slNewString As String
    Dim ilCount As Integer
    Dim slArray() As String
    Dim slfake(9) As String
    Dim slNewWord As String
    Dim ilPos As Integer
    
    slfake(0) = "File2"
    slfake(1) = "File1"
    slfake(2) = "File3"
    slfake(3) = "File4"
    slfake(4) = "File5"
    slfake(5) = "File1"
    slfake(6) = "File2"
    slfake(7) = "File3"
    slfake(8) = "File4"
    slfake(9) = "File5"
    slNewString = slFileString
    slArray = Split(slFileString, ";")
    For ilCount = 0 To UBound(slArray)
        ilPos = InStr(1, slArray(ilCount), ".mp3")
        If ilPos > 0 Then
            If ilCount < 10 Then
                slNewWord = slfake(ilCount) & Mid(slArray(ilCount), ilPos)
            Else
                slNewWord = slfake(ilCount - 9) & Mid(slArray(ilCount), ilPos)
            End If
        End If
        slNewString = Replace(slNewString, slArray(ilCount), slNewWord)
    Next ilCount
    mTestReplaceFilename = slNewString
End Function
Private Function mLoadFromIni(Section As String, Key As String, slPath As String, sValue As String) As Boolean
    'generic. from any ini with slpath pointing to ini file.
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    
    sValue = "Not Found"
    mLoadFromIni = False
    If Dir(slPath) > "" Then
        BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slPath)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                sValue = Left(sBuffer, BytesCopied)
                mLoadFromIni = True
            End If
        End If
    End If 'slPath not valid?
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function
Private Function mOpenCSF() As Integer

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
Private Function mCloseCSF() As Integer

    Dim ilRet As Integer
    
    ilRet = btrClose(hmCsf)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrClose Failed on CSF.BTR"
        btrDestroy hmCsf
        mCloseCSF = False
        Exit Function
    End If
    
    btrDestroy hmCsf
    mCloseCSF = True
    Exit Function

End Function
Private Function mGetCSFComment(lCSFCode As Long) As String

    Dim ilRet, i, ilLen, ilActualLen As Integer
    Dim ilRecLen As Integer
    Dim tlCSF As CSF
    Dim tlCsfSrchKey As LONGKEY0
    Dim slComment As String
    Dim slTemp As String
    Dim blOneChar As Byte
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    mGetCSFComment = ""
    If (lCSFCode <= 0) Then
        Exit Function
    End If
    tlCsfSrchKey.lCode = lCSFCode
    tlCSF.sComment = ""
    ilRecLen = Len(tlCSF) '5011
    ilRet = btrGetEqual(hmCsf, tlCSF, ilRecLen, tlCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If

    slComment = gStripChr0(tlCSF.sComment)
    If slComment <> "" Then
        ' Strip off any trailing non ascii characters.
        ilLen = Len(slComment)
        ' Find the first valid ascii character from the end and assume the rest of the string is good.
        For i = ilLen To 1 Step -1
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
            blOneChar = Asc(Mid(slTemp, i, 1))
            If blOneChar >= 32 Then
                slComment = slComment + Mid(slTemp, i, 1)
            Else
                slComment = slComment + " "
            End If
        Next i
        mGetCSFComment = slComment
    End If
    If igExportSource = 2 Then DoEvents
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in modPervasive-mGetCSFComment: "
        myErrors.WriteError gMsg & Err.Description & "; Error #" & Err.Number, True, True
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Function

Private Function mGetRegionName(llRafCode As Long) As String
    On Error GoTo ErrHand
    
    If igExportSource = 2 Then DoEvents
    SQLQuery = "Select rafName FROM RAF_Region_Area"
    SQLQuery = SQLQuery & " Where (rafCode = " & llRafCode & ")"
    Set raf_rst = gSQLSelectCall(SQLQuery)
    If raf_rst.EOF Then
        mGetRegionName = ""
    Else
        mGetRegionName = Trim$(raf_rst!rafName)
    End If
    If igExportSource = 2 Then DoEvents
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mGetRegionName"
    mGetRegionName = ""
    Exit Function

End Function

Private Sub txtNumberDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Function mExportRegionISCI() As Boolean
    '5437
    Dim llLoop As Long
    Dim llSplitIndex As Long
    Dim slAdvName As String
    Dim llISCIIndex As Long
    Dim slRISCI As String
    Dim slUnfilteredISCI As String
    Dim slString As String
    
    lbcSort.Clear
    For llLoop = 0 To UBound(tmIDCGeneric) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llSplitIndex = tmIDCGeneric(llLoop).lFirstSplit
        Do While llSplitIndex <> -1
            If igExportSource = 2 Then DoEvents
                slAdvName = Trim$(tmIDCSplit(llSplitIndex).sAdvName)
                llISCIIndex = tmIDCSplit(llSplitIndex).lFirstRISCI
                Do While llISCIIndex <> -1
                    If igExportSource = 2 Then DoEvents
                    slRISCI = Trim$(tmISCIByPercent(llISCIIndex).sFilterRISCI) & ".mp3"
                    slUnfilteredISCI = Trim$(tmISCIByPercent(llISCIIndex).sUnfilterRISCI)
                    slString = """" & slAdvName & """" & ", " & """" & slUnfilteredISCI & """" & "," & slRISCI & ", " & Trim$(tmIDCSplit(llSplitIndex).sCrfStartDate) & "," & Trim$(tmIDCSplit(llSplitIndex).sCrfEndDate) & ";"
                    'only write out once.
                    If SendMessageByString(lbcSort.hwnd, LB_FINDSTRING, -1, slString) < 0 Then
                        lbcSort.AddItem slString
                    End If
                    llISCIIndex = tmISCIByPercent(llISCIIndex).lNextRISCI
                Loop
            llSplitIndex = tmIDCSplit(llSplitIndex).lNextSplit
        Loop
    Next llLoop
    If igExportSource = 2 Then DoEvents
     mExportRegionISCI = mPrintListBox(lbcSort)
End Function
Private Function mPrintListBox(lbcList As ListBox) As Boolean
    Dim blRet As Boolean
    Dim myAudio As CLogger
    Dim ilLoop As Integer
    
    blRet = False
    Set myAudio = New CLogger
    If lbcList.ListCount > 0 Then
        With myAudio
            .LogPath = .CreateLogName(smExportDirectory & "IDC_Region_ISCI_")
            If .isLog Then
                .WriteFacts "For Export date of: " & smDate & " and " & CStr(imNumberDays) & " Days.", True
                For ilLoop = 0 To lbcList.ListCount - 1
                    If igExportSource = 2 Then DoEvents
                    .WriteFacts lbcList.List(ilLoop)
                    blRet = True
                    lacResult.Caption = "Exports placed into: " & smExportDirectory
                    mSetResults "Exported Audio List", MESSAGEGREEN
                Next ilLoop
                blRet = True
            Else
                mSetResults "Error generating Audio List.", MESSAGERED
            End If
        End With
    Else
        mSetResults "Did not export Audio List. There was no data to send.", MESSAGEBLACK
    End If
Cleanup:
    Set myAudio = Nothing
    mPrintListBox = blRet
End Function
Private Sub mGetGenericISCIForBlackouts(tlAstInfo As ASTINFO)
    Dim llLstCode As Long
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    SQLQuery = "SELECT lstBkoutLstCode"
    SQLQuery = SQLQuery & " FROM lst"
    SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If igExportSource = 2 Then DoEvents
        llLstCode = rst!lstBkoutLstCode
        If llLstCode > 0 Then
            SQLQuery = "SELECT lstCifCode, lstCrfCsfCode, lstISCI, lstSdfCode"
            SQLQuery = SQLQuery & " FROM lst"
            SQLQuery = SQLQuery & " WHERE lstCode =" & Str(llLstCode)
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                tlAstInfo.lCifCode = rst!lstCifCode
                tlAstInfo.lCrfCsfCode = rst!lstCrfCsfCode
                tlAstInfo.sISCI = rst!lstISCI
                tlAstInfo.lSdfCode = rst!lstSdfCode
            End If
        End If
    End If
    If igExportSource = 2 Then DoEvents
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mGetGenericISCIForBlackouts"
    Exit Sub
End Sub

Private Sub mClearAlerts()
    Dim ilVef As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim llStartDate As Long
    Dim ilRet As Integer
    
    Dim sEndDate As String
    Dim slSDate As String
    Dim slEDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    smEndDate = sEndDate
    slSDate = smDate
    slEDate = gObtainNextSunday(slSDate)
    If gDateValue(gAdjYear(sEndDate)) < gDateValue(gAdjYear(slEDate)) Then
        slEDate = sEndDate
    End If
    llSDate = gDateValue(smDate)
    llEDate = gDateValue(slEDate)
    
    slDate = gObtainPrevMonday(Format(llSDate, "m/d/yy"))
    llStartDate = gDateValue(slDate)
    For ilVef = 0 To lbcVehicles.ListCount - 1
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
Private Sub mBuildSplitISCI(llSplitIndex As Long)
    Dim blFound As Boolean
    Dim ilLoop As Integer
    Dim ilTotalCount As Integer
    Dim ilTotalPercent As Integer
    Dim llCrfCode As Long
    Dim llCpfCode As Long
    Dim llIndexRISCI As Long
    Dim ilRet As Integer
    Dim slRCartNo As String
    Dim slRProduct As String
    Dim slRISCI As String
    Dim slRCreativeTitle As String
    Dim llRCrfCsfCode As Long
    Dim llRCpfCode As Long
    Dim ilCifAdfCode As Integer
    '6504
   ' Dim slTempDate As String
    Dim slStart As String
    Dim slEnd As String
    
    ReDim tmISCICount(0 To 0) As ISCICOUNT
    
    If igExportSource = 2 Then DoEvents
    llCrfCode = tmIDCSplit(llSplitIndex).lRCrfCode
    tmIDCSplit(llSplitIndex).sCrfStartDate = ""
    tmIDCSplit(llSplitIndex).sCrfEndDate = ""
    SQLQuery = "Select * FROM crf_Copy_Rot_Header"
    SQLQuery = SQLQuery & " Where (crfCode = " & llCrfCode & ")"
    Set crf_rst = gSQLSelectCall(SQLQuery)
    If Not crf_rst.EOF Then
'        '6504
'        slTempDate = mCrfDateModified(crf_rst!crfStartDate, True)
'        'tmIDCSplit(llSplitIndex).sCrfStartDate = Format(crf_rst!crfStartDate, sgShowDateForm)
'        tmIDCSplit(llSplitIndex).sCrfStartDate = Format(slTempDate, sgShowDateForm)
'        slTempDate = mCrfDateModified(crf_rst!crfEndDate, False)
'        'tmIDCSplit(llSplitIndex).sCrfEndDate = Format(crf_rst!crfEndDate, sgShowDateForm)
'        tmIDCSplit(llSplitIndex).sCrfEndDate = Format(slTempDate, sgShowDateForm)
        '5882
        tmIDCSplit(llSplitIndex).iRotation = crf_rst!crfRotNo
        tmIDCSplit(llSplitIndex).sStartTime = Format$(crf_rst!crfStartTime, XMLTIME)
        tmIDCSplit(llSplitIndex).sEndTime = Format$(crf_rst!crfEndTime, XMLTIME)
        '6504
        'modify these dates as needed.  Also, build additional
        slStart = crf_rst!crfStartDate
        slEnd = crf_rst!crfEndDate
        mCrfDateModified slStart, slEnd, llSplitIndex
        tmIDCSplit(llSplitIndex).sCrfStartDate = Format(slStart, sgShowDateForm)
        tmIDCSplit(llSplitIndex).sCrfEndDate = Format(slEnd, sgShowDateForm)

        '6419 blackouts.
'        rsBlackout.Filter = "Index = " & llSplitIndex
'        Do While Not rsBlackout.EOF
'            rsBlackout!startDate = Trim(tmIDCSplit(llSplitIndex).sCrfStartDate)
'            rsBlackout!endDate = Trim(tmIDCSplit(llSplitIndex).sCrfEndDate)
'            rsBlackout!Rotation = tmIDCSplit(llSplitIndex).iRotation
'            rsBlackout.MoveNext
'        Loop
        'Gather ISCI
        SQLQuery = "Select * FROM cnf_Copy_Instruction"
        SQLQuery = SQLQuery & " Where (cnfCrfCode = " & llCrfCode & ")"
        Set cnf_rst = gSQLSelectCall(SQLQuery)
        Do While Not cnf_rst.EOF
            If igExportSource = 2 Then DoEvents
            blFound = False
            'ex: split of 2 regional copy (with names of '1' and '2') currently looks like this: 1,2,1,2,1.  Join as needed to make 1 60% and 2 40%
            For ilLoop = 0 To UBound(tmISCICount) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmISCICount(ilLoop).lCifCode = cnf_rst!cnfCifCode Then
                    blFound = True
                    tmISCICount(ilLoop).iCount = tmISCICount(ilLoop).iCount + 1
                    Exit For
                End If
            Next ilLoop
            If Not blFound Then
                tmISCICount(UBound(tmISCICount)).lCifCode = cnf_rst!cnfCifCode
                tmISCICount(UBound(tmISCICount)).iCount = 1
                ReDim Preserve tmISCICount(0 To UBound(tmISCICount) + 1) As ISCICOUNT
            End If
            cnf_rst.MoveNext
        Loop
        If UBound(tmISCICount) > 0 Then
            ilTotalCount = 0
            For ilLoop = 0 To UBound(tmISCICount) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                ilTotalCount = ilTotalCount + tmISCICount(ilLoop).iCount
            Next ilLoop
            ilTotalPercent = 0
            If ilTotalCount > 1 Then
                For ilLoop = 0 To UBound(tmISCICount) - 2 Step 1
                    If igExportSource = 2 Then DoEvents
                    tmISCICount(ilLoop).iPercent = (100 * tmISCICount(ilLoop).iCount) / ilTotalCount
                    ilTotalPercent = ilTotalPercent + tmISCICount(ilLoop).iPercent
                Next ilLoop
            End If
            tmISCICount(UBound(tmISCICount) - 1).iPercent = 100 - ilTotalPercent
        End If
        For ilLoop = 0 To UBound(tmISCICount) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            ilRet = gGetCopy("1", tmISCICount(ilLoop).lCifCode, 0, True, slRCartNo, slRProduct, slRISCI, slRCreativeTitle, llRCrfCsfCode, llRCpfCode, ilCifAdfCode)
            If ilRet Then
                llIndexRISCI = UBound(tmISCIByPercent)
                tmISCIByPercent(llIndexRISCI).iPercent = tmISCICount(ilLoop).iPercent
                tmISCIByPercent(llIndexRISCI).sUnfilterRISCI = Trim$(slRISCI)
                tmISCIByPercent(llIndexRISCI).sFilterRISCI = mFileNameFilter(Trim$(slRISCI))
                tmISCIByPercent(llIndexRISCI).lNextRISCI = tmIDCSplit(llSplitIndex).lFirstRISCI
                tmIDCSplit(llSplitIndex).lFirstRISCI = llIndexRISCI
                ReDim Preserve tmISCIByPercent(0 To UBound(tmISCIByPercent) + 1) As ISCIBYPERCENT
            End If
        Next ilLoop
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mBuildSplitISCI"
    Exit Sub
    
End Sub

Private Sub mBuildGenericCif(llCrfCode As Long)
    Dim blFound As Boolean
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slCartNo As String
    Dim slProduct As String
    Dim slISCI As String
    Dim slCreativeTitle As String
    Dim llCrfCsfCode As Long
    Dim llCpfCode As Long
    Dim ilCifAdfCode As Integer
    
    ReDim tmCifCode(0 To 0) As GENERICCIF
    If igExportSource = 2 Then DoEvents
    SQLQuery = "Select * FROM crf_Copy_Rot_Header"
    SQLQuery = SQLQuery & " Where (crfCode = " & llCrfCode & ")"
    Set crf_rst = gSQLSelectCall(SQLQuery)
    If Not crf_rst.EOF Then
        If igExportSource = 2 Then DoEvents
        'Gather ISCI
        SQLQuery = "Select * FROM cnf_Copy_Instruction"
        SQLQuery = SQLQuery & " Where (cnfCrfCode = " & llCrfCode & ")"
        Set cnf_rst = gSQLSelectCall(SQLQuery)
        Do While Not cnf_rst.EOF
            If igExportSource = 2 Then DoEvents
            blFound = False
            For ilLoop = 0 To UBound(tmCifCode) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tmCifCode(ilLoop).lCifCode = cnf_rst!cnfCifCode Then
                    blFound = True
                    Exit For
                End If
            Next ilLoop
            If Not blFound Then
                tmCifCode(UBound(tmCifCode)).lCifCode = cnf_rst!cnfCifCode
                ReDim Preserve tmCifCode(0 To UBound(tmCifCode) + 1) As GENERICCIF
            End If
            cnf_rst.MoveNext
        Loop
        For ilLoop = 0 To UBound(tmCifCode) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            ilRet = gGetCopy("1", tmCifCode(ilLoop).lCifCode, 0, True, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode, ilCifAdfCode)
            If ilRet Then
                tmCifCode(ilLoop).sISCI = slISCI
            End If
        Next ilLoop
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mBuildGenericCif"
    Exit Sub
    
End Sub

Private Function mGetRegionISCI() As Integer
    Dim llLoop As Long
    Dim llSplitIndex As Long
    
    mGetRegionISCI = True
    For llLoop = 0 To UBound(tmIDCGeneric) - 1 Step 1
        llSplitIndex = tmIDCGeneric(llLoop).lFirstSplit
        Do While llSplitIndex <> -1
            mBuildSplitISCI llSplitIndex
           llSplitIndex = tmIDCSplit(llSplitIndex).lNextSplit
        Loop
    Next llLoop
End Function

Private Function mUpdateIEF(llTriggerId As Long, slDate As String, llSchedule As Long) As Boolean
    Dim llDate As Long
    '6232 added schedule
    On Error GoTo ErrHand
    llDate = gDateValue(slDate)
    If llTriggerId > 0 And llSchedule > 0 Then
        If igExportSource = 2 Then DoEvents
        '10/24/13 ddf change only v60
        SQLQuery = "Insert Into ief_IDC_Enforced (iefTriggerID,iefScheduleDate, iefScheduleID) values ( " & llTriggerId & "," & llDate & "," & llSchedule & ")"
        'SQLQuery = "Insert Into ief_IDC_Enforced (iefTriggerID,iefGenericCifCode, iefSplitRafCode) values ( " & llTriggerId & "," & llDate & "," & llSchedule & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, "Export IDC-mUpdateIef"
            mUpdateIEF = False
            Exit Function
        Else
            mUpdateIEF = True
        End If
   End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mUpdateIEF"
    mUpdateIEF = False
    Exit Function
End Function
Private Function mIsDeleteNeeded() As Boolean
    Dim blRet As Boolean
    Dim slDates As String
    Dim slPastDate As String
    Dim llPastDate As Long
    Dim ilCounter
    
    If igExportSource = 2 Then DoEvents
    blRet = False
    'this is here to 'prefill'.  It will be overridden in the next step
    slPastDate = DateAdd("d", -2, smNowDate)
    For ilCounter = 0 To imNumberDays - 1 Step 1
        slPastDate = DateAdd("d", ilCounter, smDate)
        llPastDate = gDateValue(slPastDate)
        slDates = slDates & CStr(llPastDate) & " ,"
    Next ilCounter
    slDates = mLoseLastLetter(slDates)
    '6446
    '10/24/13 ddf change for v60 only!
    SQLQuery = "Select iefScheduleID as Schedule FROM ief_IDC_Enforced where iefScheduleDate in (" & slDates & ") " 'order by Schedule"
  '  SQLQuery = "Select iefSplitRafCode as Schedule FROM ief_IDC_Enforced where iefGenericCifCode in (" & slDates & ") " 'order by Schedule"
     On Error GoTo ErrHand
    Set ief_rst = gSQLSelectCall(SQLQuery)
    With ief_rst
        If igExportSource = 2 Then DoEvents
        If Not (.EOF Or .BOF) Then
            blRet = True
        End If
    End With
    mIsDeleteNeeded = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mIsDeleteNeeded"
    mIsDeleteNeeded = False
    Err.Raise 55555
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
Private Function mDeleteSchedules(myIdc As CIdc) As Integer
    Dim slDeletes As String
    Dim slSchedules As String
    Dim llPreviousSchedule As Long
    Dim blDeleteOk As Boolean
    Dim slFail As String
    
    blDeleteOk = True
    slSchedules = ""
    slFail = ""
    llPreviousSchedule = 0
    slDeletes = ""
    mDeleteSchedules = True
    If Not ief_rst Is Nothing Then
        Do While Not ief_rst.EOF
            If igExportSource = 2 Then DoEvents
                '6232
 On Error GoTo ErrHand
            If ief_rst!Schedule > 0 Then
                If llPreviousSchedule <> ief_rst!Schedule Then
                    llPreviousSchedule = ief_rst!Schedule
                    'either 'test' in url, or not sending to idc
                    If Not myIdc.isTest Then
                        If Not myIdc.DeleteFromIDCSchedule(llPreviousSchedule) Then
                            blDeleteOk = False
                            slFail = slFail & llPreviousSchedule & ","
                        Else
                            slSchedules = slSchedules & llPreviousSchedule & ","
                        End If
                    Else
                        slSchedules = slSchedules & llPreviousSchedule & ","
                    End If
                End If
            End If
 On Error GoTo 0
            ief_rst.MoveNext
        Loop
        If Len(slSchedules) > 0 Then
            slSchedules = mLoseLastLetter(slSchedules)
        End If
        If blDeleteOk Then
            If Not myIdc.isTest Then
                slDeletes = "( " & slSchedules & ")"
                '10/24/13 ddf changes for v60 only
                SQLQuery = "DELETE FROM ief_IDC_Enforced where iefScheduleID in " & slDeletes
                'SQLQuery = "DELETE FROM ief_IDC_Enforced where iefSplitRafCode in " & slDeletes
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand1:
                    Screen.MousePointer = vbDefault
                    gHandleError smPathForgLogMsg, "Export IDC-mDeleteSchedules"
                    mDeleteSchedules = False
                    Exit Function
                End If
                myFacts.WriteFacts "Deleted these schedules:" & vbCrLf & slSchedules
            Else
                myFacts.WriteFacts "Will delete these schedules:" & vbCrLf & slSchedules
            End If ' don't delete if in test mode
        Else
            mDeleteSchedules = False
            If Len(slFail) > 0 Then
                slFail = mLoseLastLetter(slFail)
            End If
            myErrors.WriteWarning "Could not delete these schedules:" & vbCrLf & slFail
        End If 'deletions ok
    Else
        myErrors.WriteWarning "No deleted spots for IDC- but entered mDeleteSchedules"
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mDeleteSchedules"
    mDeleteSchedules = False
End Function
Private Function mGetRafCode(llRCrfCode As Long) As Long

    On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    mGetRafCode = 0
    SQLQuery = "Select crfRafCode FROM crf_Copy_Rot_Header"
    SQLQuery = SQLQuery & " Where (crfCode = " & llRCrfCode & ")"
    Set crf_rst = gSQLSelectCall(SQLQuery)

    If Not crf_rst.EOF Then
        mGetRafCode = crf_rst!crfRafCode
    End If
    If igExportSource = 2 Then DoEvents
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mGetRafCode"
    Exit Function

End Function

Private Sub udcCriteria_IDCChg(ilValue As Integer)
    
    If ilValue = vbChecked Then
        If chkAll.Value = vbUnchecked Then
            chkAll.Value = vbChecked
        End If
        chkAll.Enabled = False
        lbcVehicles.Enabled = False
        If chkAllStation.Value = vbUnchecked Then
            chkAllStation.Value = vbChecked
        End If
        chkAllStation.Enabled = False
        lbcStation.Enabled = False
    Else
        chkAll.Enabled = True
        lbcVehicles.Enabled = True
        chkAllStation.Enabled = True
        lbcStation.Enabled = True
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
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("D", "IDC", "D", Trim$(edcStartDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
Private Sub mTestSites()
    Dim blIsConnectionError As Boolean
    Dim blIsStationError As Boolean
    Dim slErrorMessage As String
    Dim slUrl As String
    Dim myIdc As CIdc
    Dim slPrevious As String
    Dim slStationName As String
    Dim ilCounter As Integer
    Dim ilTotal As Integer
    Dim blCantIdcStation As Boolean
    
On Error GoTo ErrHand
    If igExportSource = 2 Then DoEvents
    Screen.MousePointer = vbHourglass
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    lacResult.Caption = ""
    'reset if they previously pressed to cancel
    slPrevious = ""
    imTerminate = False
    blCantIdcStation = False
    blIsConnectionError = False
    blIsStationError = False
    slErrorMessage = ""
    myErrors.WriteFacts "Testing site ids", True
    smExportDirectory = udcCriteria.DExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
    slUrl = mGetURL()
    If Len(slUrl) = 0 Then
        blIsConnectionError = True
        slErrorMessage = "  Could not read values from xml.ini."
        myErrors.WriteWarning slErrorMessage, True
        GoTo Cleanup
    End If 'url exists
    Set myIdc = New CIdc
    myIdc.LogPath = myIdc.CreateLogName(sgMsgDirectory & FILEDEBUG)
    ' debug mode
    If UCase(slUrl) = "TEST" Then
        'isTest: for testing only. Don't write to idc
        myIdc.isTest = True
        mSetResults "Running in 'Test Mode'", MESSAGEBLACK
    End If
    lacResult.Caption = "Attempting to connect to IDC"
    DoEvents
    myIdc.SoapUrl = slUrl
    myIdc.LogStart
    If Len(myIdc.SoapUrl) = 0 Then
        blIsConnectionError = True
        slErrorMessage = myIdc.ErrorMessage
        myErrors.WriteWarning slErrorMessage
        GoTo Cleanup
    End If
    'url didn't work, or IDC not operating. Log start is here.
    If Not myIdc.IsConnected() Then
        blIsConnectionError = True
        slErrorMessage = "Could not connect to IDC."
        myErrors.WriteWarning slErrorMessage
        GoTo Cleanup
    End If 'connected
    '4/25/13 add redundancy state
    If Not myIdc.IsMasterState Then
        myErrors.WriteWarning "IDC is in Backup Redundancy State."
        mSetResults "IDC is in Backup Redundancy State", MESSAGEBLACK
        myIdc.SoapUrl = mGetBackupUrl()
        If Len(myIdc.SoapUrl) = 0 Then
            blIsConnectionError = True
            slErrorMessage = myIdc.ErrorMessage
            myErrors.WriteWarning slErrorMessage
            GoTo Cleanup
        ElseIf Not myIdc.IsConnected() Then
            blIsConnectionError = True
            slErrorMessage = "Could not connect to IDC."
            myErrors.WriteWarning slErrorMessage
            GoTo Cleanup
        End If
    End If
    'now change address to special for testing site ids
    slUrl = mSiteIdUrl(slUrl)
    myIdc.SoapUrl = slUrl
    If Len(myIdc.SoapUrl) = 0 Then
        blIsConnectionError = True
        slErrorMessage = myIdc.ErrorMessage
        myErrors.WriteWarning slErrorMessage
        GoTo Cleanup
    End If
    Set myFacts = New CLogger
    With myFacts
        .LogPath = .CreateLogName(smExportDirectory & FILEFACTS)
        If myIdc.isTest Then
            .WriteFacts "Checking Site Ids --Test mode only", True
        Else
            .WriteFacts "Checking Site Ids ", True
        End If
    End With
    SQLQuery = "select COUNT(*) AS AMOUNT from att where rtrim(attidcreceiverid) <> '' "
    Set rst = gSQLSelectCall(SQLQuery)
    ilTotal = rst!amount
    SQLQuery = "select distinct attidcreceiverid  as id, attShfCode from att where rtrim(attidcreceiverid) <> '' ORDER BY attidcreceiverid"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        SQLQuery = "select shttCode, shttCallLetters from shtt ORDER by shttCode"
        Set rst2 = gSQLSelectCall(SQLQuery)
        Do While Not rst.EOF
            DoEvents
            ilCounter = ilCounter + 1
            lacResult.Caption = "Checking site ids " & ilCounter & " of " & ilTotal
            If rst!ID <> slPrevious Then
                slPrevious = rst!ID
                If myIdc.GetSiteId(slPrevious) Then
                    rst2.Filter = "shttcode = " & rst!attshfCode
                    If Not rst2.EOF Then
                        slStationName = Trim$(rst2!shttCallLetters)
                    Else
                        slStationName = ""
                    End If
                    myFacts.WriteFacts "Station " & slStationName & "'s id:" & Trim$(slPrevious) & " exists in production manager."
                Else
                    If Len(myIdc.ErrorMessage) > 0 Then
                        myFacts.WriteError "Cannot test stations ids (" & slStationName & " , " & slPrevious & ") : " & myIdc.ErrorMessage
                        myErrors.WriteError "Cannot test stations ids (" & slStationName & " , " & slPrevious & ") : " & myIdc.ErrorMessage
                        mSetResults "Cannot test stations.", MESSAGERED
                         blCantIdcStation = True
                        GoTo Cleanup
                    End If
                    rst2.Filter = "shttcode = " & rst!attshfCode
                    If Not rst2.EOF Then
                        slStationName = Trim$(rst2!shttCallLetters)
                    Else
                        slStationName = ""
                    End If
                    myFacts.WriteWarning "Station " & slStationName & "'s id:" & Trim$(slPrevious) & " does not exist in production manager."
                End If
            Else
                rst2.Filter = "shttcode = " & rst!attshfCode
                If Not rst2.EOF Then
                    slStationName = Trim$(rst2!shttCallLetters)
                Else
                    slStationName = ""
                End If
                myFacts.WriteFacts " Station " & slStationName & " shares id " & Trim$(slPrevious) & " with above."
            End If
            rst.MoveNext
        Loop
    Else
        mSetResults "No site ids to test", MESSAGEBLACK
        myFacts.WriteFacts "No site ids to test"
    End If
Cleanup:
    lacResult.Caption = "Checking site ids"
    If blIsConnectionError Then
        mSetResults "Could not connect to IDC.", MESSAGERED
        If Not myFacts Is Nothing Then
            myFacts.WriteWarning slErrorMessage
        End If
    ElseIf blIsStationError Then
        mSetResults "Some Site ids do not exist on IDC", MESSAGERED
    ElseIf Not blCantIdcStation Then
        mSetResults "All Site ids exist on IDC", MESSAGEGREEN
        myFacts.WriteFacts "All Site ids exist on IDC"
    End If
    lacResult.Caption = "Exports placed into: " & smExportDirectory
    If Not myFacts Is Nothing Then
        myFacts.WriteFacts "End Site id test", True
    End If
    myErrors.WriteFacts "End Site id test", True
    Set myIdc = Nothing
    Set myFacts = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "Export IDC-mTestSites"
End Sub
'6418
Private Function mSiteIdUrl(ByVal slUrl As String) As String
    Dim slTemp As String
    Dim ilPos As Integer
    
    If UCase(slUrl) = "TEST" Then
        mSiteIdUrl = slUrl
        Exit Function
    End If
    slTemp = ""
    'first '/'
    ilPos = InStr(1, slUrl, "/")
    If ilPos > 0 Then
        '2nd '/'
        ilPos = InStr(ilPos + 1, slUrl, "/")
        If ilPos > 0 Then
            'finally, the right '/'
            ilPos = InStr(ilPos + 1, slUrl, "/")
        End If
    End If
    If ilPos > 0 Then
        slTemp = Mid(slUrl, 1, ilPos - 1)
        slTemp = slTemp & ":9876/XD"
    End If
    mSiteIdUrl = slTemp
End Function
Private Function mPrepRecordsetDate() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Date", adDate
        End With
    myRs.Open
    myRs!Date.Properties("optimize") = True
    myRs.Sort = "Date"
    Set mPrepRecordsetDate = myRs
End Function
 Private Sub mBuildRsDate(myRs As ADODB.Recordset)
    Dim dlStart As Date
    Dim dlEnd As Date

    myRs.Filter = adFilterNone
 On Error GoTo ERRORBOX
    If myRs.RecordCount > 0 Then
        Do While Not myRs.EOF
            myRs.Delete
            myRs.MoveNext
        Loop
    End If
    If Not mRsRotations Is Nothing Then
        If mRsRotations.RecordCount > 0 Then
            mRsRotations.MoveFirst
            Do While Not mRsRotations.EOF
                dlStart = mRsRotations!Start
                If DateDiff("d", dlStart, smDate) > 0 Then
                    dlStart = smDate
                End If
                myRs.Filter = "Date = #" & dlStart & "#"
                If myRs.EOF Then
                    myRs.AddNew Array("Date"), Array(dlStart)
                End If
                dlEnd = mRsRotations!End
                dlEnd = DateAdd("d", 1, dlEnd)
                 'added '=' 7093
                If DateDiff("d", smEndDate, dlEnd) >= 2 Then
                    dlEnd = DateAdd("d", 1, smEndDate)
                End If
                myRs.Filter = "Date = #" & dlEnd & "#"
                If myRs.EOF Then
                    myRs.AddNew Array("Date"), Array(dlEnd)
                End If
                mRsRotations.MoveNext
            Loop
        End If
    End If
    myRs.Filter = adFilterNone
    Exit Sub
ERRORBOX:
    dlStart = dlStart
 End Sub
 Private Function mDateEarlier(ByVal slDateOne As String, slDateTwo As String) As String
    
    If gDateValue(gAdjYear(slDateOne)) < gDateValue(gAdjYear(slDateTwo)) Then
         mDateEarlier = slDateOne
    Else
        mDateEarlier = slDateTwo
    End If
End Function
Private Function mDateLater(ByVal slDateOne As String, slDateTwo As String) As String
    
    If gDateValue(gAdjYear(slDateOne)) > gDateValue(gAdjYear(slDateTwo)) Then
         mDateLater = slDateOne
    Else
        mDateLater = slDateTwo
    End If
End Function
Private Sub mCrfDateModified(slRegionStart As String, slRegionEnd As String, llSplitIndex As Long)
    'change slStart to the first 'Y' after this date.  SlEnd is the date just before the first 'N'
    'then, if there are still more left over, ?
    Dim slStart As String
    Dim slEnd As String
    Dim slMonday As String
    Dim llDateDiffStart  As Long
    Dim llDateDiffEnd As Long
    Dim blRet As Boolean
    Dim slTempStart As String
    Dim slTempEnd As String
    'get later date of region start, or export date.  Also earlier of end date
    slStart = mDateLater(slRegionStart, smDate)
    slEnd = mDateEarlier(slRegionEnd, smEndDate)
'   'days off of Monday
    slMonday = gObtainPrevMonday(slStart)
    llDateDiffStart = DateDiff("d", slMonday, slStart)
    'don't allow to look beyond this # of days
    llDateDiffEnd = DateDiff("d", slMonday, slEnd)
    'find first 'Y' after our allowed startdate
    Do
        blRet = mCrfDayIsYes(llDateDiffStart)
        If Not blRet Then
            llDateDiffStart = llDateDiffStart + 1
        End If
    Loop While Not blRet And llDateDiffStart <= llDateDiffEnd
    slRegionStart = DateAdd("d", llDateDiffStart, slMonday)
    'now try to find consecutive 'Y' for end
    llDateDiffStart = llDateDiffStart + 1
    Do
        blRet = mCrfDayIsYes(llDateDiffStart)
        If blRet Then
            llDateDiffStart = llDateDiffStart + 1
        End If
    Loop While blRet And llDateDiffStart <= llDateDiffEnd
    If llDateDiffStart < llDateDiffEnd Then
        slRegionEnd = DateAdd("d", llDateDiffStart - 1, slMonday)
    End If
    'that's our first consecutive set of 'Y'.  Now let's find more.
    Do
        slTempStart = ""
        slTempEnd = ""
        Do
            blRet = mCrfDayIsYes(llDateDiffStart)
            If Not blRet Then
                llDateDiffStart = llDateDiffStart + 1
            End If
        Loop While Not blRet And llDateDiffStart <= llDateDiffEnd
        If llDateDiffStart < llDateDiffEnd Then
            slTempStart = DateAdd("d", llDateDiffStart, slMonday)
        End If
        'now try to find consecutive 'Y' for end
        llDateDiffStart = llDateDiffStart + 1
        Do
            blRet = mCrfDayIsYes(llDateDiffStart)
            If blRet Then
                llDateDiffStart = llDateDiffStart + 1
            End If
        Loop While blRet And llDateDiffStart <= llDateDiffEnd
        If llDateDiffStart <= llDateDiffEnd + 1 Then
            slTempEnd = DateAdd("d", llDateDiffStart - 1, slMonday)
        End If
        If Len(slTempStart) > 0 And Len(slTempEnd) Then
            'here is our set of new region
           ' MsgBox "split start = " & slTempStart & " end = " & slTempEnd
            With tmIDCSplit(llSplitIndex)
                mRsCrfSplit.AddNew Array("SplitIndex", "Start", "End", "Region", "Rotation", "Found", "StartTime", "EndTime"), Array(llSplitIndex, slTempStart, slTempEnd, .lRafCode, .iRotation, False, .sStartTime, .sEndTime)
            End With
        End If
    Loop While llDateDiffStart <= llDateDiffEnd
    
End Sub
Private Function mCrfDayIsYes(llStartDiff As Long) As Boolean
    Dim blIsYes As Boolean
    
    blIsYes = False
    Select Case llStartDiff
        Case 0
            If crf_rst!crfMo = "Y" Then
                blIsYes = True
            End If
        Case 1
            If crf_rst!crfTu = "Y" Then
                blIsYes = True
            End If
        Case 2
            If crf_rst!crfWe = "Y" Then
               blIsYes = True
            End If
        Case 3
            If crf_rst!crfTh = "Y" Then
                blIsYes = True
            End If
        Case 4
            If crf_rst!crfFr = "Y" Then
                blIsYes = True
            End If
        Case 5
            If crf_rst!crfSa = "Y" Then
               blIsYes = True
            End If
        Case 6
            If crf_rst!crfSu = "Y" Then
                blIsYes = True
            End If
        Case Else
            blIsYes = False
    End Select
    mCrfDayIsYes = blIsYes
End Function
Private Function mGroupSiteIds(llIndex As Long, lRsGroup As ADODB.Recordset, slMajorSplit As String, slMinorSplit As String, slReceivers As String) As String
'O: add additional site ids to slReceivers
'6419
    Dim slSql As String
    Dim slRet As String
    Dim slSite As String
    Dim slCall As String
    Dim slNewSite As String
    Dim slNewCall As String
    Dim slGroupChoice As String
    Dim ilStation As Integer
    Dim slSQLDate As String
On Error GoTo ERRORBOX
    slRet = slReceivers & slMajorSplit
    slSQLDate = " AND attOnAir <= '" & Format(smEndDate, sgSQLDateForm) & "' AND attOffAir >= '" & Format(smDate, sgSQLDateForm) & "' and attDropDate >= '" & Format(smDate, sgSQLDateForm) & "' "
    With tmIDCReceiver(llIndex)
        slSite = Trim(.sReceiverID)
        slCall = Trim(.sCallLetters)
        'the attCode says "use these site ids when this agreement is being called"
        lRsGroup.Filter = "AttCode = " & .lAttCode
        If lRsGroup.EOF Then
            lRsGroup.AddNew Array("AttCode", "SiteId", "Call"), Array(.lAttCode, slSite, slCall)
            slSql = "select attIDCGroupType as IdcGroup, attShfCode as StationCode, attMulticast from att where attCode = " & .lAttCode
           'test before ddf changes
           ' slSql = "select attComp as IdcGroup, attShfCode as StationCode, attMulticast from att where attCode = " & .lAttCode
            Set rst = gSQLSelectCall(slSql)
            If Not rst.EOF Then
                ilStation = rst!StationCode
                slGroupChoice = rst!IDCGroup
                Select Case slGroupChoice
                    'by station
                    Case "S"
                        'this sql is for either of the cases below
                        slSql = "select distinct attIDCReceiverId as SiteID, shttCallLetters as Letters from att inner join shtt on attshfCode = shttCode where attIDCGroupType = 'S' AND SiteId <> '" & slSite & "' AND attshfCode = " & ilStation
                        'test before ddf changes
                        'slSql = "select distinct attIDCReceiverId as SiteID, shttCallLetters as Letters from att inner join shtt on attshfCode = shttCode where attcomp = 1 AND SiteId <> '" & slSite & "' AND attshfCode = " & ilStation
                        slSql = slSql & slSQLDate
                        'multicast? then get other multicasts site ids regardless of their attIdcgroup settings
                        If rst!attMulticast = "Y" Then
                            slSql = slSql & " UNION Select distinct attIDCReceiverId as SiteId, shttCallLetters as Letters from att inner join shtt on attShfCode = shttCode where shttMulticastGroupid = (select shttMulticastGroupid from shtt where shttcode = "
                            slSql = slSql & ilStation & " AND shttMulticastGroupId > 0 ) and shttcode <> " & ilStation
                            slSql = slSql & slSQLDate
                            Set rst = gSQLSelectCall(slSql)
                        Else
                            Set rst = gSQLSelectCall(slSql)
                        End If
                    'by location
                    Case "L"
                        slSql = "select attIDCReceiverId as SiteId, shttCallLetters as Letters from att inner join shtt on attShfCode = shttCode where  "
                        slSql = slSql & " ATTidcGroupType = 'L' AND attIDCReceiverId <> '" & slSite & "'  AND shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip + cast(shttmktcode as varchar(5)) = "
                        slSql = slSql & "(select shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip + cast(shttmktcode as varchar(5)) from shtt where shttCode = " & ilStation & ")"
                    'test before ddf changes
                     '   slSql = "select attIDCReceiverId as SiteId, shttCallLetters as Letters from att inner join shtt on attShfCode = shttCode where  "
                     '   slSql = slSql & " ATTCOMP = " & slGroupChoice & " AND attIDCReceiverId <> '" & slSite & "'  AND shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip = "
                     '   slSql = slSql & "(select shttOnAddress1 + shttonAddress2 + shttOnCity + shttOnState + shttOnZip from shtt where shttCode = " & ilStation & ")"
                        
                        slSql = slSql & slSQLDate
                        Set rst = gSQLSelectCall(slSql)
                    'do not group...make rst empty
                    Case Else
                        rst.Filter = "IdcGroup = 400"
                End Select
            End If
            Do While Not rst.EOF
                slNewSite = Trim(rst!SiteID)
                slNewCall = Trim(rst!letters)
                If slNewSite <> "" Then
                    If Len(slNewCall) > 10 Then
                        slNewCall = Mid(slNewCall, 1, 10)
                    End If
                    If slSite <> slNewSite Then
                        lRsGroup.AddNew Array("AttCode", "SiteId", "Call"), Array(.lAttCode, slNewSite, slNewCall)
                        If InStr(1, slMajorSplit & slRet, slMajorSplit & slNewSite & slMinorSplit) = 0 Then
                            slRet = slRet & slNewCall & slMinorSplit & slNewSite & slMajorSplit
                        End If
                    End If
                End If
                rst.MoveNext
            Loop
        Else
            Do While Not lRsGroup.EOF
                slNewSite = Trim(lRsGroup!SiteID)
                slNewCall = Trim(lRsGroup!Call)
                If slSite <> slNewSite Then
                    If InStr(1, slRet, slMinorSplit & slNewSite & slMajorSplit) = 0 Then
                        slRet = slRet & slNewCall & slMinorSplit & slNewSite & slMajorSplit
                    End If
                End If
                lRsGroup.MoveNext
            Loop
        End If
    End With
    mGroupSiteIds = mLoseLastLetter(slRet)
    Exit Function
ERRORBOX:
    slRet = ""
    gHandleError smPathForgLogMsg, "Export IDC-mGroupSiteIds"
    mSetResults "Problem grouping site ids!", MESSAGERED
    mGroupSiteIds = slRet
End Function
'Private Function mBuildSafeBlackoutStation(slStations As String, llSplitIndex As Long, slSplit As String) As String
'    Dim slRet As String
'    Dim slUnsafe As String
'    Dim slSites() As String
'    Dim c As Integer
'
'    If Len(slStations) > 2 Then
'        slUnsafe = mLoseLastLetter(slStations)
'        slSites = Split(slUnsafe, slSplit)
'        For c = 0 To UBound(slSites)
'            rsBlackout.Filter = "SiteId = '" & slSites(c) & "'"
'
'        Next c
'    Else
'        slRet = slStations
'    End If
'    mBuildSafeBlackoutStation = slRet
'End Function
'
'Private Function mBlackoutStationSafe(slSiteId As String, slGeneric As String, llSplitIndex As Long) As Boolean
'    Dim blRet As Boolean
'    Dim slSite As String * 5
'
'    slSite = slSiteId
'    rsBlackout.Filter = "SiteId= '" & slSite & "' AND Generic = '" & slGeneric & "'"
'
'End Function

 
