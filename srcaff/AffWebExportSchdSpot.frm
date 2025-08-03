VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmWebExportSchdSpot 
   Caption         =   "CSI Electronic Affidavit"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffWebExportSchdSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frWebVendor 
      Caption         =   "Web Vendor Info Export"
      Height          =   615
      Left            =   9120
      TabIndex        =   30
      Top             =   120
      Width           =   3255
      Begin ComctlLib.ProgressBar PbcWebVendor 
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.TextBox lacStationMsg 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   $"AffWebExportSchdSpot.frx":08CA
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Timer tmcFilterDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9360
      Top             =   4680
   End
   Begin VB.Frame frFilter 
      Caption         =   "Filters"
      Height          =   945
      Left            =   5280
      TabIndex        =   22
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton rbcFilter 
         Caption         =   "DMA"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "Format"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   27
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "MSA"
         Height          =   200
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "State"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "Station"
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ListBox lbcFilter 
         Height          =   840
         Left            =   1920
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1575
      TabIndex        =   1
      Top             =   105
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "01/27/2023"
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
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9345
      Top             =   1740
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVeh 
      Height          =   2055
      Left            =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1680
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
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5820
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Results"
      Top             =   1320
      Width           =   3405
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9255
      Top             =   3990
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Vehicles"
      Top             =   1440
      Width           =   2085
   End
   Begin VB.TextBox txtCallLetters 
      Height          =   360
      Left            =   7080
      TabIndex        =   14
      Top             =   3960
      Width           =   1320
   End
   Begin VB.OptionButton rbcVehicles 
      Caption         =   "Active Vehicles"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton rbcVehicles 
      Caption         =   "All Vehicles"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   4680
      Width           =   2685
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   4680
      Width           =   2685
   End
   Begin VB.ListBox lbcMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      ItemData        =   "AffWebExportSchdSpot.frx":0906
      Left            =   5640
      List            =   "AffWebExportSchdSpot.frx":0908
      TabIndex        =   12
      Top             =   1680
      Width           =   3705
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   900
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "7"
      Top             =   105
      Width           =   405
   End
   Begin VB.ListBox lbcStation 
      Height          =   1815
      ItemData        =   "AffWebExportSchdSpot.frx":090A
      Left            =   3840
      List            =   "AffWebExportSchdSpot.frx":090C
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All Stations"
      Height          =   195
      Left            =   3855
      TabIndex        =   9
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1380
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9360
      Top             =   3315
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5250
      FormDesignWidth =   9615
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   582
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Stations"
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label lbcWebType 
      Caption         =   "Production Website"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   5040
      Width           =   1455
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
      TabIndex        =   21
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label lacFTPStatus 
      Height          =   255
      Left            =   225
      TabIndex        =   18
      Top             =   4710
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCallLetters 
      Caption         =   "Processing:"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Days"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   135
      Width           =   1335
   End
End
Attribute VB_Name = "frmWebExportSchdSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private smStartDate As String
Private smEndDate As String
Private smMode As String
Private smMODate As String
Private imNumberDays As Integer
Private imVefCode As Integer
Private imExptVefCode() As Integer
Private smVefName As String
Private smShowVehName As String
Private smVehicleType As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmToCpyRot As Integer
Private hmToMultiUse As Integer
Private hmToEventInfo As Integer
Private hmToHeader As Integer
Private hmToDetail As Integer
Private hmFrom As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private smWebExports As String
Private lmLastattCode() As Long
Private lmAttCode As Long
Private lmWebAttID As Long
Private smToFileDetail As String
Private smToFileHeader As String
Private smToMultiUse As String
Private smToCpyRot As String
Private imSpotCount As Integer
Private smWebHeader As String
Private smWebSpots As String
Private smWebCopyRot As String
Private smWebMultiUse As String
Private smWebEventInfo As String
Private smFileName As String
Private smStatus As String
Private smMsg1 As String
Private smMsg2 As String
Private smDTStamp As String
Private imNeedToSend As Integer
Private smReindex As String
Private smSendEmails As String
Private lmTotalAddSpotCount As Long
Private lmTotalDeleteSpotCount As Long
Private lmTotalEmails As Long
Private lmTotalComments As Long
Private lmTotalHeaders As Long
Private imShttCode As Integer
'Private lmDupAstAdd As Long
Private lmDupAstDel As Long
Private hmAst As Integer
Private imFailedStations As Integer
Private lmFailedStaExpIdx As Long
Private lmMaxEsfCode As Long
Private lmTtlHeaders As Long
Private lmTtlComments As Long
Private lmTtlMultiUse As Long
Private smToEventInfo As String
Private lmGsfCode As Long
Private smWebWorkStatus As String
Private imWaiting As Integer
Private imSomeThingToDo As Integer
Private lmFileDeleteSpotCount As Long
Private lmFileAddSpotCount As Long
Private lmFileHeaderCount As Long
Private lmFileCommentCount As Long
Private lmFileMultiUseCount As Long
Private lmFileEventCount As Long
Private imFTPIsOn As Boolean
Private imFailures As Boolean
Private smUseActual As String   'Export to the web after posting in Traffic and Transmit the posted date and time to the web
Private smSuppressLog As String
Private lmEstimatesExist As Boolean
Private smEstimatedDate As String
Private smEstimatedStartTime As String
Private smEstimatedEndTime As String
Private imLoadMultiplier As Integer
Private imNoAirPlays As Integer
Private ilOldImpLayout As Integer
Private lmEqtCode As Long
'Web
Private lmWebTtlComments As Long
Private lmWebTtlMultiUse As Long
Private lmWebTtlHeaders As Long
Private lmWebTtlSpots As Long
Private lmWebTtlEmail As Long
Private lmTtlEventSpots As Long
Private lmWebTtlEventSpots As Long
Private imIdx As Integer
Private rstWebQ As ADODB.Recordset
Private lmTotalRecordsProcessed As Long
Private tmCsiFtpInfo As CSIFTPINFO
Private tmCsiFtpStatus As CSIFTPSTATUS
Private tmCsiFtpErrorInfo As CSIFTPERRORINFO
Private smWebImports As String
Private smCheckIniReIndex As String
Private smMinSpotsToReIndex As String
Private smMaxWaitMinutes As String
Private imFTPEvents As Boolean
Private imFtpInProgress As Boolean
Private mFtpArray() As String
Private smAttWebInterface As String
Private rst_Gsf As ADODB.Recordset
Private rst_DAT As ADODB.Recordset
Private rst_Est As ADODB.Recordset
Private smMGsOnWeb As String
Private smReplacementsOnWeb As String
Private smPrevMsg As String
Private imExportedVefArray() As Integer
Private lmWebPostedAttRecs() As Long
Private lmPostedAttRecs() As Long
Private smStaFailedToExport() As String
Private tmEsf() As ESF
Private tmEdf() As EDF
Private Type AIRTIMEINFO
    sPrevPldgSTime As String * 11
    sPrevPldgETime As String * 11
    sBaseStartTime As String * 11
    sBaseStartDate As String * 10
    sEstimatedEndTime As String * 11
    sEstimatedStartTime As String * 10
End Type
Private tmAirTimeInfo() As AIRTIMEINFO
Private smJelliExport As String
Private smVehicleExportJelli As String
Private smAttExportToJelli As String
Private hmJelli As Integer
Private smJelliFileName As String
Const cmOneMegaByte As Long = 1000000
Const cmOneSecond As Long = 1000
Const cmPathForgLogMsg As String = "WebExportLog.Txt"
Private myEnt As CENThelper
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
Private smGridTypeAhead  As String
'11/3/17
Private smEDCDate As String
Private smTxtNumberDays As String
Private smFilterType As String

Private bmInStationFill As Boolean

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
Dim blStationListAlreadyDrawn As Boolean
Dim tmVefReqAlerts() As AUF
Dim tmVefOtherAlerts() As AUF


Private Sub cmdCancel_Click()
    On Error GoTo ErrHand
    If imExporting Then
        imTerminate = True
        mCloseFiles
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmWebExportSchdSpot
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-cmdCancel_Click"
    Exit Sub
End Sub

Private Sub mFillVehicle()

    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slTemp As String
    Dim ilMaxLen As Integer
    Dim slNowDate As String
    Dim llVef As Long
    Dim ilVff As Integer
    Dim llRow As Long
    Dim llGridRow As Long
    Dim llCol As Long
    Dim blAllVehicles As Boolean
    ReDim ilVefCode(0 To 0) As Integer
    
    On Error GoTo ErrHand
    'TTP 9683
    If edcDate.Text = "" Then
        slNowDate = Format(gNow(), sgSQLDateForm)
    Else
        slNowDate = Format(edcDate.Text, sgSQLDateForm)
    End If
    '7/1/20: File not required if dates not changed. Note: changing All or active vehicles require that the grid to re-populate (smEDCDate set to blank to force pop).
    If (smEDCDate = edcDate.Text) And (smTxtNumberDays = txtNumberDays.Text) Then
        Exit Sub
    End If
    grdVeh.Visible = False
    grdVeh.Redraw = False
    blAllVehicles = False
    ReDim ilVefCode(0 To 0) As Integer
    If chkAll.Value = vbChecked Then
        blAllVehicles = True
    Else
        For llRow = grdVeh.FixedRows To grdVeh.Rows - 1
            If Trim(grdVeh.TextMatrix(llRow, VEHINDEX)) <> "" Then
                If grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                    ilVefCode(UBound(ilVefCode)) = Val(grdVeh.TextMatrix(llRow, VEHCODEINDEX))
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
            End If
        Next llRow
    End If
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
   
    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND attExportType <> 0" & " AND attExportToWeb = 'Y'"
    'Set rst = cnn.Execute(SQLQuery)
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        llVef = gBinarySearchVef(CLng(rst!attvefCode))
        If llVef = -1 Then
            llVef = gPopSellingVehicles()
            llVef = gBinarySearchVef(CLng(rst!attvefCode))
        End If
        If llVef <> -1 Then
            If (rbcVehicles(0).Value = vbTrue) Or ((rbcVehicles(1).Value = vbTrue) And (tgVehicleInfo(llVef).sState = "A")) Then
                If tgVehicleInfo(llVef).iVefCode > 0 Then   'Reference a Log vehicle
                    'Determine if vehicle is to bemerged into Log vehicle on the web
                    For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                        If tgVehicleInfo(llVef).iCode = tgVffInfo(ilVff).iVefCode Then
                            If tgVffInfo(ilVff).sMergeWeb <> "S" Then
                                'Find the Log vehicle name and code
                                For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                                    If tgVehicleInfo(ilLoop).iCode = tgVehicleInfo(llVef).iVefCode Then
                                        'Check that name has not been previously added
                                        If Not mFindDuplVeh(tgVehicleInfo(ilLoop).iCode) Then
                                            mAddToGrid llGridRow, CLng(ilLoop)
                                        End If
                                    End If
                                Next ilLoop
                            Else
                                mAddToGrid llGridRow, llVef
                            End If
                            Exit For
                        End If
                    Next ilVff
                Else
                    If Not mFindDuplVeh(tgVehicleInfo(llVef).iCode) Then
                        mAddToGrid llGridRow, llVef
                    End If
                End If
            End If
        End If
        rst.MoveNext
    Loop
    mFindAlertsForGrdVeh
    mVehSortCol VEHINDEX
    'mVehSortCol LOGINDEX
    If blAllVehicles Then
        chkAll.Value = vbChecked
    Else
        For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
            If Trim(grdVeh.TextMatrix(llRow, VEHINDEX)) <> "" Then
                grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "0"
                For llVef = 0 To UBound(ilVefCode) - 1 Step 1
                    If ilVefCode(llVef) = Val(grdVeh.TextMatrix(llRow, VEHCODEINDEX)) Then
                        grdVeh.TextMatrix(llRow, SELECTEDINDEX) = "1"
                        Exit For
                    End If
                Next llVef
                mPaintRowColor llRow
            End If
        Next llRow
        chkAllStation.Visible = True
        mFillStations
    End If
    Erase ilVefCode
    grdVeh.Row = 0
    grdVeh.Col = VEHCODEINDEX
    grdVeh.Redraw = True
    grdVeh.Visible = True
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mFillVehicle"
    Exit Sub
End Sub

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
    grdVeh.Visible = False
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
    grdVeh.Visible = True
'    If ilCount <> 1 Then
'        edcTitle3.Visible = False
'        chkAllStation.Visible = False
'        lbcStation.Visible = False
'        lbcStation.Clear
'    Else
'        edcTitle3.Visible = True
        chkAllStation.Visible = True
'        lbcStation.Visible = True
'        mFillStations
'    End If
    mFillStations
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-chkAll_Click"
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
    
    'lacStationMsg.Visible = Not lbcStation.Visible
    If chkAllStation.Value = vbChecked Then
        lbcStation.Visible = False
        lacStationMsg.Visible = True
    Else
        '11/8/19: Fill only during export and all unchecked
        If (rbcFilter(4).Value = True) And (iValue = False) Then
            imAllStationClick = True
            mFillStations True
            imAllStationClick = False
        End If
        lbcStation.Visible = True
        lacStationMsg.Visible = False
    End If
    
    
   
    If lbcStation.ListCount > 0 Then
        gSetMousePointer grdVeh, grdVeh, vbHourglass
        imAllStationClick = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
        gSetMousePointer grdVeh, grdVeh, vbDefault
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-chkAllStation_Click"
    Exit Sub
End Sub

Private Sub cmdExport_Click()
    
    Dim iRet As Integer
    Dim slResult As String
    Dim llIdx As Long
    Dim ilIdx2 As Long
    Dim ilRet As Integer
    
    cmdExport.Enabled = False
    '11/3/17
    tmcDelay.Enabled = False
    tmcFilterDelay.Enabled = False
    lmEqtCode = -1
    lgSTime1 = timeGetTime
    lgCount1 = 0
    lgCount2 = 0
    lgCount3 = 0
    lgCount4 = 0
    lgCount5 = 0
    lgCount6 = 0
    lgCount7 = 0
    lgCount8 = 0
    lgCount9 = 0
    lgCount10 = 0
    lgCount11 = 0
    lgTtlTime1 = 0
    lgTtlTime2 = 0
    lgTtlTime3 = 0
    lgTtlTime4 = 0
    lgTtlTime5 = 0
    lgTtlTime6 = 0
    lgTtlTime7 = 0
    lgTtlTime8 = 0
    lgTtlTime9 = 0
    lgTtlTime10 = 0
    lgTtlTime11 = 0
    lgTtlTime12 = 0
    lgTtlTime13 = 0
    lgTtlTime14 = 0
    lgTtlTime15 = 0
    lgTtlTime16 = 0
    lgTtlTime17 = 0
    lgTtlTime18 = 0
    lgTtlTime19 = 0
    lgTtlTime20 = 0
    lgTtlTime21 = 0
    lgTtlTime22 = 0
    lgTtlTime23 = 0
    lgTtlTime24 = 0
    
    On Error GoTo ErrHand
    lgSTime24 = timeGetTime
    
    'D.S. 09/06/17
    '11/3/17
    If Not gTestAccessToWebServer() And Not igDemoMode Then
        gMsgBox "WARNING!" & vbCrLf & vbCrLf & _
               "Web Server Access Error: The Affiliate System does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
        "No data will be exported to the web site." & vbCrLf & _
        "No data will be imported from the web site." & vbCrLf & _
        "Sign off system immediately and contact system administrator.", vbExclamation
        Exit Sub
    End If
    
    If sgWebSiteNeedsUpdating = "True" Then
        gLogMsg "Web Version: " & sgWebSiteVersion & " does not agree with Affiliate Web Version: " & sgWebSiteExpectedByAffiliate & " No Imports Are Allowed", "WebExportLog.Txt", False
        gMsgBox "Web Version: " & sgWebSiteVersion & " does not agree with Affiliate Web Version: " & sgWebSiteExpectedByAffiliate & sgCRLF & sgCRLF & "          No Exports Are Allowed Until Corrected." & sgCRLF & sgCRLF & "                       Call Counterpoint!"
        gSetMousePointer grdVeh, grdVeh, vbDefault
        If igExportSource <> 2 Then
            Unload frmWebExportSchdSpot
        End If
        cmdExport.Enabled = True
        Exit Sub
    End If
    If imExporting Then
        cmdExport.Enabled = True
        Exit Sub
    End If
    imExporting = True
    'Validate all of the user's options
    If Not mValidateUserInput Then
        igExportReturn = 2
        imExporting = False
        cmdExport.Enabled = True
        Exit Sub
    End If
    '11/8/19: Fill only during export and all unchecked
    If (rbcFilter(4).Value = True) And (chkAllStation.Value = vbChecked) Then
        mFillStations True
    End If
    If Not imFTPIsOn Then
        gLogMsg "FTP is Off, therefore the Spots will not be exported", "WebExportLog.Txt", False
    End If
    If (igDemoMode) Then
        gLogMsg "Running in Demo Mode, therefore the Spots will not be exported", "WebExportLog.Txt", False
    End If
    mSaveCustomValues
    If (Not igDemoMode) Then
        ilRet = mWebGetPostedAttRecs(smStartDate, smEndDate)
        If ilRet Then
            ilRet = mWebProcessPostedAttRecs
            If Not ilRet Then
                ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                gLogMsg "Error: mWebProcessPostedAttRecs Failed to get Posted Att records from the web.", "WebExportLog.Txt", False
                cmdExport.Enabled = True
                Exit Sub
            End If
        Else
            ilRet = gCustomEndStatus(lmEqtCode, 2, "")
            gLogMsg "Error: mWebGetPostedAttRecs Failed to get Posted Att records from the web.", "WebExportLog.Txt", False
            cmdExport.Enabled = True
            Exit Sub
        End If
    Else
        ReDim lmWebPostedAttRecs(0 To 0) As Long
    End If
    'Get all of the latest passwords and email addresses from the web
    gRemoteTestForNewEmail
    gRemoteTestForNewWebPW
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    'Open the CSF file with an API call
    If Not mOpenCSFFile Then
        igExportReturn = 2
        imExporting = False
        ilRet = gCustomEndStatus(lmEqtCode, 2, "")
        mClose
        mCloseFiles
        cmdExport.Enabled = True
        Exit Sub
    End If
    sgTaskBlockedName = "Counterpoint Affidavit System"
    lgETime24 = timeGetTime
    lgTtlTime24 = lgTtlTime24 + (lgETime24 - lgSTime24)
    gLogMsg "Start Writing Export Files", "WebExportLog.Txt", False
    If Not mInitiateExport Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        ilRet = gCustomEndStatus(lmEqtCode, 2, "")
        gLogMsg "Error: mInitiateExport Failed.  Halting Export", "WebExportLog.Txt", False
        imExporting = False
        mClose
        mCloseFiles
        Call gCloseCSFFile
        cmdExport.Enabled = True
        Exit Sub
    End If
    lgSTime24 = timeGetTime
    '*** If we made it to here then the Export Process was successful!
    '*** Close everything and get out gracefully
    'Clear the Alerts
    gLogMsg "Start Clearing Alerts", "WebExportLog.Txt", False
    mClearAlerts
    gLogMsg "Clearing Alerts Completed", "WebExportLog.Txt", False
    'Force Alert Check
    gLogMsg "Start Alert Force Check", "WebExportLog.Txt", False
    iRet = gAlertForceCheck()
    gLogMsg "Alert Force Check Completed", "WebExportLog.Txt", False
    If Not igDemoMode Then
        'Update Web Access Control
        gLogMsg "Update Web AccessControl", "WebExportLog.Txt", False
        If CDbl(sgWebSiteVersion) >= 7.1 Then
            iRet = gWebUpdateAccessControl
            If iRet Then
                gLogMsg "Update Web Access Control Completed", "WebExportLog.Txt", False
            Else
                gLogMsg "Update Web Access Control Failed", "WebExportLog.Txt", False
            End If
        End If
    End If
    'Close the CSF.btr file with an API call
    gLogMsg "Closing CSF File", "WebExportLog.Txt", False
    Call gCloseCSFFile
    gLogMsg "CSF File Closed Successfully", "WebExportLog.Txt", False
    gLogMsg "", "WebExportLog.Txt", False
    lgETime24 = timeGetTime
    lgTtlTime24 = lgTtlTime24 + (lgETime24 - lgSTime24)
    mLogTimingResults
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    mClose
    imExporting = False
    
    'D.S. 12/27/17
    If bgTaskBlocked And igExportSource <> 2 Then
        gMsgBox "*** Some Station(s) Need to be Re-Exported. ***" & vbCrLf & vbCrLf & "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""

    If imFailedStations Then
        gLogMsg " ", "StationsNotExported.Txt", False
        gLogMsg " " & "**** " & CStr(lmFailedStaExpIdx) & " station(s) failed to export due to having spots posted during the export period. ****", "StationsNotExported.Txt", False
        gMsgBox "Press OK for a listing of stations, in the Results box, that FAILED to Export ." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "A text listing is available in your Messages folder as well.  See StationsNotExported.Txt", vbOKOnly
        lbcMsg.Clear
        For llIdx = 0 To lmFailedStaExpIdx - 1 Step 1
           gLogMsg " " & Trim$(smStaFailedToExport(llIdx)) & " for the export period " & smStartDate & " - " & smEndDate, "StationsNotExported.Txt", False
           SetResults smStaFailedToExport(llIdx), RGB(255, 0, 0)
        Next llIdx
        gSetMousePointer grdVeh, grdVeh, vbHourglass
    End If
    If lmTotalAddSpotCount > 0 Then
        If igDemoMode Or (bgVendorToWebAllowed And imFTPIsOn) Then
            bgVendorExportSent = True
            igWVImportElapsed = 0
            dgWvImportLast = 0
        End If
    End If
    gSetMousePointer grdVeh, grdVeh, vbDefault
    cmdExport.Enabled = True
    Exit Sub
    
ErrHand:
    ilRet = gCustomEndStatus(lmEqtCode, 2, "")
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-cmdExport_Click"
    imExporting = False
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    'Debug
    'Resume Next
    Exit Sub
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
    'TTP 9683
    If edcDate.Text <> "" Then
        '7/1/20: Only require re-pop if fields changed
        If (smEDCDate <> edcDate.Text) Or (smTxtNumberDays <> txtNumberDays.Text) Then
            mFillVehicle
            'chkAll.Value = vbChecked
        End If
    End If
    tmcDelay.Enabled = True
End Sub

Private Sub Form_Activate()
    
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    Dim llRow As Long
    Dim llCol As Long
    Dim llPos As Long
    
    If imFirstTime Then
        gSetMousePointer grdVeh, grdVeh, vbHourglass
        mSetGridColumns
        mSetGridTitles
        'gGrid_IntegralHeight grdVeh
        'gGrid_FillWithRows grdVeh
        ''D.S. 07-28-17
        'grdVeh.Height = grdVeh.Height + 30
        'rbcVehicles(1).Value = True
        udcCriteria.Left = Label1.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = txtNumberDays.Top + txtNumberDays.Height
        'rbcVehicles(0).Top = udcCriteria.Top + llPos 'edcTitle1.Top - 60
        'rbcVehicles(1).Top = rbcVehicles(0).Top
        'rbcVehicles(1).Left = rbcVehicles(0).Left + rbcVehicles(0).Width
        edcTitle1.Visible = False
        udcCriteria.Action 6
        '11/9/18: Expand grid height
        llPos = udcCriteria.GetCtrlBottom()
        rbcVehicles(0).Top = udcCriteria.Top + llPos 'edcTitle1.Top - 60
        rbcVehicles(1).Top = rbcVehicles(0).Top
        rbcVehicles(1).Left = rbcVehicles(0).Left + rbcVehicles(0).Width
        
        'frFilter.Top = rbcVehicles(0).Top - 1200
        'frFilter.Left = rbcVehicles(1).Left + 1000
        'frFilter.Height = 1500
        
        grdVeh.Top = rbcVehicles(0).Top + rbcVehicles(0).Height + 60
        grdVeh.Height = chkAll.Top - grdVeh.Top - grdVeh.RowHeight(0) / 2
        gGrid_IntegralHeight grdVeh
        gGrid_FillWithRows grdVeh
        grdVeh.Height = grdVeh.Height + 30
        edcTitle2.Top = rbcVehicles(0).Top
        lbcMsg.Top = grdVeh.Top
        lbcMsg.Height = grdVeh.Height
        edcTitle3.Top = rbcVehicles(0).Top
        lbcStation.Top = grdVeh.Top
        lbcStation.Height = grdVeh.Height
        If UBound(tgEvtInfo) > 0 And igExportSource = 2 Then
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
                If llVef = -1 Then
                    llVef = gPopSellingVehicles()
                    llVef = gBinarySearchVef(CLng(rst!attvefCode))
                End If
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
'            If mGetGrdSelCount() = 1 Then
'                edcTitle3.Visible = True
'                chkAllStation.Visible = True
'                chkAllStation.Value = vbUnchecked
                lbcStation.Visible = True
'                mFillStations
'                chkAllStation.Value = vbChecked
'            End If
            smFilterType = "Station"
            mFillStations
            grdVeh.Redraw = True
        Else
            mFillVehicle
            chkAll.Value = vbChecked
        End If
        gSetMousePointer grdVeh, grdVeh, vbDefault
        If igExportSource = 2 Then
            slNowStart = gNow()
            cmdExport.Enabled = True
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "CSIWebResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "CSI Web Result List, Started: " & slNowStart
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
            gLogMsgWODT "W", hlResult, "CSI Web Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            hgExportResult = 0
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        lacStationMsg.Move lbcStation.Left, lbcStation.Top, lbcStation.Width, lbcStation.Height
        lacStationMsg.ZOrder
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()

    cmdExport.Visible = True
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
    Dim slTemp As String
    Dim FTPIsOn As String
    Dim ilValue1 As Integer
    Dim slSection As String
    '9851
    Dim myRemapMonitor As cRemapper
    
    On Error GoTo ErrHand
    
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    '9915
    frWebVendor.Top = lblCallLetters.Top - 100
    frWebVendor.Left = lblCallLetters.Left
    frWebVendor.Visible = False
    '11/3/17
    smTxtNumberDays = ""
    smEDCDate = ""
        
    lblNote.Visible = False
    imLastVehColSorted = -1
    imLastVehSort = -1
    smGridTypeAhead = ""
    bmInStationFill = False
    ' First load all the information we need from the ini file.
    Call gLoadOption(sgWebServerSection, "FTPIsOn", FTPIsOn)
    If Val(FTPIsOn) < 1 Then
        ' FTP is turned off. Return success.
        ' Note: This will be the case when the affiliate system and IIS is running on the same machine.
        '       Usually only while testing.
        imFTPIsOn = False
        lacFTPStatus.Caption = "FTP is Off"
        lacFTPStatus.ForeColor = vbRed
        lacFTPStatus.Visible = True
    Else
        imFTPIsOn = True
        lacFTPStatus.Caption = "FTP is On"
        lacFTPStatus.ForeColor = &HFF00& 'Green
        lacFTPStatus.Visible = True
    End If
    '10000
    lbcWebType.Left = lacFTPStatus.Left
    If igDemoMode Then
        lbcWebType.Caption = "Demo Mode"
    ElseIf gIsTestWebServer() Then
        lbcWebType.Caption = "Test Website"
    End If
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    ilRet = gPopAdvertisers()
    ilRet = gPopVff()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    ilRet = gPopAvailNames()
    ilRet = gPopVehicleOptions()
    imIdx = 0
    imFtpInProgress = False
    ReDim tmWebQueue(0 To 0) As WEBINFO
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    frmWebExportSchdSpot.Caption = "Counterpoint Affidavit - " & sgClientName
    '10/20/17: remove setting the default date
    'edcDate.Text = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    'debug
    'edcDate.Text = "2/13/17"
    edcDate.Text = ""
    imNumberDays = 7
    txtNumberDays.Text = Trim$(Str$(imNumberDays))
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imWaiting = False
    imFirstTime = True
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    lbcStation.Clear
    ReDim imExportedVefArray(0 To 0) As Integer
    Call gLoadOption(slSection, "DBPath", sgDBPath)
    sgDBPath = gSetPathEndSlash(sgDBPath, True)
    If sgDBPath = "Not Found" Then
        gMsgBox "FAIL: Affiliat.ini missing DBPath = under [Locations]"
        gMsgBox "Warning: No Copy Rotation Comments Will be Exported"
    End If
    Call gLoadOption(sgWebServerSection, "WebExports", smWebExports)
    smWebExports = gSetPathEndSlash(smWebExports, True)
    Call gLoadOption(sgWebServerSection, "WebImports", smWebImports)
    smWebImports = gSetPathEndSlash(smWebImports, True)
    txtCallLetters.Visible = False
    lblCallLetters.Visible = False
    gSetMousePointer grdVeh, grdVeh, vbDefault
    Call gLoadOption(sgWebServerSection, "ReIndex", smCheckIniReIndex)
    Call gLoadOption(sgWebServerSection, "MinSpotsToReIndex", smMinSpotsToReIndex)
    If Not gLoadOption(sgWebServerSection, "WebImports", smWebImports) Then
        gMsgBox "Affiliat.Ini [WebServer] 'WebImports' key is missing.", vbCritical
    Exit Sub
    End If
    smWebImports = gSetPathEndSlash(smWebImports, True)
    smJelliExport = "N"
    SQLQuery = "Select safFeatures1 From SAF_Schd_Attributes WHERE safVefCode = 0"
    'Set rst = cnn.Execute(SQLQuery)
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue1 = Asc(rst!safFeatures1)
        If (ilValue1 And JELLIEXPORT) = JELLIEXPORT Then
            smJelliExport = "Y"
        End If
    End If
    '9851
    Set myRemapMonitor = New cRemapper
    myRemapMonitor.MonthlyMonitoring
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-Form_Load"
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim ilRet As Integer
    
    On Error Resume Next
    If imExporting Then
        Cancel = True
        imTerminate = True
        Exit Sub
    End If
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    imIdx = 0
    Erase tmCPDat
    Erase tmAstInfo
    Erase imExptVefCode
    Erase imExportedVefArray
    Erase lmLastattCode
    Erase mFtpArray
    Erase lmWebPostedAttRecs
    Erase lmPostedAttRecs
    Erase smStaFailedToExport
    Erase tmEsf
    Erase tmEdf
    Erase tmAirTimeInfo
    rst_Gsf.Close
    rst_DAT.Close
    rst_Est.Close
    Set myEnt = Nothing
    Set frmWebExportSchdSpot = Nothing
Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-Form_Unload"
    Set frmWebExportSchdSpot = Nothing
End Sub


Private Sub Frame1_DragDrop(Source As control, X As Single, Y As Single)

End Sub

Private Sub grdVeh_Click()

    'Move to mouse up
    'lbcStation.Clear
    'If mGetGrdSelCount() = 1 Then
    '    edcTitle3.Visible = True
    '    chkAllStation.Visible = True
        lbcStation.Visible = True
    '    mFillStations D.S. 11/8/19 commented out
    'Else
    '    edcTitle3.Visible = False
    '    chkAllStation.Visible = False
    '    lbcStation.Visible = False
    'End If
    imBypassAll = True
    chkAll.Value = vbUnchecked
    imBypassAll = False
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

Private Sub lbcFilter_Click()
    tmcFilterDelay.Enabled = True
End Sub

Private Sub lbcFilter_GotFocus()
    tmcFilterDelay.Enabled = False
End Sub

Private Sub lbcFilter_LostFocus()
    tmcFilterDelay_Timer
    
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
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
    
    Exit Sub
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-lbcStation_Click"
    Exit Sub

End Sub

Private Sub rbcFilter_Click(Index As Integer)

    Dim ilRet As Integer
    
    If rbcFilter(Index).Value Then
        Select Case Index
            Case 0
                mPopDMA
                smFilterType = "DMA"
            Case 1
                mPopFormat
                smFilterType = "Format"
            Case 2
                mPopMSA
                smFilterType = "MSA"
            Case 3
                mPopState
                smFilterType = "State"
            Case 4
                smFilterType = "Station"
        End Select
        mFillStations
    End If
End Sub

Private Sub rbcVehicles_Click(Index As Integer)
    If rbcVehicles(Index).Value Then
        '7/1/20: Forec re-population of vehicles
        smEDCDate = ""
        mFillVehicle
    End If
End Sub

Private Sub edcDate_Change()
    '8163
    tmcDelay.Enabled = False
    tmcDelay.Interval = 3000
    lbcMsg.Clear
    'If cmdExport.Enabled = False Then
    '    cmdExport.Enabled = True
    '    cmdExport.Enabled = True
    '    cmdCancel.Caption = "&Cancel"
    'End If
    tmcDelay.Enabled = True
End Sub

Private Function mExportSpots() As Integer
    
    Dim slMsg As String
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim iUpper As Integer
    Dim ilOkStation As Integer
    Dim ilWriteHeader As Boolean
    Dim cprst As ADODB.Recordset
    Dim rst_Ast1 As ADODB.Recordset
    Dim ilRet As Integer
    Dim slPDate As String
    Dim llSpotsPosted As Long
    Dim slMoDate As String
    Dim slSuDate As String
    Dim llDelRecs As Long
    Dim llAddRecs As Long
    Dim llTtlRecs As Long
    Dim ilShowAlert As Integer
    Dim slNowTime As String
    Dim slNowDate As String
    Dim slErrMsg As String
    Dim slTemp As String
    Dim slFirstZeroAst As Boolean
    Dim ilMaxTries As Integer
    Dim llVpf As Long
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim ilIdx As Integer
    Dim ilStatus As Integer
    '11/28/17: Retain current task block status
    Dim blTaskBlocker As Boolean
    Dim blBpassHdSpot As Boolean
            
    On Error GoTo ErrHand
    imFailures = False
    '11/3/17
    imVefCode = imExptVefCode(0)
    smVefName = gGetVehNameByVefCode(imVefCode)
    
    Do
        bgIllegalCharsFound = False
        For ilVef = 0 To UBound(imExptVefCode) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            'Get CPTT so that Stations requiring CP can be obtained
            SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attExportToWeb, attMonthlyWebPost, attLoad, attNoAirPlays "
            SQLQuery = SQLQuery & " FROM shtt, cptt, att "
            SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
            SQLQuery = SQLQuery & " AND attCode = cpttAtfCode "
            SQLQuery = SQLQuery & " AND attExportType = 1"
            SQLQuery = SQLQuery & " AND attExportToWeb = 'Y'"
            SQLQuery = SQLQuery & " AND cpttVefCode = " & imExptVefCode(ilVef)  'imVefCode
            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(smMODate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by shttcallletters"
            'Set cprst = cnn.Execute(SQLQuery)
            Set cprst = gSQLSelectCall(SQLQuery)
            Call gPopVehicleOptions
            txtCallLetters.Text = ""
            'D. S. 1/29/18
            Do While Not cprst.EOF
                slFirstZeroAst = False
                blBpassHdSpot = False
                If imTerminate Then
                    gSetMousePointer grdVeh, grdVeh, vbDefault
                    cmdCancel.Enabled = True
                    imExporting = False
                    SetResults "** User Terminated Export **", RGB(255, 0, 0)
                    gLogMsg "** User Terminated Export **", "WebExportLog.Txt", False
                    Exit Function
                End If
                If igExportSource = 2 Then DoEvents
                lmAttCode = cprst!attCode
                If gIsVendorWithAgreement(lmAttCode, Vendors.Jelli) Then
                    smAttExportToJelli = "Y"
                Else
                    smAttExportToJelli = ""
                End If
                slErrMsg = ""
                If lbcStation.ListCount > 0 Then
                    ilOkStation = False
                    For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                        If lbcStation.Selected(iLoop) Then
                            If lbcStation.ItemData(iLoop) = cprst!shttCode Then
                                ilOkStation = True
                                imShttCode = cprst!shttCode
                                Exit For
                            End If
                        End If
                    Next iLoop
                Else
                    ilOkStation = True
                    imShttCode = cprst!shttCode
                End If
                If ilOkStation Then
                    If sgWebExport = "W" Then
                        If cprst!attExportToWeb <> "Y" Then
                            ilOkStation = False
                        End If
                    ElseIf sgWebExport = "B" Then
                        If cprst!attExportToWeb <> "Y" Then
                            ilOkStation = False
                        End If
                    End If
                End If
                If ilOkStation Then
                    If Not cprst.EOF Then
                        imLoadMultiplier = cprst!attLoad
                        imNoAirPlays = cprst!attNoAirPlays
                    Else
                        imLoadMultiplier = 1
                        imNoAirPlays = 1
                    End If
                    If imLoadMultiplier <= 1 Then
                        ReDim tmAirTimeInfo(0 To imNoAirPlays) As AIRTIMEINFO
                    Else
                        ReDim tmAirTimeInfo(0 To imLoadMultiplier) As AIRTIMEINFO
                    End If
                    For ilIdx = 0 To UBound(tmAirTimeInfo) Step 1
                        tmAirTimeInfo(ilIdx).sBaseStartDate = ""
                        tmAirTimeInfo(ilIdx).sBaseStartTime = ""
                        tmAirTimeInfo(ilIdx).sEstimatedEndTime = ""
                        tmAirTimeInfo(ilIdx).sEstimatedStartTime = ""
                        tmAirTimeInfo(ilIdx).sPrevPldgETime = ""
                        tmAirTimeInfo(ilIdx).sPrevPldgSTime = ""
                    Next ilIdx
                    ilWriteHeader = True
                    slSuDate = gAdjYear(DateAdd("d", 6, smMODate))
                    llTtlRecs = DateDiff("d", smMODate, slSuDate)
                    llTtlRecs = llTtlRecs + 1
                    ilShowAlert = False
                    lmWebAttID = gGetLogAttID(imExptVefCode(ilVef), imShttCode, lmAttCode)
                    'D.S. 11/4/12  Get alternate vehicle name
                    smShowVehName = ""
                    If lmWebAttID <> lmAttCode Then
                        ilVff = gBinarySearchVff(imExptVefCode(ilVef))
                        If ilVef = -1 Then
                            ilRet = gPopVff()
                            ilVff = gBinarySearchVff(imExptVefCode(ilVef))
                        End If
                        If ilVff <> -1 Then
                            smShowVehName = Trim$(tgVffInfo(ilVff).sWebName)
                        End If
                    End If
                    SQLQuery = "SELECT Count(datCode) "
                    SQLQuery = SQLQuery + " FROM dat"
                    SQLQuery = SQLQuery + " WHERE (datAtfCode = " & lmAttCode
                    SQLQuery = SQLQuery + " AND datEstimatedTime = " & "'" & "Y" & "')"
                    'Set rst_DAT = cnn.Execute(SQLQuery)
                    Set rst_DAT = gSQLSelectCall(SQLQuery)
                    If IsNull(rst_DAT(0).Value) Or rst_DAT(0).Value = 0 Then
                        lmEstimatesExist = False
                    Else
                        lmEstimatesExist = True
                    End If
                    If igExportSource = 2 Then DoEvents
                    'Check the Ast file to see if any spots have been posted in the requested time frame.
                    'If so don't export the station.
                    
                    SQLQuery = "Select astCode, astStatus FROM ast WHERE "
                    SQLQuery = SQLQuery + " astAtfCode = " & cprst!attCode
                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(smStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(smEndDate, sgSQLDateForm) & "')"
                    '2/21/18: Check for received and None Aired (0 = Not Received; 1= Received; 2= None Aired)
                    'SQLQuery = SQLQuery + " AND astCPStatus = 1"
                    SQLQuery = SQLQuery + " AND (astCPStatus = 1 or astCPStatus = 2)"
                    'D.S. 05/25/17 Commented out line below. It's way too slow and the 2 cases should never happen
                    'SQLQuery = SQLQuery + " AND (Mod(astStatus, 100) <= 10) or (astStatus = " & ASTAIR_MISSED_MG_BYPASS & ")"
                    'Set rst_Ast1 = cnn.Execute(SQLQuery)
                    Set rst_Ast1 = gSQLSelectCall(SQLQuery)
                    If igExportSource = 2 Then DoEvents
                    llSpotsPosted = False
                    If rst_Ast1.EOF Then
                        'nothing posted on affiliate so, check the web
                        llSpotsPosted = gBinarySearchWebPostedAttCodes(cprst!cpttatfCode)
                        'D.S. TTP 9849 shouldn't need here but...
                        If llSpotsPosted <> -1 And Not mIsServiceAgreement(cprst!cpttatfCode) Then
                            llSpotsPosted = True
                            slErrMsg = " Rejected: Posting found on the Web."
                            ilShowAlert = True
                        Else
                            llSpotsPosted = False
                        End If
                    Else
                        'D.S. 12/8/10 we have spots posted on the affiliate.  If they are not an airing post ie "Not Carried" then send
                        'out the spots.  12/8/10
                        Do While Not rst_Ast1.EOF
                            If igExportSource = 2 Then DoEvents
                                '2/21/18: Any status other then Not Carried should be checked.
                                'If tgStatusTypes(gGetAirStatus(rst_Ast1!astStatus)).iPledged <> 2 Then
                                If tgStatusTypes(gGetAirStatus(rst_Ast1!astStatus)).iStatus <> 8 Then   '8=Not Carried
                                    '3/7/16: Bypass MG and replacement spots (Spot imported that was MG in week to export, pledged in week 1 MG in Week 2 and week 1 import)
                                    ilStatus = rst_Ast1!astStatus Mod 100
                                    If (ilStatus < ASTEXTENDED_MG) Or ((sgMissedMGBypass = "Y") And (ilStatus = ASTAIR_MISSED_MG_BYPASS)) Then
                                        'D.S. TTP  9849 5/21/20 Don't show zero spots exported for service agreements. They are pre-posted
                                        If Not mIsServiceAgreement(cprst!cpttatfCode) Then
                                            llSpotsPosted = True
                                            slErrMsg = " Rejected: Posting found on the Network "
                                            ilShowAlert = True
                                        End If
                                        'D. S. 1/29/18
                                        Exit Do
                                    End If
                                End If
                            rst_Ast1.MoveNext
                        Loop
                    End If
                    '2/21/18: Moved below
                    If llSpotsPosted Then
                        'We found a posted record add it to the list of stations not exported
                        If ilOkStation Then
                            '7458
                            myEnt.ClearWhenDontSend
                            If igExportSource = 2 Then DoEvents
                            llAddRecs = 0   'UBound(tmAstInfo)
                            smStaFailedToExport(lmFailedStaExpIdx) = Trim$(cprst!shttCallLetters) & ", " & smVefName & " " & slErrMsg
                            lmFailedStaExpIdx = lmFailedStaExpIdx + 1
                            ReDim Preserve smStaFailedToExport(0 To lmFailedStaExpIdx)
                            imFailedStations = True
                            ilShowAlert = True
                        End If
                    End If
                    txtCallLetters.Visible = True
                    lblCallLetters.Visible = True
                    'If ilOkStation Then
                    If ilOkStation And (Not llSpotsPosted) Then
                        ' Check to see if this header has been output yet or not.
                        iUpper = UBound(lmLastattCode)
                        For iLoop = 0 To iUpper - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            'If lmLastattCode(iLoop) = cprst!attCode Then
                            If lmLastattCode(iLoop) = lmWebAttID Then
                                ' This header has already been written. Don't write it out again.
                                ilWriteHeader = False
                                Exit For
                            End If
                        Next iLoop
                    Else
                        ilOkStation = False
                    End If
                    
                    'If ilOkStation Then
                    If ilOkStation And (Not llSpotsPosted) Then
                        txtCallLetters.Text = Trim$(cprst!shttCallLetters)
                        'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                        tgCPPosting(0).iStatus = cprst!cpttStatus
                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                        tgCPPosting(0).iVefCode = imExptVefCode(ilVef)    'imVefCode
                        tgCPPosting(0).iShttCode = cprst!shttCode
                        tgCPPosting(0).sZone = cprst!shttTimeZone
                        tgCPPosting(0).sDate = Format$(smMODate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                        igTimes = 1 'By Week
                        If igExportSource = 2 Then DoEvents
                        iRet = False
                        ilMaxTries = 0
                        '11/28/17: Retain current task block status
                        blTaskBlocker = bgTaskBlocked
                        
                        Do While (Not iRet) And (ilMaxTries < 3)
                            If igExportSource = 2 Then DoEvents
                            ilMaxTries = ilMaxTries + 1
                            lgSTime2 = timeGetTime
                            With myEnt
                                .Vehicle = imVefCode
                                .Station = cprst!shttCode
                                .Agreement = cprst!cpttatfCode
                                .fileName = smWebSpots
                                .ProcessStart
                                .SetThirdPartyByHierarchy
                            End With
                            '11/28/17: Set current task block status
                            bgTaskBlocked = False
                            iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True)
                            '11/28/17: Save current task block status
                            If bgTaskBlocked Then
                                If Not blTaskBlocker Then
                                    SetResults "Task Blocked, not all station spots generated/exported", RGB(255, 0, 0)
                                End If
                                blTaskBlocker = bgTaskBlocked
                                blBpassHdSpot = True
                                Exit Do
                            End If
                            gFilterAstExtendedTypes tmAstInfo
                            lgETime2 = timeGetTime
                            lgTtlTime2 = lgTtlTime2 + (lgETime2 - lgSTime2)
                            'This call should only be made if a problem exist
                            If Not iRet Then
                                gClearASTInfo True
                            End If
                        'Wend
                        Loop
                        If Not blBpassHdSpot Then
                            If iRet = False Then
                                gLogMsg "Error: gGetAstInfo Retries Exceeded - Notify Counterpoint", "WebExportLog.Txt", False
                            End If
                            If ((tmAstInfo(0).lCode = 0) And UBound(tmAstInfo) > 0) Or (Not iRet) Then
                                If igExportSource = 2 Then DoEvents
                                '2/6/18
                                'ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
                                iRet = gCloseMKDFile(hmAst, "Ast.Mkd")
                                If iRet Then
                                    iRet = gOpenMKDFile(hmAst, "Ast.Mkd")
                                End If
                                gClearASTInfo True
                                If iRet Then
                                    lgSTime2 = timeGetTime
                                    '11/28/17: set task block status
                                    bgTaskBlocked = False
                                    iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True)
                                    '11/28/17: Save current task block status
                                    If bgTaskBlocked Then
                                        If Not blTaskBlocker Then
                                            SetResults "Task Blocked, not all station spots generated/exported", RGB(255, 0, 0)
                                        End If
                                        blTaskBlocker = bgTaskBlocked
                                        blBpassHdSpot = True
                                    Else
                                        gFilterAstExtendedTypes tmAstInfo
                                        lgETime2 = timeGetTime
                                        lgTtlTime2 = lgTtlTime2 + (lgETime2 - lgSTime2)
                                        If ((tmAstInfo(0).lCode = 0) And UBound(tmAstInfo) > 0) Or (Not iRet) Then
                                            slFirstZeroAst = True
                                            imFailures = True
                                            SetResults "Bad AST Code - Notify Counterpoint", RGB(255, 0, 0)
                                            gLogMsg "Error: Bad Spot AST Code - Notify Counterpoint", "WebExportLog.Txt", False
                                            gLogMsg "Error: Start Date: " & smMODate & ", Number of days: " & Trim$(txtNumberDays.Text) & ", Vehicle: " & smVefName & ", Station: " & Trim$(cprst!shttCallLetters), "WebExportLog.Txt", False
                                        End If
                                    End If
                                Else
                                    SetResults "Failed to Open AST file.", RGB(255, 0, 0)
                                    gLogMsg "Error: Failed to Open AST file.", "WebExportLog.Txt", False
        
                                End If
                            End If
                        End If
                        '11/28/17: Restore current task block status
                        bgTaskBlocked = blTaskBlocker
                        If Not blBpassHdSpot Then
                            smUseActual = "N"
                            llVpf = gBinarySearchVpf(CLng(tgCPPosting(0).iVefCode))
                            If (llVpf = -1) Then
                                gPopVehicleOptions
                                llVpf = gBinarySearchVpf(CLng(tgCPPosting(0).iVefCode))
                            End If
                            If (llVpf <> -1) Then
                                If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
                                    smUseActual = "Y"
                                End If
                            End If
                            smSuppressLog = "N"
                            If (llVpf <> -1) Then
                                If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And SUPPRESSWEBLOG) = SUPPRESSWEBLOG Then
                                    smSuppressLog = "Y"
                                End If
                            End If
                            If ilWriteHeader Then
                                smMode = "A"
                                If igExportSource = 2 Then DoEvents
                                iRet = mBuildHeaders(cprst)
                            End If
                            If igExportSource = 2 Then DoEvents
                            If Not llSpotsPosted Then
                                'Output AST
                                If Not slFirstZeroAst Then
                                    lgSTime4 = timeGetTime
                                    iRet = mBuildDetailRecs(cprst, tmAstInfo(), "A", smEndDate)
                                    lgETime4 = timeGetTime
                                    lgTtlTime4 = lgTtlTime4 + lgETime4 - lgSTime4
                                End If
                                If imSpotCount > 0 Then
                                    If slMsg = "" Then
                                        slMsg = Trim$(cprst!shttCallLetters) & ", "
                                    End If
                                    'D.S. 2/9/18
                                    'D.S. 2/12/18
                                    'If blTaskBlocker Then
                                        slMsg = slMsg & CStr(imSpotCount) & " Add Spots."
                                        lmFileAddSpotCount = lmFileAddSpotCount + imSpotCount
                                        llAddRecs = imSpotCount
                                        mExportedVefArray imExptVefCode(ilVef)
                                    'End If
                                Else
                                    If Not slFirstZeroAst Then
                                        ilShowAlert = True
                                        slErrMsg = "Zero Spots. See Agreement."
                                    End If
                                End If
                                'Delete Cptt records if necessary
                                If Not slFirstZeroAst Then
                                    iRet = mAdjCpttRecs(cprst)
                                End If
                            End If
                            '2/21/15: Moved above
                            'If llSpotsPosted Then
                            '    'We found a posted record add it to the list of stations not exported
                            '    If ilOkStation Then
                            '        '7458
                            '        myEnt.ClearWhenDontSend
                            '        If igExportSource = 2 Then DoEvents
                            '        llAddRecs = UBound(tmAstInfo)
                            '        smStaFailedToExport(lmFailedStaExpIdx) = Trim$(cprst!shttCallLetters) & ", " & smVefName & " " & slErrMsg
                            '        lmFailedStaExpIdx = lmFailedStaExpIdx + 1
                            '        ReDim Preserve smStaFailedToExport(0 To lmFailedStaExpIdx)
                            '        imFailedStations = True
                            '        ilShowAlert = True
                            '    End If
                            'End If
                        End If
                    End If
                End If
                If ilOkStation And (Not slFirstZeroAst) And (Not blBpassHdSpot) Then
                    If igExportSource = 2 Then DoEvents
                    'Time Zones
                    slTemp = Trim$(cprst!shttTimeZone)
                    Select Case slTemp
                        Case "EST"
                        Case "CST"
                        Case "MST"
                        Case "PST"
                        Case Else
                            If ilShowAlert Then
                                slErrMsg = slErrMsg & ", No Zone."
                            Else
                                slErrMsg = "Missing Zone."
                                ilShowAlert = True
                            End If
                    End Select
                    slNowTime = Format$(gNow(), sgSQLTimeForm)
                    slNowDate = Format$(gNow(), sgSQLDateForm)
                    'Add new detail record
                    SQLQuery = "Insert Into EDF_Export_Detail ( "
                    SQLQuery = SQLQuery & "edfCode, "
                    SQLQuery = SQLQuery & "edfEsfCode, "
                    SQLQuery = SQLQuery & "edfAttCode, "
                    SQLQuery = SQLQuery & "edfVefCode, "
                    SQLQuery = SQLQuery & "edfShttCode, "
                    SQLQuery = SQLQuery & "edfStartTime, "
                    SQLQuery = SQLQuery & "edfStartDate, "
                    SQLQuery = SQLQuery & "edfTtlAdd, "
                    SQLQuery = SQLQuery & "edfTtlDel, "
                    SQLQuery = SQLQuery & "edfTtlAddDel, "
                    SQLQuery = SQLQuery & "edfUser, "
                    SQLQuery = SQLQuery & "edfAlert "
                    SQLQuery = SQLQuery & ") "
                    SQLQuery = SQLQuery & "Values ( "
                    SQLQuery = SQLQuery & 0 & ", "
                    SQLQuery = SQLQuery & lmMaxEsfCode & ", "
                    SQLQuery = SQLQuery & CLng(cprst!cpttatfCode) & ", "
                    SQLQuery = SQLQuery & imExptVefCode(ilVef) & ", "
                    SQLQuery = SQLQuery & CInt(cprst!shttCode) & ", "
                    SQLQuery = SQLQuery & "'" & Format$(slNowTime, sgSQLTimeForm) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(smStartDate, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & llAddRecs & ", "
                    SQLQuery = SQLQuery & llDelRecs & ", "
                    SQLQuery = SQLQuery & llTtlRecs & ", "
                    SQLQuery = SQLQuery & "'" & gFixQuote(sgUserName) & "', "
                    If ilShowAlert Then
                        SQLQuery = SQLQuery & "'" & slErrMsg & "'"
                    Else
                        SQLQuery = SQLQuery & "'" & "" & "'"
                    End If
                    SQLQuery = SQLQuery & ") "
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    llAddRecs = 0
                    llDelRecs = 0
                    llTtlRecs = 0
                    ilShowAlert = False
                    If igExportSource = 2 Then DoEvents
                    If Not myEnt.CreateEnts(Incomplete) Then
                        gLogMsg "Failed to Create ENT record.", "WebExportLog.Txt", False
                        llTtlRecs = llTtlRecs
                    End If
                End If
                If slMsg <> "" Then
                    gLogMsg slMsg, "WebExportLog.Txt", False
                    slMsg = ""
                End If
                '1/11/21: Move here for below and change fields. TTP 10059
                If (Len(slErrMsg) > 0) And (cprst!attCode > 0) Then
                    If Not mIsServiceAgreement(cprst!attCode) Then
                        gLogMsg Trim$(cprst!shttCallLetters) & "  on: " & smVefName & " Station: " & slErrMsg & " Date Range " & smStartDate & " - " & smEndDate, "StationsNotExported.Txt", False
                    End If
                End If
                If igExportSource = 2 Then DoEvents
                cprst.MoveNext
            Loop
            If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
                gClearASTInfo True
            Else
                gClearASTInfo False
            End If
            '1/11/21: Moved within agreement loop. place just above MoveNext. TTP 10059
            'If Len(slErrMsg) > 0 Then
            '    If Not mIsServiceAgreement(tgCPPosting(0).lAttCode) Then
            '        gLogMsg Trim$(txtCallLetters.Text) & "  on: " & smVefName & " Station: " & slErrMsg & " Date Range " & smStartDate & " - " & smEndDate, "StationsNotExported.Txt", False
            '    End If
            'End If
            '12/11/17: Clear abf
            If (lbcStation.ListCount <= 0) Or (chkAllStation.Value = vbChecked) Then
                gClearAbf imExptVefCode(ilVef), 0, smMODate, gObtainNextSunday(smMODate), False
            End If
        Next ilVef
        smMODate = DateAdd("d", 7, smMODate)
    Loop While DateValue(gAdjYear(smMODate)) < DateValue(gAdjYear(smEndDate))
    'reset the date back for the next vehicle
    smMODate = gObtainPrevMonday(smStartDate)
    smEndDate = DateAdd("d", imNumberDays - 1, smStartDate)
    mExportSpots = True
    Exit Function

ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mExportSpots"
End Function

Private Sub mFillStations(Optional blFillStations As Boolean = False)
    If blStationListAlreadyDrawn = True Then Exit Sub
    Dim ilRet As Integer
    Dim llVef As Long
    Dim ilVff As Integer
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilFilterIdx As Integer
    Dim slDate As String
    Dim slPrevCallLetters As String
    Dim ilFound As Integer
    Dim slTemp As String
    
    On Error GoTo ErrHand
    
    lbcStation.Clear
    If chkAllStation.Value = vbChecked Then
        lbcStation.Visible = False
        lacStationMsg.Visible = True
    End If
    
    If mGetGrdSelCount() <= 0 Then
        mSetCommands
        Exit Sub
    End If
    
    If rbcFilter(4).Value = False Then
        If lbcFilter.SelCount <= 0 Then
            mSetCommands
            Exit Sub
        End If
    End If
    
    '11/8/19: Fill only during export and all unchecked
    If (rbcFilter(4).Value = True) And (blFillStations = False) And (chkAllStation.Value = vbChecked) Then
        mSetCommands
        Exit Sub
    End If
    bmInStationFill = True
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    'DoEvents
    'Only get agreements that are to be sent to web and are active
    If edcDate.Text = "" Then
        slDate = Format(gNow(), sgSQLDateForm)
    ElseIf gIsDate(edcDate.Text) = False Then
        slDate = Format(gNow(), sgSQLDateForm)
    Else
        slDate = Format(edcDate.Text, sgSQLDateForm)
    End If
    If Not rst.EOF Then
        lbcStation.Clear
    End If
    lbcStation.Visible = False
    'If mGetGrdSelCount() > 0 Then
        For ilLoop = 1 To grdVeh.Rows - 1 Step 1
            'DoEvents
            If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
                If grdVeh.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                    imVefCode = grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)
                    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode, shttState, shttMktCode, shttMarket,shttMetCode, shttFmtCode, attVefCode"
                    SQLQuery = SQLQuery & " FROM shtt, att"
                    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
                    'SQLQuery = SQLQuery & " WHERE (attExportType = 1"
                    SQLQuery = SQLQuery & " AND  attDropDate > '" & slDate & "' AND attOffAir > '" & slDate & "' AND attExportToWeb = 'Y'"
                    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
                    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
                    Set rst = gSQLSelectCall(SQLQuery)
    
                   While Not rst.EOF
                        '11/8/19: Replaced SendMessage call with BinarySearch which is much faster
                        'have we already added the call letters?
                        'llRow = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, Trim$(rst!shttCallLetters))
                        llRow = gBinarySearchListCtrl(lbcStation, Trim$(rst!shttCallLetters))
                        ilFound = 1
                        If llRow < 0 Then
                            If rbcFilter(0).Value = True Then   'DMA
                                For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                                    If lbcFilter.Selected(ilFilterIdx) Then
                                        If Trim(lbcFilter.List(ilFilterIdx)) = Trim$(rst!shttMarket) Then
                                            lbcStation.AddItem Trim$(rst!shttCallLetters)
                                            lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                        End If
                                    End If
                                Next ilFilterIdx
                            End If
                            If rbcFilter(1).Value = True Then    'Format
                                slTemp = mGetFormat(rst!shttFmtCode)
                                For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                                    If lbcFilter.Selected(ilFilterIdx) Then
                                        If Trim(lbcFilter.List(ilFilterIdx)) = Trim(slTemp) Then
                                            lbcStation.AddItem Trim$(rst!shttCallLetters)
                                            lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                        End If
                                    End If
                                Next ilFilterIdx
                            End If
                            If rbcFilter(2).Value = True Then   'MSA
                                slTemp = mGetMSA(rst!shttMetCode)
                                For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                                    If lbcFilter.Selected(ilFilterIdx) Then
                                        If Trim(lbcFilter.List(ilFilterIdx)) = Trim$(slTemp) Then
                                            lbcStation.AddItem Trim$(rst!shttCallLetters)
                                            lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                        End If
                                    End If
                                Next ilFilterIdx
                            End If
                            If rbcFilter(3).Value = True Then    'State
                                For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                                    If lbcFilter.Selected(ilFilterIdx) Then
                                        If Left(lbcFilter.List(ilFilterIdx), 2) = Trim$(rst!shttState) Then
                                            lbcStation.AddItem Trim$(rst!shttCallLetters)
                                            lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                        End If
                                    End If
                                Next ilFilterIdx
                            End If
                            If rbcFilter(4).Value = True Then   'Stations
                                lbcStation.AddItem Trim$(rst!shttCallLetters)
                                lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                            End If
                        End If
                        rst.MoveNext
                    Wend
                End If
            End If
        Next ilLoop
    'End If
    'If Log vehicle, check all vehicles that are part of that log vehicle
    llVef = gBinarySearchVef(CLng(imVefCode))
    If llVef = -1 Then
        llVef = gPopSellingVehicles()
        llVef = gBinarySearchVef(CLng(imVefCode))
    End If
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
                                'Set rst = cnn.Execute(SQLQuery)
                                Set rst = gSQLSelectCall(SQLQuery)
                                While Not rst.EOF
                                    'Check that name has not been previously added
                                    'llRow = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, Trim$(rst!shttCallLetters))
                                    llRow = gBinarySearchListCtrl(lbcStation, Trim$(rst!shttCallLetters))
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
    'chkAllStation.Value = vbChecked
    bmInStationFill = False
    gSetMousePointer grdVeh, grdVeh, vbDefault
    blStationListAlreadyDrawn = True
    chkAllStation_Click
    
    blStationListAlreadyDrawn = False
    lbcStation.Visible = True
    mSetCommands
    Exit Sub
ErrHand:
    bmInStationFill = False
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mFillStations"
    Exit Sub
End Sub

Private Function mValidateUserInput() As Integer

    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilRet As Integer
    
    ReDim tmEsf(0 To 0) As ESF
    ReDim tmEdf(0 To 0) As EDF
    On Error GoTo ErrHand
    'Debug
    'txtDate.text = "5/18/09"
    If edcDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        mValidateUserInput = False
        Exit Function
    End If
    If gIsDate(edcDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        mValidateUserInput = False
        Exit Function
    Else
        smStartDate = Format(edcDate.Text, sgShowDateForm)
    End If
    tmEsf(0).sExpDate = smStartDate
    slNowTime = Format$(gNow(), sgSQLTimeForm)
    tmEsf(0).sStartTime = slNowTime
    slNowDate = Format$(gNow(), sgSQLDateForm)
    tmEsf(0).sStartDate = slNowDate
    imNumberDays = Val(txtNumberDays.Text)
    tmEsf(0).iNumDays = imNumberDays
    smMODate = gObtainPrevMonday(smStartDate)
    smEndDate = DateAdd("d", imNumberDays - 1, smStartDate)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        mValidateUserInput = False
        Exit Function
    End If
    tmEsf(0).sUser = sgUserName
    tmEsf(0).sMachine = gGetComputerName
    mValidateUserInput = True
    Exit Function
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mValidateUserInput"
    Exit Function
End Function

Private Function mOpenFiles() As Integer

    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slTemp As String
    
    imExporting = True
    ilRet = 0
    On Error GoTo cmdExportErr:
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & ".txt"
    slTemp = slTemp & "_" & sgUserName & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    smWebSpots = "WebSpots_" & slTemp
    smToFileDetail = smWebExports & smWebSpots
    smWebHeader = "WebHeaders_" & slTemp
    smToFileHeader = smWebExports & smWebHeader
    smWebCopyRot = "CpyRotCom_" & slTemp
    smToCpyRot = smWebExports & smWebCopyRot
    smWebMultiUse = "MultiUse_" & slTemp
    smToMultiUse = smWebExports & smWebMultiUse
    'D.S. Check bit map to see if using games.  If not no sense in exporting of showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
    smWebEventInfo = "EventInfo_" & slTemp
    smToEventInfo = smWebExports & smWebEventInfo
    End If
    ilRet = gFileExist(smToFileDetail)
    If ilRet = 0 Then
        gSetMousePointer grdVeh, grdVeh, vbDefault
        ilRet = gMsgBox("Export Previously Created " & slDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        If ilRet = vbCancel Then
            gLogMsg "** Terminated Because Export File Existed **", "WebExportLog.Txt", False
            Close #hmToDetail
            imExporting = False
            mOpenFiles = False
            Exit Function
        End If
        gSetMousePointer grdVeh, grdVeh, vbHourglass
        Kill smToFileDetail
        Kill smToFileHeader
        Kill smToCpyRot
        Kill smToMultiUse
    End If
    On Error GoTo 0
    ilRet = 0
    On Error GoTo cmdExportErr:
    hmToDetail = FreeFile
    Open smToFileDetail For Output Lock Write As hmToDetail
    If ilRet <> 0 Then
        gLogMsg "** Terminated - " & smToFileDetail & " failed to open. **", "WebExportLog.Txt", False
        Close #hmToDetail
        imExporting = False
        gSetMousePointer grdVeh, grdVeh, vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & smToFileDetail, vbOKOnly, "Open Error"
        mOpenFiles = False
        Exit Function
    End If
    hmToHeader = FreeFile
    Open smToFileHeader For Output Lock Write As hmToHeader
    If ilRet <> 0 Then
        gLogMsg "** Terminated - " & smToFileHeader & " failed to open. **", "WebExportLog.Txt", False
        Close #hmToDetail
        Close #hmToHeader
        imExporting = False
        gSetMousePointer grdVeh, grdVeh, vbDefault
        gMsgBox "Open Error #" & Str$(Err.Number) & smToFileHeader, vbOKOnly, "Open Error"
        mOpenFiles = False
        Exit Function
    End If
    hmToCpyRot = FreeFile
    Open smToCpyRot For Output Lock Write As hmToCpyRot
    If ilRet <> 0 Then
        gLogMsg "** Terminated - " & smToCpyRot & " failed to open. **", "WebExportLog.Txt", False
        Close #hmToDetail
        Close #hmToHeader
        Close #hmToCpyRot
        imExporting = False
        gSetMousePointer grdVeh, grdVeh, vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & smToCpyRot, vbOKOnly, "Open Error"
        mOpenFiles = False
        Exit Function
    End If
    hmToMultiUse = FreeFile
    Open smToMultiUse For Output Lock Write As hmToMultiUse
    If ilRet <> 0 Then
        gLogMsg "** Terminated - " & smToMultiUse & " failed to open. **", "WebExportLog.Txt", False
        Close #hmToDetail
        Close #hmToHeader
        Close #hmToCpyRot
        Close #hmToMultiUse
        imExporting = False
        gSetMousePointer grdVeh, grdVeh, vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & smToMultiUse, vbOKOnly, "Open Error"
        mOpenFiles = False
        Exit Function
    End If
    'Check bit map to see if using games.  If not no sense in exporting or showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        hmToEventInfo = FreeFile
        Open smToEventInfo For Output Lock Write As hmToEventInfo
        If ilRet <> 0 Then
                gLogMsg "** Terminated - " & smToEventInfo & " failed to open. **", "WebExportLog.Txt", False
            Close #hmToDetail
            Close #hmToHeader
            Close #hmToCpyRot
            Close #hmToMultiUse
            Close #hmToEventInfo
            imExporting = False
            gSetMousePointer grdVeh, grdVeh, vbDefault
            gMsgBox "Open Error #" & Str$(Err.Numner) & smToEventInfo, vbOKOnly, "Open Error"
            mOpenFiles = False
            Exit Function
        End If
    End If
    If igExportSource = 2 Then DoEvents
    ' Print the headers
    Print #hmToDetail, "attCode, Advt, Prod, TranType, PledgeStartDate, PledgeEndDate, PledgeStartTime, PledgeEndTime, FeedDate, FeedTime, SpotLen, Cart, ISCI, CreativeTitle, astCode, CpyRotCode, AvailName, OrgStatusCode, gsfCode, RotEndDate, IsDaypart, EstimatedDay, EstimatedTime, BeforeOrAfter, TrueDaysPledged, ActualDateTime, srcAttCode, showVehName, FlightDays, CntrNumber, FlightStartTime, FlightEndTime, AdfCode, Blackout, EmbeddedOrROS"
    Print #hmToHeader, gBuildWebHeaderDetail()
    Print #hmToCpyRot, "Code, Comment"
    'Check bit map to see if using games.  If not no sense in exporting of showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
    Print #hmToEventInfo, "Code, GameDate, GameStartTime, VisitTeamName, VisitTeamAbbr, HomeTeamName, HomeTeamAbbr, LanguageCode, FeedSource, EventCarried, AttCode"
    End If
    gLogMsg "** Storing Output into " & smToFileDetail & " And " & smToFileHeader & "**", "WebExportLog.Txt", False
    mOpenFiles = True
    Exit Function
    
cmdExportErr:
    ilRet = Err
    Resume Next
End Function

Private Function mFTPFiles() As Integer

    Dim ilRet As Integer
    Dim ilRetry As Integer
    
    On Error GoTo ErrHand
    If Not imFTPIsOn Then
        mFTPFiles = True
        Exit Function
    End If
    mFTPFiles = False
    If igDemoMode Then
        mFTPFiles = True
        Exit Function
    End If
    mFTPFiles = False
    imFtpInProgress = True
    ReDim mFtpArray(0 To 0)
    gLogMsg "Sending Files to web site.", "WebActivityLog.Txt", False
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        SQLQuery = "SELECT wqfCode FROM WQF_Web_Queue WHERE wqfFileName = " & "'" & smWebEventInfo & "'"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            gLogMsg "FTP " & smWebEventInfo & " to web site.", "WebExportLog.Txt", False
            SetResults "FTP " & smWebEventInfo & " to web site.", 0
            ilRet = csiFTPFileToServer(Trim$(smWebEventInfo))
            mFtpArray(UBound(mFtpArray)) = Trim$(smWebEventInfo)
            ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
        End If
    End If
    SQLQuery = "SELECT wqfCode FROM WQF_Web_Queue WHERE wqfFileName = " & "'" & smWebCopyRot & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gLogMsg "FTP " & smWebCopyRot & " to web site.", "WebExportLog.Txt", False
        SetResults "FTP " & smWebCopyRot & " to web site.", 0
        ilRet = csiFTPFileToServer(Trim$(smWebCopyRot))
        mFtpArray(UBound(mFtpArray)) = Trim$(smWebCopyRot)
        ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
    End If
    
    SQLQuery = "SELECT wqfCode FROM WQF_Web_Queue WHERE wqfFileName = " & "'" & smWebMultiUse & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gLogMsg "FTP " & smWebMultiUse & " to web site.", "WebExportLog.Txt", False
        SetResults "FTP " & smWebMultiUse & " to web site.", 0
        ilRet = csiFTPFileToServer(Trim$(smWebMultiUse))
    mFtpArray(UBound(mFtpArray)) = Trim$(smWebMultiUse)
    ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
    End If
    SQLQuery = "SELECT wqfCode FROM WQF_Web_Queue WHERE wqfFileName = " & "'" & smWebHeader & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gLogMsg "FTP " & smWebHeader & " to web site.", "WebExportLog.Txt", False
        SetResults "FTP " & smWebHeader & " to web site.", 0
        ilRet = csiFTPFileToServer(Trim$(smWebHeader))
    mFtpArray(UBound(mFtpArray)) = Trim$(smWebHeader)
    ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
    End If
    SQLQuery = "SELECT wqfCode FROM WQF_Web_Queue WHERE wqfFileName = " & "'" & smWebSpots & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gLogMsg "FTP " & smWebSpots & " to web site.", "WebExportLog.Txt", False
        SetResults "FTP " & smWebSpots & " to web site.", 0
        ilRet = csiFTPFileToServer(Trim$(smWebSpots))
        mFtpArray(UBound(mFtpArray)) = Trim$(smWebSpots)
        ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
    End If
    'End FTP Time
    mFTPFiles = True
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mFTPFiles"
    Exit Function
End Function

Private Sub mWaitForWebLock()
    On Error GoTo ErrHandler
    Dim ilLoop As Integer
    Dim ilTotalMinutes As Integer
    Dim ilNotSaidWebServerWasBusy As Boolean
    Dim slLastMessage As String
    Dim slThisMessage As String
    Dim ilRow As Integer
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String

    ilNotSaidWebServerWasBusy = False
    slLastMessage = "Nothing"
    While 1
        If igExportSource = 2 Then DoEvents
        ilTotalMinutes = gStartWebSession("WebExportLog.Txt")
        If ilTotalMinutes = 0 Then
            'Start the Export Process
            gLogMsg "Web Session Started Successfully", "WebExportLog.Txt", False
            Exit Sub
        End If
        If Not ilNotSaidWebServerWasBusy Then
            ilNotSaidWebServerWasBusy = True
            SetResults "The Server is Busy. Standby...", 0
        End If
        If ilTotalMinutes > 1 Then
            slThisMessage = "  -Max wait time is " & Trim(Str(ilTotalMinutes)) & " Minutes."
        Else
            slThisMessage = "  -Max wait time is " & Trim(Str(ilTotalMinutes)) & " Minute."
        End If
        If slThisMessage <> slLastMessage Then
            ilRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slLastMessage)
            If lbcMsg.ListCount And ilRow >= 0 Then
                lbcMsg.RemoveItem ilRow
            End If
            lbcMsg.Refresh
            lbcMsg.AddItem slThisMessage
            slLastMessage = slThisMessage
            DoEvents
        End If
        ' Wait here for 15 seconds. This loop allows the cancel button to be pressed as well.
        For ilLoop = 0 To 15
            If imTerminate Then
                Exit Sub
            End If
            DoEvents
            Sleep (1000)   ' Wait 1 of a second
        Next
    Wend
    Exit Sub

ErrHandler:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in frmWebExportSchdSpot - mWaitForWebLock: "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            gLogMsg gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot - mWaitForWebLock: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        gLogMsg gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
    End If
    Exit Sub
End Sub

Private Function mInitiateExport() As Integer

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim llFileSize As Long
    Dim ilStartNewFiles As Integer
    Dim slEndDate As String
    Dim llLoop As Long
    Dim rst As ADODB.Recordset
    Dim ilRetries As Integer
    Dim ilInsertFailed As Integer
    Dim llDupBuffSize As Long
    Dim ilTargetFileSizeInMB As Integer
    Dim llTargetFileValue As Long
    Dim ilFileSent As Boolean
    Dim llMinSpotsToReIndex As Long
    Dim ilIdx As Integer
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilVff As Integer
    
    lgSTime25 = timeGetTime
    smAttWebInterface = ""
    ilFileSent = False
    mInitiateExport = False
    ilRet = mInitFTP()
    'Init Affiliate var counts
    lmTtlHeaders = 0
    lmTtlComments = 0
    lmTtlMultiUse = 0
    lmTotalAddSpotCount = 0
    lmTotalDeleteSpotCount = 0
    'Init Web var counts
    lmWebTtlHeaders = 0
    lmWebTtlEventSpots = 0
    lmWebTtlComments = 0
    lmWebTtlMultiUse = 0
    lmWebTtlSpots = 0
    lmWebTtlEmail = 0
    ReDim tgCopyRotInfo(0 To 0) As CPYROTCOM
    ReDim tgGameInfo(0 To 0) As GAMEINFO
    
    'D.S. 6/14/05 The constant, clTargetFileSize, below is the Target file size we want to
    'export. Each time before a new vehicle begins to export the spot file size is checked.
    'If it meets or exceeds the Target size then that group of files is closed and sent out
    'via FTP to the web server and imported.  Then a new group of files is created for the
    'spots, copy rotation and headers.  This was done to take a load off of the machine
    'that's doing the exporting and the web server. Also, we don't have to wait 6 hours to
    'see if Dial's export is working or not.
    'What the desired target file should reach in MB to trigger exporting the file
    If igSmallFiles Then
        ilTargetFileSizeInMB = 1  'Note: needs to even MB or it will be rounded
    Else
        ilTargetFileSizeInMB = 8  'Note: needs to even MB or it will be rounded
    End If
    'Just set ilTargetFileSizeInMB in the line above. llTargetFileValue is the final value
    'that the program uses.
    llTargetFileValue = ilTargetFileSizeInMB * cmOneMegaByte
    'This is the minimum number spots needed to trigger a reindex of the web server.
    'Reindexing the server takes at least 3 minutes and pretty much shuts everyone
    'out while its running so don't do it unless it's a big export
    If smMinSpotsToReIndex = "" Then
        llMinSpotsToReIndex = 500000
    Else
        llMinSpotsToReIndex = CLng(smMinSpotsToReIndex)
    End If
    On Error GoTo ErrHand:
    imTerminate = False
    If Not gPopCopy(smMODate, "Web Export Scheduled Spots") Then
        imExporting = False
        Exit Function
    End If
    'Arrays to make sure that we don't insert a duplicate ast into the web site
    'that's all that they are used for
    llDupBuffSize = 500
    gLogMsg "*** Export Starting. ***", "WebExpSummary.Txt", False
    'Build file names, open the files and write out their first record headers
    If Not mOpenFiles Then
        Exit Function
    End If
    tmEsf(0).sFileName = smToFileDetail
    ilStartNewFiles = False
    imNeedToSend = False
    'Dan 6/26/19 I moved this below web vendors
'    SetResults "Standby, Gathering Export Data.", 0
    SQLQuery = "Select MAX(EsfCode) from ESF_Export_Summary"
    'Set rst = cnn.Execute(SQLQuery)
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If IsNull(rst(0).Value) Then
            lmMaxEsfCode = 1
        Else
            lmMaxEsfCode = rst(0).Value + 1
        End If
    Else
        lmMaxEsfCode = 1
    End If
    tmEsf(0).lCode = lmMaxEsfCode
    On Error GoTo IncCode
    ilRetries = 0
    ilInsertFailed = True
    gEraseEventDate
    Do While ilInsertFailed And ilRetries < 5
        ilInsertFailed = False
        SQLQuery = "Insert Into ESF_Export_Summary ( "
        SQLQuery = SQLQuery & "esfCode, "
        SQLQuery = SQLQuery & "esfStartTime, "
        SQLQuery = SQLQuery & "esfStartDate, "
        SQLQuery = SQLQuery & "esfExpDate, "
        SQLQuery = SQLQuery & "esfNumDays, "
        SQLQuery = SQLQuery & "esfUser, "
        SQLQuery = SQLQuery & "esfMachine, "
        SQLQuery = SQLQuery & "esfFileName "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & tmEsf(0).lCode & ", "
        SQLQuery = SQLQuery & "'" & Format$(tmEsf(0).sStartTime, sgSQLTimeForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$(tmEsf(0).sStartDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$(tmEsf(0).sExpDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & tmEsf(0).iNumDays & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(tmEsf(0).sUser) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(tmEsf(0).sMachine) & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(tmEsf(0).sFileName) & "'"
        SQLQuery = SQLQuery & ") "
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        Else
            ilInsertFailed = False
        End If
        ilRetries = ilRetries + 1
    Loop
    If ilRetries = 5 And ilInsertFailed = False Then
        Exit Function
    End If
    On Error GoTo ErrHand
    imFailedStations = False
    lmFailedStaExpIdx = 0
    ReDim Preserve smStaFailedToExport(0 To 0)
    ilRet = mGetMissedReasons()
    ilRet = mGetReplacementReasons()
    If bgVendorToWebAllowed Then
        '10000
        If gIsTestWebServer() And Not bgTestSystemAllowWebVendors Then
            mSetWebVendorsToTest
        End If
        'Dan 6/26/19
        SetResults "Standby, Sending Web Vendor Data.", 0
        '10467
        gLogMsg "Sending Web Vendor Data", "WebExportLog.Txt", False
        If gUpdateWebVendorsOnWeb() = -1 Then
            SetResults "Web vendor issue. See 'AffErrorLog.txt'", RGB(255, 0, 0)
        End If
        gUpdateWebVendorsHeaderOnWeb Me
    Else
        '10320
        If gWebVendorIsUsed() Then
            SetResults "Web vendor issue. Could not send web vendor information.  See 'AffErrorLog.txt'", RGB(255, 0, 0)
            gLogMsg "Web vendor issue in WebExport-mInitiateExport. Could not send web vendor information because bgVendorToWebAllowed was false", "AffErrorLog.txt", False
        End If
    End If
    'Dan 6/26/19 moved from above
    SetResults "Standby, Gathering Export Data.", 0
    hmJelli = -1
    Set myEnt = New CENThelper
    With myEnt
        .TypeEnt = Sentunpostedtoweb
        .User = igUstCode
        .ThirdParty = Vendors.Web
        .ErrorLog = cmPathForgLogMsg
    End With
    'D.S. 11/16/17
    bgTaskBlocked = False
    For ilLoop = 1 To grdVeh.Rows - 1 Step 1
        If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
            If grdVeh.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                imVefCode = grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)
                smVefName = grdVeh.TextMatrix(ilLoop, VEHINDEX)
                ReDim imExptVefCode(0 To 1) As Integer
                imExptVefCode(0) = imVefCode
                For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
                    If imVefCode = tgVehicleInfo(ilVef).iCode Then
                        smVehicleType = tgVehicleInfo(ilVef).sVehType
                        If smVehicleType = "L" Then
                            For ilIdx = 0 To UBound(tgVehicleInfo) - 1 Step 1
                                If tgVehicleInfo(ilVef).iCode = tgVehicleInfo(ilIdx).iVefCode Then
                                    For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                                        If tgVehicleInfo(ilIdx).iCode = tgVffInfo(ilVff).iVefCode Then
                                             If (tgVffInfo(ilVff).sMergeAffiliate = "S") And (tgVffInfo(ilVff).sMergeWeb <> "S") Then
                                                 imExptVefCode(UBound(imExptVefCode)) = tgVffInfo(ilVff).iVefCode
                                                 ReDim Preserve imExptVefCode(0 To UBound(imExptVefCode) + 1) As Integer
                                             End If
                                             Exit For
                                        End If
                                    Next ilVff
                                End If
                            Next ilIdx
                        End If
                        Exit For
                    End If
                Next ilVef
    
                If ilStartNewFiles Then
                    'Build file names, open the files and write out their first record headers
                    If Not mOpenFiles Then
                        gCloseRegionSQLRst
                        imExporting = False
                        Exit Function
                    End If
                End If
                If sgShowByVehType = "Y" Then
                    smVefName = Mid$(smVefName, 3)
                End If
                gSetMousePointer grdVeh, grdVeh, vbHourglass
                smVehicleExportJelli = "N"
                ilVff = gBinarySearchVff(imVefCode)
                If ilVff = -1 Then
                    ilVff = gPopVff()
                    ilVff = gBinarySearchVff(imVefCode)
                End If
                If ilVff <> -1 Then
                    smVehicleExportJelli = Trim$(tgVffInfo(ilVff).sExportJelli)
                End If
                If (smVehicleExportJelli = "Y") And (hmJelli = -1) Then
                    mOpenJelliFile
                End If
                ReDim lmLastattCode(0 To 0) As Long
                SetResults "Creating " & Trim$(smVefName), 0
                gLogMsg "*** Creating Export for " & Trim$(smVefName) & " - " & "Starting " & smStartDate & " for " & Trim$(txtNumberDays.Text) & " days ***", "WebExportLog.Txt", False
                gLogMsg "    Creating Export for " & Trim$(smVefName) & " - " & "Starting " & smStartDate & " for " & Trim$(txtNumberDays.Text) & " days ***", "WebActivityLog.Txt", False
                lgETime25 = timeGetTime
                lgTtlTime25 = lgTtlTime25 + lgETime25 - lgSTime25
                lgSTime23 = timeGetTime
                ilRet = mExportSpots()
                lgETime23 = timeGetTime
                lgTtlTime23 = lgTtlTime23 + lgETime23 - lgSTime23
                lgSTime25 = timeGetTime
                If (ilRet = False) And (Not igDemoMode) Then
                    gCloseRegionSQLRst
                    gLogMsg "** Terminated - mExportSpots returned False **", "WebExportLog.Txt", False
                    Close #hmToDetail
                    imExporting = False
                    gSetMousePointer grdVeh, grdVeh, vbDefault
                    Exit Function
                End If
                If imTerminate Then
                    gCloseRegionSQLRst
                    gLogMsg "** User Terminated **", "WebExportLog.Txt", False
                    Close #hmToDetail
                    imExporting = False
                    gSetMousePointer grdVeh, grdVeh, vbDefault
                    Exit Function
                End If
                If (igDemoMode) Then
                    imNeedToSend = False
                Else
                    llFileSize = FileLen(smToFileDetail)
                    If lmFileAddSpotCount > 30000 Or ilLoop = (grdVeh.Rows - 1) Then
                        imNeedToSend = False
                        'Write out the WebCopyRot.txt file from the tgCopyRotInfo array
                        mWriteCSFFile
                        'Check bit map to see if using games.  If not no sense in exporting or showing it
                        If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                            mWriteEventInfoFile
                        End If
                        ilFileSent = True
                        mCloseFiles
                        mInsertIntoWQF
                        'FTP all the files we just created
                        While imFtpInProgress
                            Sleep (250)
                            ilRet = mCheckFTPStatus()
                        Wend
                        If Not imFtpInProgress Then
                            gLogMsg "Calling to FTP Files", "WebExportLog.Txt", False
                            lgSTime3 = timeGetTime
                            If Not mFTPFiles Then
                                lgETime3 = timeGetTime
                                lgTtlTime3 = lgTtlTime3 + (lgETime3 - lgSTime3)
                            
                                gCloseRegionSQLRst
                                imExporting = False
                                Exit Function
                            End If
                        End If
                        ilStartNewFiles = True
                        If imFTPIsOn Then
                            If Not myEnt.UpdateIncompleteByFilename(Successful) Then
                                 gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
                            End If
                        Else
                            If Not myEnt.UpdateIncompleteByFilename(NotSent) Then
                                 gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
                            End If
                        End If
                    Else
                        ilStartNewFiles = False
                        If imFtpInProgress Then
                            ilRet = mCheckFTPStatus()
                        End If
                        If imWaiting = True Then
                            ilRet = mCheckStatus()
                        End If
                        If Not imWaiting And ilFileSent = True Then
                            'Start Web Process Time
                            lgSTime6 = timeGetTime
                            mProcessWebQueue
                            'End Web Process Time
                            lgETime6 = timeGetTime
                            lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                        End If
                    End If
                End If
                If igExportSource = 2 Then DoEvents
            End If
        End If
    Next ilLoop
    If hmJelli > 0 Then
        Close #hmJelli
        hmJelli = -1
    End If
    gCloseRegionSQLRst
    txtCallLetters.Visible = False
    lblCallLetters.Visible = False
    'This covers the case where the file wasn't big enough to meet target size,
    'but stills needs to be sent before we exit
    If (igDemoMode) Then
        mCloseFiles
    Else
        If imNeedToSend Then
            mWriteCSFFile
            'D.S. Check bit map to see if using games.  If not, no sense in exporting of showing it
            If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
            mWriteEventInfoFile
            End If
            mCloseFiles
                imExporting = False
            
            'Most likely we insert one more time into the WQF - the file that was smaller than llTargetFileValue
            mInsertIntoWQF
            mFTPFiles
            '7458
            If imFTPIsOn Then
                If Not myEnt.UpdateIncompleteByFilename(Successful) Then
                     gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
                End If
            Else
                If Not myEnt.UpdateIncompleteByFilename(NotSent) Then
                     gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
                End If
            End If
        Else
            mCloseFiles
        End If
        'Wait on any unfinished FTP jobs
        lgSTime3 = timeGetTime
        While imFtpInProgress
            ilRet = mCheckFTPStatus()
            Sleep (100)
            DoEvents
        Wend
        If ilRet Then    ' JD 01-05-22 TTP: 10372
            lgETime3 = timeGetTime
            lgTtlTime3 = lgTtlTime3 + (lgETime3 - lgSTime3)
            'Check the WQF table and see if there any files that were not FTPed.  If so, send them
            SQLQuery = "SELECT * FROM WQF_Web_Queue WHERE wqfFTPStatus = 0"
            Set rst = gSQLSelectCall(SQLQuery)
            Do While Not rst.EOF    ' JD 01-05-22 TTP: 10372
                ReDim mFtpArray(0 To 0)
                gLogMsg "Sending " & Trim$(mFtpArray(0)), "WebActivityLog.Txt", False
                mFtpArray(0) = Trim$(rst!wqffilename)
                imFtpInProgress = True
                If imFTPIsOn Then
                ilRet = csiFTPFileToServer(Trim$(mFtpArray(0)))
                End If
                ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
                While imFtpInProgress
                    ilRet = mCheckFTPStatus()
                    Sleep (1000)
                    DoEvents
                Wend
                rst.MoveNext
                If ilRet = False Then
                    ' JD 01-05-22 TTP: 10372
                    Exit Do
                End If
            Loop
        imSomeThingToDo = True
        End If
        'Start Web Process Time
        lgSTime6 = timeGetTime
        'This looks like a possible Endless loop
        ilRetries = 0
        While (imSomeThingToDo = True Or imWaiting = True) 'And ilRetries < 5
            If imWaiting Then
                ilRet = mCheckStatus()
                Sleep (1000)
            Else
                mProcessWebQueue
            End If
        Wend
        If ilRetries = 5 Then
            gLogMsg "Error: Retries were exceeded in frmWebExportSchdSpot - mInitiateExport", "WebExportLog.Txt", False
        End If
        'End Web Process Time
        lgETime6 = timeGetTime
        lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
        If igExportSource = 2 Then DoEvents
        'We have exported all of the files so reindex the web server
        If lmTotalAddSpotCount >= llMinSpotsToReIndex Then
            ilRet = mReIndexServer()
        End If
        If lmTotalAddSpotCount > 0 Then
            'Show the final message with the totals of spots imported an emails sent
            Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports", "WebEmails")
            ilRet = mSendEmails()
        End If
    End If
    'D.S. 11/16/17 - TTP #8684
    'D.S. 12/27/17 - moved to cmdExport_Click
    'If bgTaskBlocked And igExportSource <> 2 Then
    '    gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_mmddyyyy.txt for further information.", vbCritical
    'End If
    'bgTaskBlocked = False
    Erase lmLastattCode
    mInitiateExport = True
    Set myEnt = Nothing
    Exit Function
IncCode:
    lmMaxEsfCode = lmMaxEsfCode + 1
    tmEsf(0).lCode = lmMaxEsfCode
    ilInsertFailed = True
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mInitiateExport"
    'Debug
    'Resume Next
    Exit Function
End Function
Private Sub mClose()
    
    Dim llTtlTime As Long
    Dim slTtlTime As String
    Dim llHour As Long
    Dim llMin As Long
    Dim llSec As Long
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim ilSuccess As Integer
    Dim slErrorMsg As String
    Dim ilRet As Integer
    
    lmTotalRecordsProcessed = 0
    imExporting = False
    lmTotalRecordsProcessed = lmTotalAddSpotCount + lmTotalDeleteSpotCount
    ilSuccess = False
    'Check to see if all the counts between the web and the affiliate match
    If (lmTotalRecordsProcessed = lmWebTtlSpots) Or igDemoMode Then
        If (lmTtlHeaders = lmWebTtlHeaders) Or igDemoMode Then
            If (lmTtlComments = lmWebTtlComments) Or igDemoMode Then
                ilSuccess = True
                SetResults "All Export Counts were Reconciled", RGB(0, 155, 0)
                gLogMsg "All Export Counts were Reconciled", "WebActivityLog.Txt", False
            Else
                gLogMsg "ERROR: " & "Comments Exported: " & CLng(lmTtlComments) & " Web Comments Imported " & lmWebTtlComments & ".", "WebExportLog.Txt", False
                gLogMsg "ERROR: " & "Comments Exported: " & CLng(lmTtlComments) & " Web Comments Imported " & lmWebTtlComments & ".", "WebActivityLog.Txt", False
                slErrorMsg = "ERROR: " & "Comments Exported: " & CLng(lmTtlComments) & " Web Comments Imported " & lmWebTtlComments & "."
                SetResults "Comment Counts did not reconcile.", RGB(255, 0, 0)
            End If
        Else
            gLogMsg "ERROR: " & "Headers Exported: " & CLng(lmTtlHeaders) & " Web Comments Imported " & lmWebTtlHeaders & ".", "WebExportLog.Txt", False
            gLogMsg "ERROR: " & "Headers Exported: " & CLng(lmTtlHeaders) & " Web Comments Imported " & lmWebTtlHeaders & ".", "WebActivityLog.Txt", False
            slErrorMsg = "ERROR: " & "Headers Exported: " & CLng(lmTtlHeaders) & " Web Comments Imported " & lmWebTtlHeaders & "."
            SetResults "Header Counts did not reconcile.", RGB(255, 0, 0)
        End If
    Else
        If lmTotalRecordsProcessed > 0 Then
            gLogMsg "ERROR: " & "Spots Exported: " & CLng(lmTotalRecordsProcessed) & " Web Spots Imported " & lmWebTtlSpots & ".", "WebExportLog.Txt", False
            gLogMsg "ERROR: " & "Spots Exported: " & CLng(lmTotalRecordsProcessed) & " Web Spots Imported " & lmWebTtlSpots & ".", "WebActivityLog.Txt", False
            slErrorMsg = " ERROR: " & "Spots Exported: " & CLng(lmTotalRecordsProcessed) & " Web Spots Imported " & lmWebTtlSpots & "."
            SetResults "Spots Counts did not reconcile.", RGB(255, 0, 0)
        End If
    End If
    If Not imTerminate Then
        SetResults "Total Emails Sent: " & CLng(lmWebTtlEmail), RGB(0, 155, 0)
        SetResults "Total Add Spots Exported: " & CLng(lmTotalAddSpotCount), RGB(0, 155, 0)
        SetResults "Total Spots Processed: " & CLng(lmTotalRecordsProcessed), RGB(0, 155, 0)
    End If
    If bgIllegalCharsFound Then
        SetResults "*** Illegal Characters were found. Please see: AffBadCharLog.Txt", RGB(0, 155, 0)
        SetResults "in the messages folder.  Call Counterpoint. ***", RGB(0, 155, 0)
        gLogMsg "", "AffBadCharLog.Txt", False
    End If
    If ilSuccess And Not imTerminate Then
        If Not imFailures Then
        SetResults "*** Export Completed Successfully. ***", RGB(0, 155, 0)
    Else
            SetResults "*** Export Completed, but had at least one Failure. ***", RGB(255, 0, 0)
            SetResults "*** See WebExportLog.txt. ***", RGB(255, 0, 0)
        End If
    Else
        SetResults "--- Export Failed. ---", RGB(255, 0, 0)
    End If
    'Web Summary Log
    gLogMsg "Total Add Spots Exported: " & CLng(lmTotalAddSpotCount), "WebExpSummary.Txt", False
    gLogMsg "Total Spots Processed: " & CLng(lmTotalRecordsProcessed), "WebExpSummary.Txt", False
    If ilSuccess And Not imTerminate Then
        gLogMsg "*** Export Completed Successfully. ***", "WebExpSummary.Txt", False
    Else
        If imTerminate Then
            gLogMsg "** User Terminated **", "WebExpSummary.Txt", False
            slErrorMsg = "** User Terminated **"
        Else
            gLogMsg slErrorMsg, "WebExpSummary.Txt", False
        End If
    End If
    gLogMsg "", "WebExpSummary.Txt", False
    'Web Export Log
    gLogMsg "Total Add Spots Exported: " & CLng(lmTotalAddSpotCount), "WebExportLog.Txt", False
    gLogMsg "Total Spots Processed: " & CLng(lmTotalRecordsProcessed), "WebExportLog.Txt", False
    If ilSuccess And Not imTerminate Then
        gLogMsg "*** Export Completed Successfully. ***", "WebExportLog.Txt", False
    Else
        If imTerminate Then
            gLogMsg "** User Terminated **", "WebExportLog.Txt", False
            slErrorMsg = "** User Terminated **"
        Else
            gLogMsg slErrorMsg, "WebExportLog.Txt", False
        End If
    End If
    gLogMsg "", "WebExportLog.Txt", False
    SQLQuery = "Update ESF_Export_Summary Set "
    SQLQuery = SQLQuery & "esfTtlAdd = " & CLng(lmTotalAddSpotCount) & ", "
    SQLQuery = SQLQuery & "esfTtlDel = " & CLng(lmTotalDeleteSpotCount) & ", "
    SQLQuery = SQLQuery & "esfTtlAddDel = " & CLng(lmTotalRecordsProcessed) & ", "
    SQLQuery = SQLQuery & "esfTtlEmails = " & lmWebTtlEmail & ", "
    SQLQuery = SQLQuery & "esfTtlHdrs = " & lmTtlHeaders & ", "
    SQLQuery = SQLQuery & "esfTtlComments = " & lmTtlComments & ", "
    SQLQuery = SQLQuery & "esfEndTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "esfEndDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "esfElspdTime = '" & slTtlTime & "', "
    If ilSuccess Then
        SQLQuery = SQLQuery & "esfExpSuccess = '" & "Y" & "', "
    Else
        SQLQuery = SQLQuery & "esfExpSuccess = '" & "N" & "', "
    End If
    SQLQuery = SQLQuery & "esfErrors = '" & slErrorMsg & "' "
    SQLQuery = SQLQuery & " Where esfCode = " & lmMaxEsfCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        GoSub ErrHand:
    End If
    lbcMsg.ListIndex = -1   ' Finish with nothing selected
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    gSetMousePointer grdVeh, grdVeh, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mClose"
    Exit Sub
End Sub

Private Sub mWriteCSFFile()

    Dim ilLoop As Integer
    Dim slStr As String
    
    On Error GoTo ErrHand
    For ilLoop = 0 To UBound(tgCopyRotInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        slStr = tgCopyRotInfo(ilLoop).lCode & ","
        slStr = slStr & """" & gFixDoubleQuoteWithSingle(tgCopyRotInfo(ilLoop).sComment) & """"
        Print #hmToCpyRot, gRemoveIllegalCharsAndLog(slStr, smToCpyRot, ilLoop + 1, True)
        lmFileCommentCount = lmFileCommentCount + 1
    Next ilLoop
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mWriteCSFFile"
    Exit Sub
End Sub

Private Sub SetResults(Msg As String, FGC As Long)

    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String
    Dim ilLoop As Integer
    
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    If igExportSource = 2 Then DoEvents
    
    If lbcMsg.ListCount > 0 Then
        For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
            If igExportSource = 2 Then DoEvents
            slTemp = lbcMsg.List(ilLoop)
            'create horz. scrool bar if the text is wider than the list box
            ilLen = Me.TextWidth(slTemp)
            If Me.ScaleMode = vbTwips Then
                ilLen = ilLen / Screen.TwipsPerPixelX  ' if twips change to pixels
            End If
            If ilLen > ilMaxLen Then
                ilMaxLen = ilLen
            End If
        Next ilLoop
        SendMessageByNum lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, ilMaxLen + 250, 0
        If igExportSource = 2 Then DoEvents
    End If
End Sub

Private Sub mClearAlerts()
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim ilLoop As Integer
    Dim iRet As Integer
    
    On Error GoTo ErrHand
    'Alerts are only defined with a Monday date
    llSDate = DateValue(gObtainPrevMonday(gAdjYear(smStartDate)))
    llEDate = DateValue(gAdjYear(Format$(DateAdd("d", imNumberDays - 1, smStartDate), "mm/dd/yy")))
    For ilLoop = 0 To UBound(imExportedVefArray) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        '11/25/12
        imVefCode = imExportedVefArray(ilLoop)
        For llDate = llSDate To llEDate Step 7
            slDate = Format$(llDate, "ddddd")
            iRet = gAlertClearFinalAndReprint("A", "F", "S", imExportedVefArray(ilLoop), slDate)
            If igExportSource = 2 Then DoEvents
        Next llDate
    Next ilLoop
    Exit Sub
ErrHand:
    'debug
    'Resume Next
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mClearAlerts"
    Exit Sub
End Sub

Private Function mBuildDetailRecs(cprst As ADODB.Recordset, sArray() As ASTINFO, sTransType As String, sEndDate As String) As Integer

    Dim llLoop As Long
    Dim llBinRet As Long
    Dim ilRet As Integer
    Dim iAddDelete As Integer
    Dim sAdvt As String
    Dim sCart As String
    Dim sISCI As String
    Dim sCreative As String
    Dim sProd As String
    Dim lCrfCsfCode As Long
    Dim sPledgeStartDate As String
    Dim sPledgeEndDate As String
    Dim sPledgeStartTime As String
    Dim sPledgeEndTime As String
    ReDim iDays(0 To 6) As Integer
    Dim iDay As Integer
    Dim iIndex As Integer
    Dim sLen As String
    Dim iFound As Integer
    Dim iExport As Integer
    Dim slStr As String
    Dim rst_Temp As ADODB.Recordset
    'Dim rst_Lst As ADODB.Recordset
    Dim llMax As Long
    Dim llAstTest As Long
    Dim slTemp As String
    Dim slTemp1 As String
    Dim llIdx As Long
    Dim sRCart As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim sRProd As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim ilOrgStatusCode As Integer
    Dim ilWebStatusCode As Integer
    Dim llRet As Long
    Dim sFeedDate As String
    Dim sFeedTime As String
    Dim ilLstAdfCode As Integer
    Dim ilLstAnfCode As Integer
    Dim slLstProd As String
    Dim slLstCart As String
    Dim slLstISCI As String
    Dim slLiveCopy As String
    Dim llLstCrfCsfCode As Long
    Dim llLstCpfCode As Long
    Dim llLstGsfCode As Long
    Dim slRotCpyEndDate As String
    Dim llCifCode As Long
    Dim slIsDaypart As Boolean
    Dim slEstDayAndTime As String
    Dim slBeforeOrAfter As String * 1
    Dim llVpf As Long
    Dim slActualAirDateTime As String
    Dim slPrevPldgSTime As String
    Dim slPrevPldgETime As String
    Dim slBaseStartTime As String
    Dim slBaseStartDate As String
    Dim ilAirPlayNo As Integer
    Dim slAirDateAndTime As String
    Dim llEstimatesExist As Boolean
    Dim slEstimateResult As String
    Dim ilStart As Integer
    Dim ilIdx As Integer
    Dim ilPos As Integer
    Dim ilLenOption As Integer
    
    On Error GoTo ErrHand
    
    mBuildDetailRecs = False
    imSpotCount = 0
    slBaseStartTime = ""
    slBaseStartDate = ""
    For llLoop = LBound(sArray) To UBound(sArray) - 1 Step 1
        DoEvents
        ilAirPlayNo = sArray(llLoop).iAirPlay
        If ilAirPlayNo > UBound(tmAirTimeInfo) Then
            ilStart = UBound(tmAirTimeInfo)
            ReDim Preserve tmAirTimeInfo(0 To ilAirPlayNo + 1) As AIRTIMEINFO
            For ilIdx = ilStart + 1 To ilAirPlayNo Step 1
                tmAirTimeInfo(ilIdx).sBaseStartDate = ""
                tmAirTimeInfo(ilIdx).sBaseStartTime = ""
                tmAirTimeInfo(ilIdx).sEstimatedEndTime = ""
                tmAirTimeInfo(ilIdx).sEstimatedStartTime = ""
                tmAirTimeInfo(ilIdx).sPrevPldgETime = ""
                tmAirTimeInfo(ilIdx).sPrevPldgSTime = ""
            Next ilIdx
        End If
        slBaseStartDate = Trim(tmAirTimeInfo(ilAirPlayNo).sBaseStartDate)
        slBaseStartTime = Trim(tmAirTimeInfo(ilAirPlayNo).sBaseStartTime)
        smEstimatedEndTime = Trim(tmAirTimeInfo(ilAirPlayNo).sEstimatedEndTime)
        smEstimatedStartTime = Trim(tmAirTimeInfo(ilAirPlayNo).sEstimatedStartTime)
        slPrevPldgETime = Trim(tmAirTimeInfo(ilAirPlayNo).sPrevPldgETime)
        slPrevPldgSTime = Trim(tmAirTimeInfo(ilAirPlayNo).sPrevPldgSTime)
        If igExportSource = 2 Then DoEvents
        slRotCpyEndDate = ""
        'No Regional Copy Exists
        If sArray(llLoop).iRegionType = 0 Then
            llCifCode = sArray(llLoop).lCifCode
        End If
        'Regional Copy Exists
        If sArray(llLoop).iRegionType > 0 Then
            llCifCode = sArray(llLoop).lRCifCode
        End If
        'I don't know how this could happen but....
        If sArray(llLoop).iRegionType < 0 Then
            llCifCode = -1
        End If
        lgSTime15 = timeGetTime
        slRotCpyEndDate = ""
        If llCifCode <> -1 Then
            'slRotCpyEndDate = mObtainRotEndDate(llCifCode)
            llRet = gBinarySearchCifCpf(llCifCode)
            'TTP 9923 changed below llRet from ilRet
            If llRet = -1 Then
                ilRet = gPopCifCpfInfo(slBaseStartDate)
                llRet = gBinarySearchCifCpf(llCifCode)
            End If
            If llRet <> -1 Then
                slRotCpyEndDate = Format$(tgCifCpfInfo1(llRet).cifRotEndDate, sgSQLDateForm)
            Else
                slRotCpyEndDate = Format$(mObtainRotEndDate(llCifCode), sgSQLDateForm)
            End If
            lgCount8 = lgCount8 + 1
        End If
        lgETime15 = timeGetTime
        lgTtlTime15 = lgTtlTime15 + (lgETime15 - lgSTime15)
        lgSTime19 = timeGetTime
        If Trim$(slRotCpyEndDate) = "" Then
            'Set it to the export periods end date
            slRotCpyEndDate = Format$(smEndDate, sgSQLDateForm)
        End If
        lgETime19 = timeGetTime
        lgTtlTime19 = lgTtlTime19 + (lgETime19 - lgSTime19)
        If imTerminate Then
            gSetMousePointer grdVeh, grdVeh, vbDefault
            cmdCancel.Enabled = True
            imExporting = False
            Exit Function
        End If
        '3/9/16: Bypass MG that was posted in a future week not yet exported to the Web
        'and the MG week imported back from the Web
        'If (sArray(llLoop).iStatus Mod 100 <= 10) And (DateValue(gAdjYear(sArray(llLoop).sFeedDate)) >= DateValue(gAdjYear(smStartDate))) And (DateValue(gAdjYear(sArray(llLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) Then
        If ((sArray(llLoop).iStatus Mod 100 <= 10) Or (sArray(llLoop).iStatus Mod 100 = ASTAIR_MISSED_MG_BYPASS)) And (DateValue(gAdjYear(sArray(llLoop).sFeedDate)) >= DateValue(gAdjYear(smStartDate))) And (DateValue(gAdjYear(sArray(llLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) Then
            If (tgStatusTypes(sArray(llLoop).iStatus).iPledged <> 2) Then
                iAddDelete = 0
                sAdvt = "Missing"
                sCart = ""
                sISCI = ""
                sCreative = ""
                lCrfCsfCode = 0
                ilLstAdfCode = sArray(llLoop).iAdfCode
                ilLstAnfCode = sArray(llLoop).iAnfCode
                slLstProd = sArray(llLoop).sProd
                slLstCart = sArray(llLoop).sCart
                slLstISCI = sArray(llLoop).sISCI
                ' slLiveCopy= sArray(llLoop).
                llLstCrfCsfCode = sArray(llLoop).lCrfCsfCode
                llLstCpfCode = sArray(llLoop).lCpfCode
                llLstGsfCode = sArray(llLoop).lgsfCode
                'If Not rst_Lst.EOF Then
                If sTransType <> "D" Then
                    If igExportSource = 2 Then DoEvents
                    ilRet = gBinarySearchAdf(CLng(ilLstAdfCode))  '(rst_Lst!lstAdfCode)
                    If ilRet = -1 Then
                        ilRet = gPopAdvertisers()
                        ilRet = gBinarySearchAdf(CLng(ilLstAdfCode))
                    End If
                    If ilRet <> -1 Then
                        sAdvt = Trim$(tgAdvtInfo(ilRet).sAdvtName)
                    Else
                        sAdvt = ""
                    End If
                    If igExportSource = 2 Then DoEvents
                    sProd = Trim$(slLstProd)
                    If (Trim$(slLstCart) = "") Or (Left$(slLstCart, 1) = Chr$(0)) Then
                        sCart = ""
                    Else
                        sCart = Trim$(slLstCart) 'Trim$(rst_Lst!lstCart)
                    End If
                    sISCI = Trim$(slLstISCI)
                    llRet = gBinarySearchCpf(llLstCpfCode)  '(rst_Lst!lstCpfCode)
                    If llRet = -1 Then
                        ilRet = gPopCpf()
                        llRet = gBinarySearchCpf(llLstCpfCode)  '(rst_Lst!lstCpfCode)
                    End If
                    If llRet <> -1 Then
                        sCreative = Trim$(tgCpfInfo(llRet).sCreative)
                    Else
                        sCreative = ""
                    End If
                    If igExportSource = 2 Then DoEvents
                    lCrfCsfCode = llLstCrfCsfCode   'rst_Lst!lstCrfCsfCode
                    lmGsfCode = llLstGsfCode  'CLng(rst_Lst!lstGsfCode)
                Else
                    If sTransType = "D" Then
                        If igExportSource = 2 Then DoEvents
                        ilRet = gBinarySearchAdf(CLng(ilLstAdfCode))  '(rst_Lst!lstAdfCode)
                        If ilRet = -1 Then
                            ilRet = gPopAdvertisers()
                            ilRet = gBinarySearchAdf(CLng(ilLstAdfCode))
                        End If
                        If ilRet <> -1 Then
                            sAdvt = Trim$(tgAdvtInfo(ilRet).sAdvtName)
                        Else
                            sAdvt = ""
                        End If
                        If igExportSource = 2 Then DoEvents
                        sProd = Trim$(slLstProd)
                        sCart = ""
                        sISCI = ""
                        sCreative = ""
                    End If
                    lmGsfCode = 0
                End If
                If sArray(llLoop).iRegionType > 0 Then
                    '9/10/14: If region assigned the same copy as the generic, retain the generic product
                    '11/21/14: Replace cart with ISCI to handle clients not using cart numbers
                    'If sCart <> Trim$(sArray(llLoop).sRCart) Then
                    If sISCI <> Trim$(sArray(llLoop).sRISCI) Then
                        sProd = Trim$(sArray(llLoop).sRProduct)  'sRProd
                        sCart = Trim$(sArray(llLoop).sRCart)  'sRCart
                        sISCI = Trim$(sArray(llLoop).sRISCI)  'sRISCI
                        sCreative = Trim$(sArray(llLoop).sRCreativeTitle)  'sRCreative
                    End If
                    lCrfCsfCode = sArray(llLoop).lRCrfCsfCode  'lRCrfCsfCode
                End If
                If sCreative <> "" Then
                    If Asc(sCreative) = 0 Then
                        sCreative = ""
                    End If
                End If
                sPledgeStartDate = Format$(sArray(llLoop).sPledgeDate, "m/d/yyyy")
                'D.S. Please let me know if this mapping can made more confusing. I don't think so!
                'Also if they pledge a 2 Delay B'cast or a 10 'Delay Comm/Prgm  The it maps to a
                '2 Delay B'cast per Jim F. 7/28/06
                'the original status ast status code as pledged for this avail
                ilOrgStatusCode = sArray(llLoop).iPledgeStatus
                'D.S. Check bit map to see if using games.  If not no sense in exporting of showing it
                If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
                    If lmGsfCode <> 0 Then
                        Call mBuildEventInfo(lmGsfCode)
                    End If
                End If
                Select Case ilOrgStatusCode
                    Case 0
                        'C - Program and Commercial Aired Live.
                        '   Send Web = 1  Status = 0  Screen = 1-Aired Live
                        ilWebStatusCode = 1
                    Case 1
                        'K - Delay B'cast
                        '   Send Web = 5  Status = 1   Screen = 2-Delay B'cast
                        ilWebStatusCode = 2
                    
                    Case 9
                        'D - Program and Commercial were both delayed.
                        '   Send Web = 2  Status = 9   Screen = 10-Delay Cmml/Prg
                        ilWebStatusCode = 2
                    Case 10
                        'D - Air Comm Only
                        '   Web returns = 3  Status = 10   Screen = 11-Air Comm Only
                        ilWebStatusCode = 3
                End Select
                If tgStatusTypes(sArray(llLoop).iPledgeStatus).iPledged = 0 Then
                    sPledgeEndDate = sPledgeStartDate
                Else
                    gUnMapDays sArray(llLoop).sPdDays, iDays()
                    iDay = Weekday(sPledgeStartDate, vbMonday) - 1
                    sPledgeEndDate = sPledgeStartDate
                    For iIndex = iDay + 1 To 6 Step 1
                        If iDays(iIndex) Then
                            sPledgeEndDate = DateAdd("d", 1, sPledgeEndDate)
                        Else
                            Exit For
                        End If
                    Next iIndex
                End If
                If Second(sArray(llLoop).sPledgeStartTime) <> 0 Then
                    sPledgeStartTime = Format$(sArray(llLoop).sPledgeStartTime, "h:mm:ssa/p")
                Else
                    sPledgeStartTime = Format$(sArray(llLoop).sPledgeStartTime, "h:mma/p")
                End If
                If Len(Trim$(sArray(llLoop).sPledgeEndTime)) <= 0 Then
                    sPledgeEndTime = sPledgeStartTime
                Else
                    If Second(sArray(llLoop).sPledgeEndTime) <> 0 Then
                        sPledgeEndTime = Format$(sArray(llLoop).sPledgeEndTime, "h:mm:ssa/p")
                    Else
                        sPledgeEndTime = Format$(sArray(llLoop).sPledgeEndTime, "h:mma/p")
                    End If
                End If
                ' Gather up the Feed date and time
                sFeedDate = Format$(sArray(llLoop).sFeedDate, "m/d/yyyy")
                If Second(sArray(llLoop).sFeedTime) <> 0 Then
                    sFeedTime = Format$(sArray(llLoop).sFeedTime, "h:mm:ssa/p")
                Else
                    sFeedTime = Format$(sArray(llLoop).sFeedTime, "h:mma/p")
                End If
                sLen = Trim$(Str$(sArray(llLoop).iLen))
                iFound = False
                iExport = 1
                If igExportSource = 2 Then DoEvents
                If InStr(1, sAdvt, "Missing", vbTextCompare) = 1 And sTransType <> "D" Then
                    SetResults Trim$(smVefName) & ": Advertiser Missing on " & Format$(sArray(llLoop).sAirDate, "ddddd") & " at " & Format$(sArray(llLoop).sAirTime, "ttttt"), 0
                    gLogMsg "Error: " & Trim$(smVefName) & ": Advertiser Missing on " & Format$(sArray(llLoop).sAirDate, "ddddd") & " at " & Format$(sArray(llLoop).sAirTime, "ttttt"), "WebExportLog.Txt", False
                    '7458
                    If Not myEnt.Add(sArray(llLoop).sFeedDate, sArray(llLoop).lgsfCode, Asts) Then
                        gLogMsg "ERROR: Failed to Create ENT Add record.", "WebExportLog.Txt", False
                    End If
                Else
                    If iExport <> 0 Then
                        If sTransType = "D" Then
                            slStr = Trim$(Str$(gGetLogAttID(sArray(llLoop).iVefCode, imShttCode, sArray(llLoop).lAttCode))) & ","
                        Else
                            slStr = Trim$(Str$(lmWebAttID)) & ","
                        End If
                        'Increment start and end times to true their times
                        If sArray(llLoop).sPdTimeExceedsFdTime = "Y" Then
                            slIsDaypart = True
                        Else
                            slIsDaypart = False
                        End If
                        slEstDayAndTime = ""
                        slAirDateAndTime = ""
                        llEstimatesExist = False
                        If slIsDaypart And lmEstimatesExist Then
                            slEstimateResult = mEstimatedDayAndTime(sArray(), llLoop)
                            If slEstimateResult <> "" Then
                                llEstimatesExist = True
                            End If
                        End If
                        If slIsDaypart And llEstimatesExist Then
                            lgCount11 = lgCount11 + 1
                            slEstDayAndTime = slEstimateResult  'mEstimatedDayAndTime(sArray(), llLoop)
                            sArray(llLoop).sFeedTime = sArray(llLoop).sFeedTime
                            sArray(llLoop).iLen = sArray(llLoop).iLen
                            'debug
                            'If llLoop = 3 Then
                            '    llLoop = llLoop
                            'End If
                            If slPrevPldgSTime = slPrevPldgETime Or slBaseStartDate <> sPledgeStartDate Then
                                slBaseStartTime = smEstimatedStartTime
                                slBaseStartDate = sPledgeStartDate
                                smEstimatedEndTime = DateAdd("s", sLen, smEstimatedStartTime)
                                slPrevPldgETime = smEstimatedEndTime
                                slPrevPldgSTime = smEstimatedStartTime
                            Else
                                If slBaseStartTime = smEstimatedStartTime Then
                                    smEstimatedStartTime = slPrevPldgETime
                                    slPrevPldgSTime = slPrevPldgETime
                                    smEstimatedEndTime = DateAdd("s", sLen, smEstimatedEndTime)
                                    slPrevPldgETime = smEstimatedEndTime
                                    smEstimatedStartTime = slPrevPldgSTime
                                Else
                                    slBaseStartTime = smEstimatedStartTime
                                    slBaseStartDate = sPledgeStartDate
                                    smEstimatedEndTime = DateAdd("s", sLen, smEstimatedStartTime)
                                    slPrevPldgETime = smEstimatedEndTime
                                    slPrevPldgSTime = smEstimatedStartTime
                                End If
                            End If
                            slEstDayAndTime = Format$(smEstimatedDate, "yyyy-mm-dd") & "," & Format$(smEstimatedStartTime, "h:mm:ssa/p") & ","
                            slAirDateAndTime = Format$(smEstimatedDate, "yyyy-mm-dd") & " " & Format$(smEstimatedStartTime, "h:mm:ssam/pm")
                        Else
                            '9968 yes, date is funky, but it's what the Javascript code expects
'                            slEstDayAndTime = """" & "2000-01-01" & """" & "," & "00:00:00a" & ","
                            slEstDayAndTime = """" & "1/1/2000" & """" & "," & "00:00:00a" & ","
                            slAirDateAndTime = "NULL"
                        End If
                        If Not slIsDaypart Then
                            If slPrevPldgSTime = slPrevPldgETime Or slBaseStartDate <> sPledgeStartDate Then
                                slBaseStartTime = sPledgeStartTime
                                slBaseStartDate = sPledgeStartDate
                                sPledgeEndTime = DateAdd("s", sLen, sPledgeStartTime)
                                slPrevPldgETime = sPledgeEndTime
                                slPrevPldgSTime = sPledgeStartTime
                            Else
                                If slBaseStartTime = sPledgeStartTime Then
                                    sPledgeStartTime = slPrevPldgETime
                                    slPrevPldgSTime = slPrevPldgETime
                                    sPledgeEndTime = DateAdd("s", sLen, sPledgeStartTime)
                                    slPrevPldgETime = sPledgeEndTime
                                Else
                                    slBaseStartTime = sPledgeStartTime
                                    sPledgeEndTime = DateAdd("s", sLen, sPledgeStartTime)
                                    slPrevPldgETime = sPledgeEndTime
                                    slPrevPldgSTime = sPledgeStartTime
                                End If
                            End If
                            sPledgeStartTime = Format$(sPledgeStartTime, "h:mm:ssa/p")
                            sPledgeEndTime = Format$(sPledgeEndTime, "h:mm:ssa/p")
                            slEstDayAndTime = slBaseStartDate & "," & sPledgeStartTime & ","
                            slAirDateAndTime = Format$(slBaseStartDate, "yyyy-mm-dd") & " " & Format$(sPledgeStartTime, "hh:mm:ssam/pm")
                        End If
                        'End New
                        slStr = slStr & """" & sAdvt & """" & ","
                        If sProd <> "" Then
                            slStr = slStr & """" & sProd & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        slStr = slStr & "~" & "," & sPledgeStartDate & "," & sPledgeEndDate & "," & _
                                sPledgeStartTime & "," & sPledgeEndTime & "," & sFeedDate & "," & sFeedTime & _
                                "," & sLen & ","
                        
                        If sCart <> "" Then
                            slStr = slStr & """" & sCart & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        If sISCI <> "" Then
                            If Not mIsRemoveForIDC() Then
                                slStr = slStr & """" & sISCI & """" & ","
                            Else
                                slStr = slStr & ","
                            End If
                        Else
                            slStr = slStr & ","
                        End If
                        If sCreative <> "" Then
                            slStr = slStr & """" & sCreative & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        'Get the Ast Code
                        slStr = slStr & Trim$(Str$(sArray(llLoop).lCode))
                        llAstTest = sArray(llLoop).lCode
                        lgSTime16 = timeGetTime
                        ilLenOption = 0
                        If lCrfCsfCode <> 0 Then
                            sgCopyComment = ""
                            If mGetCSFComment(lCrfCsfCode, True) Then
                                 slStr = slStr & "," & CStr(lCrfCsfCode) & ","
                                 'D.S. 10/26/19
                                 ilPos = InStr(1, sgCopyComment, ":10", vbTextCompare)
                                 If ilPos > 0 Then
                                    ilPos = InStr(1, sgCopyComment, ":15", vbTextCompare)
                                    If ilPos > 0 Then
                                        ilLenOption = 1
                                    End If
                                End If
                            Else
                                slStr = slStr & ",0,"
                                ' Log this as an error in the message file.
                                gLogMsg "** Spot Comment is missing for lstCrfCsfCode = " & Str(lCrfCsfCode) & " **", "WebExportLog.Txt", False
                            End If
                        Else
                            slStr = slStr & ",0,"
                        End If
                        lgETime16 = timeGetTime
                        lgTtlTime16 = lgTtlTime16 + (lgETime16 - lgSTime16)
                        'Show the avails name
                        If ilLstAnfCode > 0 Then
                            If igSendAvails And sTransType <> "D" Then
                                ilRet = gBinarySearchAnf(ilLstAnfCode)
                                If ilRet = -1 Then
                                    gPopAvailNames
                                    ilRet = gBinarySearchAnf(ilLstAnfCode)
                                End If
                                If ilRet <> -1 Then
                                    slStr = slStr & """" & Trim$(tgAvailNamesInfo(ilRet).sName) & """" & ","
                                Else
                                    slStr = slStr & """" & """" & ","
                                End If
                                If igExportSource = 2 Then DoEvents
                            Else
                                slStr = slStr & """" & """" & ","
                            End If
                        Else
                                slStr = slStr & """" & """" & ","
                        End If
                        slStr = slStr & """" & ilWebStatusCode & """" & ","
                        slStr = slStr & CStr(lmGsfCode) & ","
                        slStr = slStr & slRotCpyEndDate & ","
                        If sArray(llLoop).sPdTimeExceedsFdTime = "Y" Then
                            slStr = slStr & "Y,"
                            slIsDaypart = True
                        Else
                            slStr = slStr & "N,"
                            slIsDaypart = False
                        End If
                         slStr = slStr & slEstDayAndTime
                        'Is the pledge before of after it is fed
                        If sArray(llLoop).sPdDayFed = "B" Then
                            slStr = slStr & """" & "B" & ""","
                        Else
                            slStr = slStr & """" & "A" & ""","
                        End If
                        slStr = slStr & """" & Trim$(sArray(llLoop).sTruePledgeDays) & ""","
                        If smUseActual = "Y" Then
                            slActualAirDateTime = Trim$(sArray(llLoop).sAirDate) & " " & Format(Trim$(sArray(llLoop).sAirTime), "hh:mm:ssam/pm")
                            slStr = slStr & slActualAirDateTime
                        Else
                            If slAirDateAndTime <> "" Then
                                slStr = slStr & slAirDateAndTime
                            Else
                                slStr = slStr & "NULL"
                            End If
                        End If
                        slStr = slStr & ","
                        slStr = slStr & lmAttCode & ","
                        slStr = slStr & """" & smShowVehName & ""","
                        If sArray(llLoop).iLstMon = 0 Then
                            slTemp = 0
                        Else
                            slTemp = 1
                        End If
                        If sArray(llLoop).iLstTue = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                        If sArray(llLoop).iLstWed = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                        If sArray(llLoop).iLstThu = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                        If sArray(llLoop).iLstFri = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                        If sArray(llLoop).iLstSat = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                        If sArray(llLoop).iLstSun = 0 Then
                            slTemp = slTemp & 0
                        Else
                            slTemp = slTemp & 1
                        End If
                         slStr = slStr & """" & slTemp & ""","
                        slStr = slStr & sArray(llLoop).lCntrNo & ","
                        'Acceptable flight start and end times
                        slStr = slStr & Format$(sArray(llLoop).sLstLnStartTime, "h:mm:ssa/p") & ","
                        slStr = slStr & Format$(sArray(llLoop).sLstLnEndTime, "h:mm:ssa/p") & ","
                        'Acceptable advertiser code
                        slStr = slStr & sArray(llLoop).iAdfCode & ","
                        'Is this a blackout, NO = 0, other than 0 it is a blackout
                        slStr = slStr & """" & sArray(llLoop).lLstBkoutLstCode & ""","
                        slStr = slStr & """" & Trim$(sArray(llLoop).sEmbeddedOrROS) & """"
                        
                        'D.S. 10/29/19 iLenOption is set above to let the web know if it needs to provide the user an option for either 10 or 15 second spots
                        'D.S. 10/29/19 The replace overwrites the TransType field with the lenOptions info: 0 = no choice, 1 = choice between 10 or 15 secs.
                        slStr = Replace(slStr, "~", ilLenOption)
                        '7458
                        If Not myEnt.Add(sFeedDate, sArray(llLoop).lgsfCode) Then
                            gLogMsg "ERROR: Failed to Create ENT Add record for games.", "WebExportLog.Txt", False
                        End If
                        imSpotCount = imSpotCount + 1
                        lgSTime18 = timeGetTime
                        Print #hmToDetail, gRemoveIllegalCharsAndLog(slStr, smToFileDetail, llLoop, False)
                        lgETime18 = timeGetTime
                        lgTtlTime18 = lgTtlTime18 + (lgETime18 - lgSTime18)
                        imNeedToSend = True
                        mCreateJelliRecord sArray(llLoop), sISCI, sCreative, sAdvt
                    End If
                End If
            ElseIf Not myEnt.Add(sFeedDate, sArray(llLoop).lgsfCode, Asts) Then
                gLogMsg "Warning: Records Pledge status = 2.", "WebExportLog.Txt", False
            End If
        End If
        tmAirTimeInfo(ilAirPlayNo).sBaseStartDate = slBaseStartDate
        tmAirTimeInfo(ilAirPlayNo).sBaseStartTime = slBaseStartTime
        tmAirTimeInfo(ilAirPlayNo).sEstimatedEndTime = smEstimatedEndTime
        tmAirTimeInfo(ilAirPlayNo).sEstimatedStartTime = smEstimatedStartTime
        tmAirTimeInfo(ilAirPlayNo).sPrevPldgETime = slPrevPldgETime
        tmAirTimeInfo(ilAirPlayNo).sPrevPldgSTime = slPrevPldgSTime
    Next llLoop
    If igExportSource = 2 Then DoEvents
    mBuildDetailRecs = True
Exit Function
ErrHand:
    'debug
    'Resume Next
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mBuildDetailRecs"
Exit Function
End Function

Private Function mBuildHeaders(cprst As ADODB.Recordset) As Integer

    '***** Important Note *******
    'If you change the way that the headers are built then you have to make the
    'same changes to the way the header is built in Stations and in Agreements.
    'They both put out headers when either station information is changed or
    'an Agreement's information is changed.

    Dim rst_Temp As ADODB.Recordset
    Dim slStr As String
    Dim slStr2 As String
    Dim slFTP As String
    Dim slTemp As String
    Dim ilWriteHeader As Integer
    Dim ilUpper As Integer
    Dim slEndDate As String
    
    On Error GoTo ErrHand
    lgSTime5 = timeGetTime
    imShttCode = cprst!shttCode
    smAttWebInterface = gGetWebInterface(lmAttCode)
    slStr = "ERROR"
    If udcCriteria.CSendEMails = vbChecked Then
        slStr = gBuildWebHeaders(cprst, imVefCode, smVefName, imShttCode, smAttWebInterface, True, smMode, smStartDate, smEndDate, smUseActual, smSuppressLog)
    Else
        slStr = gBuildWebHeaders(cprst, imVefCode, smVefName, imShttCode, smAttWebInterface, False, smMode, smStartDate, smEndDate, smUseActual, smSuppressLog)
    End If
    Print #hmToHeader, gRemoveIllegalCharsAndLog(slStr, smToFileHeader, 0, False)
    If igExportSource = 2 Then DoEvents
    lmFileHeaderCount = lmFileHeaderCount + 1
    lmLastattCode(UBound(lmLastattCode)) = lmWebAttID
    ReDim Preserve lmLastattCode(0 To UBound(lmLastattCode) + 1) As Long
    lgETime5 = timeGetTime
    lgTtlTime5 = lgTtlTime5 + (lgETime5 - lgSTime5)
    mBuildHeaders = True
Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mBuildHeaders"
    Exit Function
End Function


Private Function mAdjCpttRecs(cprst As ADODB.Recordset) As Integer
    
    Dim slTrueDropDate As String
    Dim slEndDate As String
    Dim slStartDate As String
    
    On Error GoTo ErrHand
    'Make sure to set the Posting Status to Outstanding
    SQLQuery = "UPDATE cptt SET "
    SQLQuery = SQLQuery + "cpttStatus = " & 0 & ", "
    SQLQuery = SQLQuery + "cpttPostingStatus = " & 0
    SQLQuery = SQLQuery + " WHERE cpttCode = " & cprst!cpttCode & ""
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        GoSub ErrHand:
    End If
    'Find the true date that the show stops airing for this agreemenet by saving
    'the lesser of the agreements Off Air Date and it's Drop Date
    If DateValue(gAdjYear(Trim$(cprst!attOffAir))) <= DateValue(gAdjYear(Trim$(cprst!attDropDate))) Then
        slTrueDropDate = Trim$(cprst!attOffAir)
    Else
        slTrueDropDate = Trim$(cprst!attDropDate)
    End If
    ' Delete ast records using Station, Vehicle, Dates
    If DateValue(gAdjYear(cprst!CpttStartDate)) < DateValue(gAdjYear(cprst!attOnAir)) Or DateValue(gAdjYear(cprst!CpttStartDate)) > DateValue(gAdjYear(slTrueDropDate)) Then
        slStartDate = Format$(DateValue(gAdjYear(cprst!CpttStartDate)), sgSQLDateForm)
        slEndDate = Format$(DateValue(gAdjYear(cprst!CpttStartDate)) + 6, sgSQLDateForm)
        SQLQuery = "DELETE FROM Ast"
        SQLQuery = SQLQuery & " WHERE astAtfCode = " & cprst!cpttatfCode
        SQLQuery = SQLQuery & " And astFeedDate >= '" & slStartDate & "'"
        SQLQuery = SQLQuery & " And astFeedDate <= '" & slEndDate & "'"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        End If
    End If
    'Determine if this is a valid week date for this agreement.
    'If not then delete the cptt record.
    If DateValue(gAdjYear(cprst!CpttStartDate)) < DateValue(gAdjYear(cprst!attOnAir)) Or DateValue(gAdjYear(cprst!CpttStartDate)) > DateValue(gAdjYear(slTrueDropDate)) Then
        SQLQuery = "DELETE FROM Cptt"
        SQLQuery = SQLQuery & " WHERE CpttCode = " & cprst!cpttCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        End If
    End If
    gFileChgdUpdate "cptt.mkd", True
    mAdjCpttRecs = True
Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "WebExportLog.txt", "Export CSI Web-mAdjCpttRecs"
    Exit Function
End Function

Sub mCloseFiles()

    Close #hmToDetail
    Close #hmToHeader
    Close #hmToCpyRot
    Close #hmToMultiUse
    Close #hmToEventInfo
    Exit Sub
End Sub

Private Function mReIndexServer() As Integer

    Dim ilRet As Integer
    Dim ilReIndexServer As Integer
    Dim ilRetry As Integer
    Dim WebCmds As New WebCommands
    Dim slResp As String

    On Error GoTo ErrHand
    If smCheckIniReIndex <> "" Then
        If CInt(smCheckIniReIndex) <> 1 Then
            SetResults "No Web Reindex was Attempted.", 0
            gLogMsg "No Web Reindex was Attempted. Turned Off in the INI File.", "WebExportLog.Txt", False
            gLogMsg "No Web Reindex was Attempted. Turned Off in the INI File.", "WebActivityLog.Txt", False
            mReIndexServer = True
            Exit Function
        End If
    End If
    mReIndexServer = False
    smReindex = "ReIndex" & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    gLogMsg "Instructing Web Site to Reindex SQL Server.", "WebExportLog.Txt", False
    gLogMsg "Instructing Web Site to Reindex SQL Server.", "WebActivityLog.Txt", False
    SetResults "Reindexing the Web Database.", 0
    If Not gExecExtStoredProc(smReindex, "Reindex.exe", False, False) Then
        SetResults "FAIL: Unable reindex the web database...", RGB(255, 0, 0)
        gLogMsg "FAIL: Unable to to reindex the web database...", "WebExportLog.Txt", False
        Exit Function
    End If
    ilRet = mExCheckWebWorkStatus(smReindex, "ReIndex")
    If Not ilRet Then
        For ilRetry = 0 To 4 Step 1
            'wait one minute
            Sleep (60000)
            ilRet = mExCheckWebWorkStatus(smReindex, "ReIndex")
            If ilRet Then
                Exit For
            End If
        Next ilRetry
    End If
    If ilRet = True Then
        gLogMsg "Reindexing Completed Successfully", "WebExportLog.Txt", False
        gLogMsg "Reindexing Completed Successfully", "WebActivityLog.Txt", False
        SetResults "Reindexing Completed Successfully.", 0
        mReIndexServer = True
    Else
        gLogMsg "Reindexing Failed", "WebExportLog.Txt", False
        gLogMsg "Reindexing Failed", "WebActivityLog.Txt", False
        SetResults "Reindexing Failed.", 0
    End If
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mReIndexServer"
    Exit Function
End Function

'***************************************************************************************
' JD 08-22-2007
' This function was added to handle a special case occurring in the function
' mCheckWebWorkStatus. We believe a network error is causing the error handler
' to fire. Adding retry code to the function mCheckWebWorkStatus itself did not
' seem feasable because we did not know where the error was actually occuring and
' simplying calling a resume next could cause even more trouble.
'
'***************************************************************************************
Private Function mExCheckWebWorkStatus(sFileName As String, sTypeExpected As String) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    On Error GoTo Err_Handler
    
    mExCheckWebWorkStatus = -1
    If (igDemoMode) Then
        mExCheckWebWorkStatus = 0
        Exit Function
    End If
    For ilLoop = 1 To 10
        ilRet = mCheckWebWorkStatus(sFileName, sTypeExpected)
        mExCheckWebWorkStatus = ilRet
        If ilRet <> -2 Then ' Retry only when this status is returned.
            Exit Function
        End If
        gLogMsg "mExCheckWebWorkStatus is retrying due to an error in mCheckWebWorkStatus", "WebExpRetryLog.txt", False
        If igExportSource = 2 Then DoEvents
        Sleep 2000  ' Delay for two seconds when retrying.
    Next
    If ilRet = -2 Then
        ilRet = -1  ' Keep the original error of -1 so all callers can process the error normally.
        gMsg = "A timeout has occured in frmWebExportSchdSpot - mExCheckWebWorkStatus"
        gLogMsg gMsg, "WebExpRetryLog.txt", False
        gLogMsg " ", "WebExpRetryLog.txt", False
    End If
    Exit Function
Err_Handler:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    mExCheckWebWorkStatus = -1
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "A general error has occured in frmWebExportSchdSpot - mExCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsg gMsg, "WebExportLog.txt", False
    gLogMsg " ", "WebExportLog.txt", False
    Exit Function
End Function

Private Function mCheckWebWorkStatus(sFileName As String, sTypeExpected As String) As Integer
    'input - sFilemane is the unique file name that is the key into the web
    'server database to check it's status
    'Web Server Status - 0 = Done, 1 = Working and 2 = Error
    'Loop while the web server is busy processing spots and emails
    'Check the server every 10 seconds Report status
    Dim llWaitTime As Long
    Dim ilModResult As Integer
    Dim imStatus As Integer
    Dim slResult As String
    Dim llNumRows As Long
    Dim ilTimedOut As Integer
    Dim ilRet As Integer
    Dim ilWaitValue As Integer
    'Debug information
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    'Number of Seconds to Sleep
    Const clNumSecsToSleep As Long = 1
    Const clSleepValue As Long = clNumSecsToSleep * cmOneSecond
    
    'Assuming clNumSecsToSleep is 10 then a mod value of 6 would
    'be 6 loops at 10 seconds each or 1 minute
    Const clModValue As Integer = 6
    On Error GoTo ErrHand
    mCheckWebWorkStatus = False
    If Not gHasWebAccess() Then
        Exit Function
    End If
    If sTypeExpected = "ReIndex" Or sTypeExpected = "WebEmails" Then
        ilWaitValue = 1350
    Else
        ilWaitValue = 2
    End If
    llWaitTime = 0
    imStatus = 1
    ilRet = False
    'Do While imStatus = 1 And llWaitTime < 1350 'We will wait 45 minutes based on 1350 - 1350/60 * 2 seconds
    'Do While imStatus = 1 And llWaitTime < ilWaitValue And ilRet = False
    Do While imStatus = 1 And llWaitTime < ilWaitValue And ilRet = False
        If igExportSource = 2 Then DoEvents
        If imTerminate Then
            gSetMousePointer grdVeh, grdVeh, vbDefault
            cmdCancel.Enabled = True
            imExporting = False
            SetResults "Export was canceled.", 0
            gLogMsg "** User Terminated **", "WebExportLog.Txt", False
            gLogMsg "** User Terminated **", "WebActivityLog.Txt", False
            Exit Function
        End If
        SQLQuery = "Select Count(*) from WorkStatus Where FileName = " & "'" & sFileName & "'"
        llNumRows = gExecWebSQLWithRowsEffected(SQLQuery)
        llWaitTime = llWaitTime + 1
        ilModResult = llWaitTime Mod clModValue
        If llNumRows = -1 Then
            'An error was returned
            imStatus = 2
            smStatus = "2"
        End If
        If llNumRows > 0 Then
            SQLQuery = "Select FileName, Status, Msg1, Msg2, DTStamp from WorkStatus Where FileName = " & "'" & sFileName & "'"
            'Get the status information from the web server database and write it to a file
            Call gRemoteExecSql(SQLQuery, smWebWorkStatus, "WebExports", True, True, 30)
            If igExportSource = 2 Then DoEvents
            smStatus = "1"
            ilRet = mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports", sTypeExpected)
            llWaitTime = llWaitTime + 1
            ilModResult = llWaitTime Mod clModValue
            imStatus = CInt(smStatus)
            'Handle Web Error Condition
            If imStatus = 2 Then
                If smPrevMsg <> smMsg2 Then
                    gLogMsg "   " & "The Web Server Returned an ERROR. See Below. ", "WebExportLog.Txt", False
                    gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
                End If
                If smPrevMsg <> smMsg2 Then
                    Call gEndWebSession("WebExportLog.Txt", "Y")
                Else
                    Call gEndWebSession("WebExportLog.Txt", "N")
                End If
                mCheckWebWorkStatus = False
                smPrevMsg = smMsg2
                If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                    gLogMsg myEnt.ErrorMessage, myEnt.fileName, False
                End If
                Exit Function
            End If
            If ilModResult = 0 And imStatus = 1 Then
                If igExportSource = 2 Then DoEvents
                gLogMsg "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
                gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
                If igExportSource = 2 Then DoEvents
            End If
        End If
        If imStatus = 1 Then
            Sleep clSleepValue
        End If
    Loop
    If llWaitTime >= 900 Then
        'We timed out
        gLogMsg "   " & "A timeout occured while waiting on the web server for a response.", "WebExportLog.Txt", False
        SetResults "A timeout waiting on a web server response.", RGB(255, 0, 0)
        Call gEndWebSession("WebExportLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
        
    End If
    'Show the final message with the totals of spots imported an emails sent
    'Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports")
    'Handle case that smStatus not a number to avoid error 13 (Type mismatch)
    imStatus = 0
    On Error Resume Next
    imStatus = CInt(smStatus)
    On Error GoTo ErrHand
    'Handle Web Error Condition
    If imStatus = 2 Then
        gLogMsg "   " & "The Web Server Returned an ERROR. See Below. ", "WebExportLog.Txt", False
        gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
        Call gEndWebSession("WebExportLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
    End If
    If ilRet Then
    gLogMsg "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
    gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebExportLog.Txt", False
    mCheckWebWorkStatus = True
    End If
Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    mCheckWebWorkStatus = -2
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mCheckWebWorkStatus"
    Exit Function
End Function

Private Function mProcessWebWorkStatusResults(sFileName As String, sIniValue As String, sTypeExpected) As Boolean

    Dim slLocation As String
    Dim hlFrom As Integer
    Dim ilRet  As Integer
    Dim ilLen As Integer
    Dim ilPos As Integer
    Dim slTemp As String
    Dim llCount As Long
    
    On Error GoTo ErrHand
    mProcessWebWorkStatusResults = False
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    slLocation = slLocation & sFileName
    On Error GoTo FileErrHand:
    hlFrom = FreeFile
    ilRet = 0
    Open slLocation For Input Access Read As hlFrom
    If ilRet <> 0 Then
        'D.S. J.D. 07/12/10 added 3 lines below
        gLogMsg "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable to open the file.", "WebExportLog.Txt", False
        smStatus = "1"
        Exit Function
    End If
    'Skip past the header record
    ilRet = 0
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Close #hlFrom
    
    
    If smMsg2 = "Execute Step 4 started." Then
        smMsg2 = smMsg2
    End If
    
    If ilRet <> 0 Then
        gLogMsg "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable read/input statement.", "WebExportLog.Txt", False
        smStatus = "1"
        Exit Function
    End If
    On Error GoTo ErrHand
    slTemp = smMsg2
    ilLen = Len(slTemp)
    ilPos = InStr(slTemp, ":") Or InStr(slTemp, ".")
    If ilPos > 0 Then
        llCount = Val(Mid$(slTemp, ilPos + 1, ilLen))
        If slTemp <> smMsg2 Then
            gLogMsg "   WorkStatus: " & smFileName & ", " & smStatus & ", " & smMsg1 & ", " & smMsg2 & ", " & smDTStamp, "WebExportLog.Txt", False
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Event Info Imported:") Then
            If Trim$(sTypeExpected) <> "WebEventInfo" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Event Info Imported:", "WebExportLog.Txt", False
                Exit Function
            End If
            If llCount <> rstWebQ!wqfAddCount Then
                gLogMsg "Error Counts Not Matching: WQF = " & rstWebQ!wqfAddCount & " Web = " & llCount, "WebExportLog.Txt", False
            End If
            lmTtlEventSpots = lmTtlEventSpots + rstWebQ!wqfAddCount
            lmWebTtlEventSpots = lmWebTtlEventSpots + llCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Comments Imported:") Then
            If Trim$(sTypeExpected) <> "WebComments" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Comments Imported:", "WebExportLog.Txt", False
                Exit Function
            End If
            If llCount <> rstWebQ!wqfAddCount Then
                gLogMsg "Error Counts Not Matching: WQF = " & rstWebQ!wqfAddCount & " Web = " & llCount, "WebExportLog.Txt", False
            End If
            lmWebTtlComments = lmWebTtlComments + llCount
            lmTtlComments = lmTtlComments + rstWebQ!wqfAddCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total MultiUse Imported:") Then
            If Trim$(sTypeExpected) <> "WebMultiUse" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total MultiUse Imported:", "WebExportLog.Txt", False
                Exit Function
            End If
            If llCount <> rstWebQ!wqfAddCount Then
                gLogMsg "Error Counts Not Matching: WQF = " & rstWebQ!wqfAddCount & " Web = " & llCount, "WebExportLog.Txt", False
            End If
            lmWebTtlMultiUse = lmWebTtlMultiUse + llCount
            lmWebTtlMultiUse = lmWebTtlMultiUse + rstWebQ!wqfAddCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Headers Imported:") Then
            If Trim$(sTypeExpected) <> "WebHeaders" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Headers Imported:", "WebExportLog.Txt", False
                Exit Function
            End If
            If llCount <> rstWebQ!wqfAddCount Then
                gLogMsg "Error Counts Not Matching: WQF = " & rstWebQ!wqfAddCount & " Web = " & llCount, "WebExportLog.Txt", False
            End If
            lmWebTtlHeaders = lmWebTtlHeaders + llCount
            lmTtlHeaders = lmTtlHeaders + rstWebQ!wqfAddCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Records Processed:") And sTypeExpected = "WebSpots" Then
            If Trim$(sTypeExpected) <> "WebSpots" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Records Processed:", "WebExportLog.Txt", False
                Exit Function
            End If
            If llCount <> rstWebQ!wqfAddCount Then
                gLogMsg "Error Counts Not Matching: WQF = " & rstWebQ!wqfAddCount & " Web = " & llCount, "WebExportLog.Txt", False
                '7458
                If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                     gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
                End If
            End If
            lmWebTtlSpots = lmWebTtlSpots + llCount
            'D.S. 2/9/18
            'D.S. 2/12/18
            'If bgTaskBlocked Then
                lmTotalAddSpotCount = lmTotalAddSpotCount + rstWebQ!wqfAddCount
                lmTotalDeleteSpotCount = lmTotalDeleteSpotCount + rstWebQ!wqfDelCount
            'End If
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Records Processed:") And sTypeExpected = "TotalSpots" Then
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Emails Sent:") Then
            If Trim$(sTypeExpected) <> "WebEmails" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Emails Sent:", "WebExportLog.Txt", False
                Exit Function
            End If
            lmWebTtlEmail = lmWebTtlEmail + llCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "ReIndex Complete.") Then
            If Trim$(sTypeExpected) <> "ReIndex" Then
                mProcessWebWorkStatusResults = False
                gLogMsg "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: ReIndex Complete.", "WebExportLog.Txt", False
                Exit Function
            End If
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
    End If
    Exit Function
FileErrHand:
    ilRet = -1
    Resume Next
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mProcessWebWorkStatusResults"
    Exit Function
End Function


Private Function mSendEmails() As Integer

    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If (igDemoMode) Then
        mSendEmails = True
        Exit Function
    End If
    If igExportSource = 2 Then DoEvents
    mSendEmails = False
    smSendEmails = "SendEmails" & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    gLogMsg "Instructing Web Site to Send Emails.", "WebExportLog.Txt", False
    gLogMsg "Instructing Web Site to Send Emails.", "WebActivityLog.Txt", False
    SetResults "Instructing Web to Send Emails.", 0
    If Not gExecExtStoredProc(smSendEmails, "SendEmails.exe", False, False) Then
        SetResults "FAIL: Unable Send Emails...", RGB(255, 0, 0)
        gLogMsg "FAIL: Unable Send Emails...", "WebExportLog.Txt", False
        gLogMsg "FAIL: Unable Send Emails...", "WebActivityLog.Txt", False
        Exit Function
    End If
    ilRet = mExCheckWebWorkStatus(smSendEmails, "WebEmails")
    If igExportSource = 2 Then DoEvents
    If ilRet = -1 Then
        gLogMsg "Send Emails Completed Successfully", "WebExportLog.Txt", False
        gLogMsg "Send Emails Completed Successfully", "WebActivityLog.Txt", False
        SetResults "Send Emails Completed Successfully.", 0
        mSendEmails = True
    Else
        gLogMsg "Send Emails Failed", "WebExportLog.Txt", False
        gLogMsg "Send Emails Failed", "WebActivityLog.Txt", False
        SetResults "Send Emails Failed.", 0
    End If
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-gSendEmails"
    Exit Function
End Function

Private Sub mExportedVefArray(ilVefCode As Integer)

    Dim ilIdx As Integer
    Dim ilFound As Integer

    On Error GoTo ErrHand
    ilFound = False
    For ilIdx = 0 To UBound(imExportedVefArray) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If imExportedVefArray(ilIdx) = ilVefCode Then
            ilFound = True
            Exit For
        End If
    Next ilIdx
    If ilFound = False Then
        imExportedVefArray(UBound(imExportedVefArray)) = ilVefCode
        ReDim Preserve imExportedVefArray(0 To UBound(imExportedVefArray) + 1)
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-gExportedVefArray"
    Exit Sub
End Sub


Public Function mWebGetPostedAttRecs(slSDate As String, slEDate As String) As Integer

    Dim sFTPAddress As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mWebGetPostedAttRecs = False
    gSetMousePointer grdVeh, grdVeh, vbHourglass
    mWebGetPostedAttRecs = True
    Call gLoadOption(sgWebServerSection, "FTPAddress", sFTPAddress)
    SQLQuery = "Select Distinct attCode FROM Spots WHERE PostedFlag = 1 "
    SQLQuery = SQLQuery & " AND PledgeStartDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND PledgeStartDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
    ilRet = gRemoteExecSql(SQLQuery, "WebPostedAttRecs.txt", "WebImports", True, True, 30)
    If ilRet Then
        mWebGetPostedAttRecs = True
    End If
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mWebGetPostedAttRecs"
    mWebGetPostedAttRecs = False
End Function

Private Function mWebProcessPostedAttRecs() As Integer

    Dim slWebAttRecs As String
    Dim slPathFileName As String
    Dim llCurMaxRecs As Long
    Dim llCount As Long
    Dim ilRet As Integer
    Dim slAttCode As String
    Dim hlFrom As Integer
    
    On Error GoTo ErrHand
    ReDim lmWebPostedAttRecs(0 To 10000) As Long
    llCurMaxRecs = 10000
    mWebProcessPostedAttRecs = True
    'Request the records
    gLogMsg "** Starting Getting Web ATT Records **", "WebExportLog.Txt", False
    slPathFileName = smWebImports & "WebPostedAttRecs.txt"
    'Open the and fill the array
    hlFrom = FreeFile
    On Error GoTo ErrFileHand
    Open slPathFileName For Input Access Read Lock Write As hlFrom
    If ilRet <> 0 Then
        mWebProcessPostedAttRecs = False
        Exit Function
    End If
    On Error GoTo ErrHand
    llCount = 0
    ' Skip past the header definition record.
    Input #hlFrom, slAttCode
    Do While Not EOF(hlFrom)
        Input #hlFrom, slAttCode
        If slAttCode <> "" Then
            If IsNumeric(slAttCode) Then
                lmWebPostedAttRecs(llCount) = CLng(slAttCode)
                llCount = llCount + 1
            End If
        End If
        
        If llCount = llCurMaxRecs Then
            llCurMaxRecs = llCurMaxRecs + 10000
            ReDim Preserve lmWebPostedAttRecs(0 To llCurMaxRecs) As Long
        End If
    Loop
    If llCount > 0 Then
        ReDim Preserve lmWebPostedAttRecs(0 To llCount - 1) As Long
        ArraySortTyp fnAV(lmWebPostedAttRecs(), 0), UBound(lmWebPostedAttRecs) + 1, 0, LenB(lmWebPostedAttRecs(1)), 0, -2, 0
    Else
        ReDim Preserve lmWebPostedAttRecs(0 To 0) As Long
    End If
    Close hlFrom
    gLogMsg "** Finished Getting Web ATT Records **", "WebExportLog.Txt", False
    Exit Function
ErrFileHand:
    ilRet = 1
    Resume Next
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mWebProcessPostedAttRecs"
    'Debug
    'Resume Next
    Exit Function
End Function
    

Private Function gBinarySearchWebPostedAttCodes(lCode As Long) As Long
    
    'D.S. 03/13/06
    'Returns the index number of the matching AttCode that was passed in.
    'If we find it then we don't want to export it again
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    llMin = LBound(lmWebPostedAttRecs)
    llMax = UBound(lmWebPostedAttRecs)
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If lCode = lmWebPostedAttRecs(llMiddle) Then
            'found the match
            gBinarySearchWebPostedAttCodes = llMiddle
            Exit Function
        ElseIf lCode < lmWebPostedAttRecs(llMiddle) Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchWebPostedAttCodes = -1
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-gBinarySearchWebPostedAttCodes"
    gBinarySearchWebPostedAttCodes = False
    Exit Function
    
End Function
Private Function gBinarySearchPostedAttCodes(lCode As Long) As Long
    
    'D.S. 03/13/06
    'Returns the index number of the matching AttCode that was passed in.
    'If we find it then we don't want to export it again
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    llMin = LBound(lmPostedAttRecs)
    llMax = UBound(lmPostedAttRecs)
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If lCode = lmPostedAttRecs(llMiddle) Then
            'found the match
            gBinarySearchPostedAttCodes = llMiddle
            Exit Function
        ElseIf lCode < lmPostedAttRecs(llMiddle) Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchPostedAttCodes = -1
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mBinarySearchPostedAttCodes"
    Exit Function
End Function


Public Function mGetPostedAttRecs(slSDate As String, slEDate As String) As Integer

    Dim ilRet As Integer
    Dim rst_Ast As ADODB.Recordset
    Dim llCount As Long
    Dim slAttCode As String
    Dim hlFrom As Integer
    Dim llCurMaxRecs As Long
    
    On Error GoTo ErrHand
    ReDim lmPostedAttRecs(0 To 100000) As Long
    llCurMaxRecs = 100000
    mGetPostedAttRecs = True
    gLogMsg "** Starting Getting Web ATT Records **", "WebExportLog.Txt", False
    SQLQuery = "SELECT Distinct astAtfCode FROM ast WHERE astAtfCode > 0"
    SQLQuery = SQLQuery & " And astFeedDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND astCPStatus = 1"
    Set rst_Ast = gSQLSelectCall(SQLQuery)
    llCount = 0
    While Not rst_Ast.EOF
        lmPostedAttRecs(llCount) = rst_Ast!astAtfCode
        llCount = llCount + 1
        If llCount = llCurMaxRecs Then
            llCurMaxRecs = llCurMaxRecs + 100000
            ReDim Preserve lmPostedAttRecs(0 To llCurMaxRecs) As Long
        End If
        rst_Ast.MoveNext
    Wend
    ReDim Preserve lmPostedAttRecs(0 To llCount - 1) As Long
    ArraySortTyp fnAV(lmPostedAttRecs(), 0), UBound(lmPostedAttRecs) + 1, 0, LenB(lmPostedAttRecs(1)), 0, -2, 0
    Close hlFrom
    gLogMsg "** Finished Getting Local ATT Records **", "WebExportLog.Txt", False
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mGetPostedAttRecs"
    Exit Function
End Function

Public Function mConvertSecsToTime(lNumSecs As Long) As String

    Dim slHr As String
    Dim slMin As String
    Dim slSec As String

    slHr = CDate(lNumSecs)
    slHr = Round(CStr(lNumSecs / 3600))
    slMin = Round(CStr(lNumSecs / 3600))
    slHr = Round(CStr(lNumSecs / 3600))
    slSec = Round(CStr(lNumSecs / 1))
    slMin = Round(CStr(lNumSecs / 60))
    mConvertSecsToTime = slHr & ":" & slMin & ":" & slSec
    Exit Function
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mConvertSecsToTime"
    Exit Function
    
End Function


Private Function mBuildEventInfo(lgsfCode As Long) As Integer

    
    Dim slFeedSource As String
    Dim slLanguageCode As String
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slVisitTeamName As String
    Dim slVisitTeamAbbr As String
    Dim slHomeTeamName As String
    Dim slHomeTeamAbbr As String
    Dim slGameDate As String
    Dim slGameStartTime As String
    Dim slPledgeByEvent As String
    Dim slStr As String
    Dim llLoop As Long
    Dim llGameIdx As Long
    Dim llUpper As Long
    Dim ilVff As Integer
    Dim ilVefCode As Integer
    Dim slEventCarried As String
    Dim ilTimeAdj As Integer
    Dim ilShttCode As Integer
    Dim slAttCode As String
    Dim slFed As String
    
    On Error GoTo ErrHand
    lgSTime17 = timeGetTime
    'now we sent send game information for every agreement that uses sports
    'rather than only once per export
    'Get out if we already have the incoming gsfcode
    For llLoop = 0 To UBound(tgGameInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgGameInfo(llLoop).lgsfCode = lgsfCode And tgGameInfo(llLoop).lAttCode = lmAttCode Then
            Exit Function
        End If
    Next llLoop
    If igExportSource = 2 Then DoEvents
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE gsfCode = " & lgsfCode
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    If igExportSource = 2 Then DoEvents
    If Not rst_Gsf.EOF Then
        'Feed Source
        slFeedSource = ""
        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
            If rst_Gsf!gsfFeedSource = "V" Then
                slFeedSource = "Visting"
            Else
                slFeedSource = "Home"
            End If
        End If
        'Language
        slLanguageCode = ""
        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
            For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tgLangInfo(ilLang).iCode = rst_Gsf!gsfLangMnfCode Then
                    slLanguageCode = Trim$(tgLangInfo(ilLang).sName)
                    Exit For
                End If
            Next ilLang
        End If
        'Visiting Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                If igExportSource = 2 Then DoEvents
                slVisitTeamName = Trim$(tgTeamInfo(ilTeam).sName)
                slVisitTeamAbbr = Trim$(tgTeamInfo(ilTeam).sShortForm)
                Exit For
            End If
        Next ilTeam
        'Home Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                If igExportSource = 2 Then DoEvents
                slHomeTeamName = Trim$(tgTeamInfo(ilTeam).sName)
                slHomeTeamAbbr = Trim$(tgTeamInfo(ilTeam).sShortForm)
                Exit For
            End If
        Next ilTeam
        'Air Date
        slGameDate = Format$(rst_Gsf!gsfAirDate, "m/d/yyyy")
        'TTP 9645 - adj. date and time for the correct time zone
        'Start Time
        slAttCode = CStr(lmAttCode)
        ilShttCode = gGetShttCodeFromAttCode(slAttCode)
        ilTimeAdj = gGetTimeAdj(ilShttCode, CLng(imVefCode), slFed)
        slGameStartTime = rst_Gsf!gsfAirTime
        gAdjustEventTime ilTimeAdj, slGameDate, slGameStartTime
        slGameStartTime = Format$(slGameStartTime, sgShowTimeWSecForm)
    End If
    llGameIdx = UBound(tgGameInfo)
    tgGameInfo(llGameIdx).lgsfCode = lgsfCode
    tgGameInfo(llGameIdx).sGameDate = Format$(slGameDate, "m/d/yyyy")
    tgGameInfo(llGameIdx).sGameStartTime = slGameStartTime
    tgGameInfo(llGameIdx).sVisitTeamName = slVisitTeamName
    tgGameInfo(llGameIdx).sVisitTeamAbbr = slVisitTeamAbbr
    tgGameInfo(llGameIdx).sHomeTeamName = slHomeTeamName
    tgGameInfo(llGameIdx).sHomeTeamAbbr = slHomeTeamAbbr
    tgGameInfo(llGameIdx).sLanguageCode = slLanguageCode
    tgGameInfo(llGameIdx).sFeedSource = slFeedSource
    'Get the Declare Status
    SQLQuery = "SELECT * FROM Pet WHERE petAttCode = " & lmAttCode & " And petGsfCode = " & lgsfCode
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    'D.S. 10/31/12 Per Dick and Jeff always send a "U" out because Jeff used radio buttons
    'instead of a toggle
    slEventCarried = "U"
    tgGameInfo(llGameIdx).sEventCarried = slEventCarried
    tgGameInfo(llGameIdx).lAttCode = lmWebAttID
    ReDim Preserve tgGameInfo(0 To llGameIdx + 1)
    lgETime17 = timeGetTime
    lgTtlTime17 = lgTtlTime17 + (lgETime17 - lgSTime17)
    If igExportSource = 2 Then DoEvents
Exit Function

ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mBuildEventInfo"
    Exit Function
End Function

Private Sub mWriteEventInfoFile()

    Dim ilLoop As Integer
    Dim slStr As String
    
    On Error GoTo ErrHand
    For ilLoop = 0 To UBound(tgGameInfo) - 1 Step 1
        slStr = tgGameInfo(ilLoop).lgsfCode & ","
        slStr = slStr & Trim$(tgGameInfo(ilLoop).sGameDate) & ","
        slStr = slStr & Trim$(tgGameInfo(ilLoop).sGameStartTime) & ","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sVisitTeamName) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sVisitTeamAbbr) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sHomeTeamName) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sHomeTeamAbbr) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sLanguageCode) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sFeedSource) & ""","
        slStr = slStr & """" & Trim$(tgGameInfo(ilLoop).sEventCarried) & ""","
        slStr = slStr & tgGameInfo(ilLoop).lAttCode
        lmTtlEventSpots = lmTtlEventSpots + 1
        Print #hmToEventInfo, gRemoveIllegalCharsAndLog(slStr, smToEventInfo, ilLoop + 1, False)
        lmFileEventCount = lmFileEventCount + 1
    Next ilLoop
Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mWriteEventInfoFile"
    Exit Sub
End Sub

Private Sub mInsertIntoWQF()

    Dim ilRet As Integer
    Dim ilRetry As Integer
    Dim ilRetries As Integer
    Dim ilInsertFailed As Integer

    If imTerminate Then
        gSetMousePointer grdVeh, grdVeh, vbDefault
        cmdCancel.Enabled = True
        imExporting = False
        SetResults "Export was canceled.", 0
        gLogMsg "** User Terminated **", "WebExportLog.Txt", False
        gLogMsg "** User Terminated **", "WebActivityLog.Txt", False
        Exit Sub
    End If
    ilRetries = 0
    ilInsertFailed = True
    If lmFileCommentCount > 0 Then
        Do While ilInsertFailed And ilRetries < 5
            ilInsertFailed = False
            SQLQuery = "Insert Into WQF_Web_Queue ( "
            SQLQuery = SQLQuery & "wqfCode, "
            SQLQuery = SQLQuery & "wqfFileName, "
            SQLQuery = SQLQuery & "wqfExeToRun, "
            SQLQuery = SQLQuery & "wqfTypeExpected, "
            SQLQuery = SQLQuery & "wqfFTPStatus, "
            SQLQuery = SQLQuery & "wqfProcStatus, "
            SQLQuery = SQLQuery & "wqfAddCount, "
            SQLQuery = SQLQuery & "wqfDelCount "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Trim$(smWebCopyRot) & "', "
            SQLQuery = SQLQuery & "'" & "ImportCopyRotCom.exe" & "', "
            SQLQuery = SQLQuery & "'" & "WebComments" & "', "
            If imFTPEvents Then
                SQLQuery = SQLQuery & 1 & ","
            Else
            SQLQuery = SQLQuery & 0 & ","
            End If
            SQLQuery = SQLQuery & 0 & ","
            SQLQuery = SQLQuery & lmFileCommentCount & ","
            SQLQuery = SQLQuery & 0
            SQLQuery = SQLQuery & ") "
    
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            Else
                ilInsertFailed = False
                End If
            ilRetries = ilRetries + 1
        Loop
    Else
        ilInsertFailed = False
    End If
    If ilRetries = 5 And ilInsertFailed = False Then
        Exit Sub
    End If
    lmFileCommentCount = 0
    ilRetries = 0
    ilInsertFailed = True

    If lmFileMultiUseCount > 0 Then
        Do While ilInsertFailed And ilRetries < 5
            ilInsertFailed = False
            SQLQuery = "Insert Into WQF_Web_Queue ( "
            SQLQuery = SQLQuery & "wqfCode, "
            SQLQuery = SQLQuery & "wqfFileName, "
            SQLQuery = SQLQuery & "wqfExeToRun, "
            SQLQuery = SQLQuery & "wqfTypeExpected, "
            SQLQuery = SQLQuery & "wqfFTPStatus, "
            SQLQuery = SQLQuery & "wqfProcStatus, "
            SQLQuery = SQLQuery & "wqfAddCount, "
            SQLQuery = SQLQuery & "wqfDelCount "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Trim$(smWebMultiUse) & "', "
            SQLQuery = SQLQuery & "'" & "ImportMultiFile.exe" & "', "
            SQLQuery = SQLQuery & "'" & "WebMultiUse" & "', "
            If imFTPEvents Then
                SQLQuery = SQLQuery & 1 & ","
            Else
            SQLQuery = SQLQuery & 0 & ","
            End If
            SQLQuery = SQLQuery & 0 & ","
            SQLQuery = SQLQuery & lmFileMultiUseCount & ","
            SQLQuery = SQLQuery & 0
            SQLQuery = SQLQuery & ") "

            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            Else
                ilInsertFailed = False
                End If
            ilRetries = ilRetries + 1
        Loop
    Else
        ilInsertFailed = False
    End If
    If ilRetries = 5 And ilInsertFailed = False Then
        Exit Sub
    End If
    lmFileMultiUseCount = 0
    ilRetries = 0
    ilInsertFailed = True
    'Check bit map to see if using games.  If not no sense in exporting of showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        gLogMsg "Instructing Web Site to Import Event Info.", "WebActivityLog.Txt", False
        ilRetries = 0
        ilInsertFailed = True
        If lmFileEventCount > 0 Then
            Do While ilInsertFailed And ilRetries < 5
                ilInsertFailed = False
                SQLQuery = "Insert Into WQF_Web_Queue ( "
                SQLQuery = SQLQuery & "wqfCode, "
                SQLQuery = SQLQuery & "wqfFileName, "
                SQLQuery = SQLQuery & "wqfExeToRun, "
                SQLQuery = SQLQuery & "wqfTypeExpected, "
                SQLQuery = SQLQuery & "wqfFTPStatus, "
                SQLQuery = SQLQuery & "wqfProcStatus, "
                SQLQuery = SQLQuery & "wqfAddCount, "
                SQLQuery = SQLQuery & "wqfDelCount "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & 0 & ", "
                SQLQuery = SQLQuery & "'" & Trim$(smWebEventInfo) & "', "
                SQLQuery = SQLQuery & "'" & "ImportGameInfo.exe" & "', "
                SQLQuery = SQLQuery & "'" & "WebEventInfo" & "', "
                If imFTPEvents Then
                    SQLQuery = SQLQuery & 1 & ","
                Else
                SQLQuery = SQLQuery & 0 & ","
                End If
                SQLQuery = SQLQuery & 0 & ","
                SQLQuery = SQLQuery & lmFileEventCount & ","
                SQLQuery = SQLQuery & 0
                SQLQuery = SQLQuery & ") "
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    GoSub ErrHand:
                Else
                    ilInsertFailed = False
                End If
                ilRetries = ilRetries + 1
            Loop
        Else
            ilInsertFailed = False
        End If
        If ilRetries = 5 And ilInsertFailed = False Then
            Exit Sub
        End If
    End If
    lmFileEventCount = 0
    ilRetries = 0
    ilInsertFailed = True
    If lmFileHeaderCount > 0 Then
        Do While ilInsertFailed And ilRetries < 5
            ilInsertFailed = False
            SQLQuery = "Insert Into WQF_Web_Queue ( "
            SQLQuery = SQLQuery & "wqfCode, "
            SQLQuery = SQLQuery & "wqfFileName, "
            SQLQuery = SQLQuery & "wqfExeToRun, "
            SQLQuery = SQLQuery & "wqfTypeExpected, "
            SQLQuery = SQLQuery & "wqfFTPStatus, "
            SQLQuery = SQLQuery & "wqfProcStatus, "
            SQLQuery = SQLQuery & "wqfAddCount, "
            SQLQuery = SQLQuery & "wqfDelCount "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Trim$(smWebHeader) & "', "
            SQLQuery = SQLQuery & "'" & "ImportHeaders.exe" & "', "
            SQLQuery = SQLQuery & "'" & "WebHeaders" & "', "
            If imFTPEvents Then
                SQLQuery = SQLQuery & 1 & ","
            Else
            SQLQuery = SQLQuery & 0 & ","
            End If
            SQLQuery = SQLQuery & 0 & ","
            SQLQuery = SQLQuery & lmFileHeaderCount & ","
            SQLQuery = SQLQuery & 0
            SQLQuery = SQLQuery & ") "
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            Else
                ilInsertFailed = False
            End If
            ilRetries = ilRetries + 1
        Loop
    Else
        ilInsertFailed = False
    End If
    If ilRetries = 5 And ilInsertFailed = False Then
        Exit Sub
    End If
    lmFileHeaderCount = 0
    ilRetries = 0
    ilInsertFailed = True
    If lmFileAddSpotCount > 0 Then
        Do While ilInsertFailed And ilRetries < 5
            ilInsertFailed = False
            SQLQuery = "Insert Into WQF_Web_Queue ( "
            SQLQuery = SQLQuery & "wqfCode, "
            SQLQuery = SQLQuery & "wqfFileName, "
            SQLQuery = SQLQuery & "wqfExeToRun, "
            SQLQuery = SQLQuery & "wqfTypeExpected, "
            SQLQuery = SQLQuery & "wqfFTPStatus, "
            SQLQuery = SQLQuery & "wqfProcStatus, "
            SQLQuery = SQLQuery & "wqfAddCount, "
            SQLQuery = SQLQuery & "wqfDelCount "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Trim$(smWebSpots) & "', "
            SQLQuery = SQLQuery & "'" & "ImportSpots.exe" & "', "
            SQLQuery = SQLQuery & "'" & "WebSpots" & "', "
            If imFTPEvents Then
                SQLQuery = SQLQuery & 1 & ","
            Else
            SQLQuery = SQLQuery & 0 & ","
            End If
            SQLQuery = SQLQuery & 0 & ","
            SQLQuery = SQLQuery & lmFileAddSpotCount & ", "
            SQLQuery = SQLQuery & lmFileDeleteSpotCount
            SQLQuery = SQLQuery & ") "
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            Else
                ilInsertFailed = False
            End If
            ilRetries = ilRetries + 1
        Loop
    Else
        ilInsertFailed = False
    End If
    If ilRetries = 5 And ilInsertFailed = False Then
        Exit Sub
    End If
    lmFileAddSpotCount = 0
    lmFileDeleteSpotCount = 0
    Exit Sub
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mInsertIntoWQF"
    Exit Sub
End Sub

Private Sub mProcessWebQueue()

    Dim ilRet As Integer
    Dim ilRetry As Integer
    Dim ilLoop As Integer
    Dim ilAllWentOK As Integer
    Dim slTemp As String
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim ilCount As Integer
    Dim slType As String
    Dim WebCmds As New WebCommands
    Dim slResp As String
    Dim slResponse As String
    
    On Error GoTo ErrHand:
    ilCount = 1
    slType = ""
    'While slType <> "ImportSpots.exe" Or ilCount = 1
    SQLQuery = "SELECT * FROM WQF_Web_Queue WHERE wqfProcStatus = 0 And wqfFTPStatus = 1 order by wqfCode asc"
    Set rstWebQ = gSQLSelectCall(SQLQuery)

    If rstWebQ.EOF Then
        imSomeThingToDo = False
        Exit Sub
    End If
    imSomeThingToDo = True
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
        ilCount = ilCount + 1
    smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & ".txt"
    Call mWaitForWebLock
    ilAllWentOK = True
    
    slType = Trim$(rstWebQ!wqfExeToRun)
    SetResults "- Imp. " & rstWebQ!wqfExeToRun, 0
    If Not gExecExtStoredProc(rstWebQ!wqffilename, rstWebQ!wqfExeToRun, False, False) Then
        SetResults "FAIL: Unable to instruct Web site to run " & rstWebQ!wqfExeToRun, RGB(255, 0, 0)
        gLogMsg "FAIL: Unable to instruct Web site to run " & rstWebQ!wqfExeToRun, "WebExportLog.Txt", False
        cmdExport.Enabled = False
        cmdCancel.Caption = "&Done"
        gSetMousePointer grdVeh, grdVeh, vbDefault
        ilAllWentOK = False
    Else
        gLogMsg "Importing " & rstWebQ!wqffilename, "WebExportLog.Txt", False
        gLogMsg "Importing " & rstWebQ!wqffilename, "WebActivityLog.Txt", False
        imWaiting = True
    End If
    Exit Sub
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mProcessWebQueue"
    'debug
    'Resume Next
    Exit Sub

End Sub

Private Function mCheckStatus() As Integer

    Dim ilRet As Integer
    Dim slFile As String
    Dim slTemp As String
    Dim ilPos As Integer
    
    On Error GoTo ErrHand
    If (igDemoMode) Then
        imWaiting = False
        Exit Function
    End If
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    On Error GoTo ErrHand2:
    slTemp = smWebExports & Trim$(rstWebQ!wqffilename)
    '8886
'    slFile = Dir(slTemp)
'    If Len(slFile) = 0 Then
    If gFileExist(slTemp) = FILEEXISTSNOT Then
        Exit Function
    End If
    On Error GoTo ErrHand
    If rstWebQ.EOF Then
        imWaiting = False
        Exit Function
    End If
    ilRet = mExCheckWebWorkStatus(Trim$(rstWebQ!wqffilename), Trim$(rstWebQ!wqfTypeExpected))
    If ilRet = True Then
        gLogMsg Trim$(rstWebQ!wqffilename) & " Import Successful.", "WebExportLog.Txt", False
        gLogMsg Trim$(rstWebQ!wqffilename) & " Import Successful.", "WebActivityLog.Txt", False
        SetResults "     -- " & Trim$(rstWebQ!wqfTypeExpected) & " Imp. Successful.", 0
        SQLQuery = "Update WQF_Web_Queue Set wqfProcStatus = 1 where wqfCode = " & rstWebQ!wqfCode
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        End If
        
        ilPos = 0
        ilPos = InStr(slTemp, "WebSpots")
        If ilPos > 0 Then
            ilRet = mSetVefExpDate(rstWebQ!wqffilename)
        End If
        If Trim$(rstWebQ!wqfTypeExpected) = "webspots" Then
            ' JD 06-02-2015 - This code needed to be changed back to it's original behavoir. Calling this
            ' each time causes multiple emails to be sent to clients due to not waiting until the very end.
        
            ' Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports", "WebEmails")
            ' mSendEmails
        imWaiting = False
        Call gEndWebSession("WebExportLog.Txt")
    Else
            imWaiting = False
            Call gEndWebSession("WebExportLog.Txt")
        End If
    Else
        imWaiting = True
    End If
    Exit Function
    
ErrHand2:
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mCheckStatus"
    'debug
    'Resume Next
    Exit Function
End Function


Private Function mInitFTP() As Boolean

    Dim slTemp As String
    Dim ilRet As Integer
    Dim slSection As String
    
    mInitFTP = False
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    
    'Support for CSI_Utils FTP functions
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tmCsiFtpInfo.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tmCsiFtpInfo.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tmCsiFtpInfo.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tmCsiFtpInfo.sPWD)
    Call gLoadOption(sgWebServerSection, "WebExports", tmCsiFtpInfo.sSendFolder)
    Call gLoadOption(sgWebServerSection, "WebImports", tmCsiFtpInfo.sRecvFolder)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tmCsiFtpInfo.sServerDstFolder)
    Call gLoadOption(sgWebServerSection, "FTPExportDir", tmCsiFtpInfo.sServerSrcFolder)
    Call gLoadOption(slSection, "DBPath", tmCsiFtpInfo.sLogPathName)
    tmCsiFtpInfo.sLogPathName = Trim$(tmCsiFtpInfo.sLogPathName) & "\" & "Messages\FTPLog.txt"
    ilRet = csiFTPInit(tmCsiFtpInfo)
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tgCsiFtpFileListing.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tgCsiFtpFileListing.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tgCsiFtpFileListing.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tgCsiFtpFileListing.sPWD)
    Call gLoadOption(slSection, "DBPath", tgCsiFtpFileListing.sLogPathName)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tgCsiFtpFileListing.sPathFileMask)
    Exit Function
End Function
Public Function mCheckFTPStatus() As Boolean

    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    If (igDemoMode) Then
        imFtpInProgress = False
        Exit Function
    End If
    mCheckFTPStatus = False
    imFtpInProgress = True
    ilRet = csiFTPGetStatus(tmCsiFtpStatus)
    '1 = Busy, 0 = Not Busy
    If tmCsiFtpStatus.iState = 1 Then
        Exit Function
    Else
        If tmCsiFtpStatus.iStatus <> 0 Then
            ' Errors occured.
            ilRet = csiFTPGetError(tmCsiFtpErrorInfo)
            MsgBox "FTP Failed. " & tmCsiFtpErrorInfo.sInfo
            gLogMsg "Error: " & "FAILED to FTP " & tmCsiFtpErrorInfo.sFileThatFailed, "WebExportLog.Txt", False
            gLogMsg "Error: " & "FAILED to FTP " & tmCsiFtpErrorInfo.sFileThatFailed, "WebActivityLog.Txt", False
            SetResults "FTP Failed. ", 0
            '7458
            If Not myEnt.UpdateIncompleteByFilename(EntError) Then
                 gLogMsg myEnt.ErrorMessage, "WebExportLog.Txt", False
            End If
            If tmCsiFtpStatus.iStatus = 2 Then
                ' JD 01-05-22 TTP: 10372
                ' When this is the case, we need to exit out. Otherwise the affiliate will just keep failing
                ' and the user will have to use task manager to kill it.
                imFtpInProgress = False
                mCheckFTPStatus = False
            End If
            Exit Function
        Else
            ' JD 01-05-22 TTP: 10372
            ' Added code here to prevent the affilate from getting into an endless loop.
            If tmCsiFtpStatus.iJobCount = 0 And UBound(mFtpArray) < 1 Then
                imFtpInProgress = False
                mCheckFTPStatus = False
                Exit Function
            End If
            
            For ilLoop = 0 To UBound(mFtpArray) - 1 Step 1
                ilRet = gTestFTPFileExists(mFtpArray(ilLoop))
                If ilRet = 1 Then
                    SQLQuery = "Update WQF_Web_Queue Set wqfFtpStatus = 1 Where wqfFileName = " & "'" & mFtpArray(ilLoop) & "'"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    SetResults "Success, FTP - " & mFtpArray(ilLoop), 0
                    gLogMsg "   Success, FTP - " & mFtpArray(ilLoop), "WebActivityLog.Txt", False
                    imFtpInProgress = False
                    mCheckFTPStatus = True
                End If
            Next ilLoop
        End If
    End If
    Exit Function
    
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mCheckFTPStatus"
    'debug
    'Resume Next
    Exit Function
End Function


Private Sub cmdFTPTest_Click()
    Dim ilRet As Integer
    Dim FTPInfo As CSIFTPINFO
    Dim FTPStatus As CSIFTPSTATUS
    Dim FTPErrorInfo As CSIFTPERRORINFO
    
    FTPInfo.sIPAddress = "172.16.10.10"
    FTPInfo.sUID = "administrator"
    FTPInfo.sPWD = "cps6x52"
    FTPInfo.nPort = 21
    FTPInfo.sSendFolder = sgRootDrive & "CSI\Dev\CSI_UTILS\Release\TestIt"
    FTPInfo.sRecvFolder = sgRootDrive & "CSI\Dev\CSI_UTILS\Release\TestIt"
    FTPInfo.sServerDstFolder = "\WebAffiliate\AffWeb_Dev\Import"
    FTPInfo.sServerSrcFolder = "\WebAffiliate\AffWeb_Dev\Export"
    FTPInfo.sLogPathName = sgRootDrive & "CSI\Dev\CSI_UTILS\Release\TestIt\FTPLog.txt"
    ilRet = csiFTPInit(FTPInfo)
    ' Send the following files to the server.
    ilRet = csiFTPFileToServer("FTPTestFile_1.txt")
    ilRet = csiFTPFileToServer("FTPTestFile_2.txt")
    ilRet = csiFTPFileToServer("FTPTestFile_3.txt")
    ilRet = csiFTPFileToServer("FTPTestFile_4.txt")
    ilRet = csiFTPGetStatus(FTPStatus)
    While FTPStatus.iState = 1
        If igExportSource = 2 Then DoEvents
        Sleep (200)
        ilRet = csiFTPGetStatus(FTPStatus)
    Wend
    If FTPStatus.iStatus <> 0 Then
        ' Errors occured.
        ilRet = csiFTPGetError(FTPErrorInfo)
        MsgBox "FTP Failed. " & FTPErrorInfo.sInfo
        MsgBox "The file name was " & FTPErrorInfo.sFileThatFailed
        Exit Sub
    End If
    MsgBox "Send Complete"
 
    ' Receive the following files from the server.
    ilRet = csiFTPFileFromServer("FTPTestFile_1.txt")
    ilRet = csiFTPFileFromServer("FTPTestFile_2.txt")
    ilRet = csiFTPFileFromServer("FTPTestFile_3.txt")
    ilRet = csiFTPFileFromServer("FTPTestFile_4.txt")
    ilRet = csiFTPGetStatus(FTPStatus)
    While FTPStatus.iState = 1
        If igExportSource = 2 Then DoEvents
        Sleep (200)
        ilRet = csiFTPGetStatus(FTPStatus)
    Wend
    If FTPStatus.iStatus <> 0 Then
        ' Errors occured.
        ilRet = csiFTPGetError(FTPErrorInfo)
        MsgBox "FTP Failed. " & FTPErrorInfo.sInfo
        MsgBox "The file name was " & FTPErrorInfo.sFileThatFailed
        Exit Sub
    End If
    MsgBox "Recv Complete"
End Sub

Private Function mObtainRotEndDate(lCode As Long) As String

    Dim rst As ADODB.Recordset
        
    On Error GoTo ErrHand
    mObtainRotEndDate = ""
    SQLQuery = "SELECT CifRotEndDate FROM CIF_Copy_Inventory WHERE cifCode = " & lCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst!cifRotEndDate) Then
            mObtainRotEndDate = Trim$(rst!cifRotEndDate)
        Else
            mObtainRotEndDate = ""
        End If
    Else
        mObtainRotEndDate = ""
    End If
Exit Function

ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mObtainRotEndDate"
    Exit Function
End Function

Private Function mEstimatedDayAndTime(sArray() As ASTINFO, lLoop As Long) As String

    Dim slResults As String
    Dim slMonDate As String
    Dim slEstDate As String
    Dim ilDayNum As Integer
    Dim ilestDay As Integer
    
    If igExportSource = 2 Then DoEvents
    mEstimatedDayAndTime = ""
    SQLQuery = "SELECT datCode "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datCode = " & sArray(lLoop).lDatCode & ")"
    Set rst_Est = gSQLSelectCall(SQLQuery)
    
    If Not rst_Est.EOF Then
        Select Case Weekday(Trim$(sArray(lLoop).sPledgeDate))
            Case vbSunday
                slResults = "Su"
                ilDayNum = 6
            Case vbMonday
                slResults = "Mo"
                ilDayNum = 0
            Case vbTuesday
                slResults = "Tu"
                ilDayNum = 1
            Case vbWednesday
                slResults = "We"
                ilDayNum = 2
            Case vbThursday
                slResults = "Th"
                ilDayNum = 3
            Case vbFriday
                slResults = "Fr"
                ilDayNum = 4
            Case vbSaturday
                slResults = "Sa"
                ilDayNum = 5
            Case Else
                slResults = "??"
        End Select
    
        SQLQuery = "SELECT eptEstimatedDay, eptEstimatedTime "
        SQLQuery = SQLQuery + " FROM ept"
        SQLQuery = SQLQuery + " WHERE (eptDatCode = " & rst_Est!datCode
        SQLQuery = SQLQuery + " AND eptFdAvailTime = " & "'" & Format$(Trim$(sArray(lLoop).sFeedTime), sgSQLTimeForm) & "'"
        
        ' TTP 10633 JD 01-27-2023
        ' Prev: SQLQuery = SQLQuery + " AND eptFdAvailDay = " & "'" & slResults & "')"
        ' This allows the user to use a pledge date that is not on the same day as the feed date.
        SQLQuery = SQLQuery + " AND eptEstimatedDay = " & "'" & slResults & "')"
        
        Set rst_Est = gSQLSelectCall(SQLQuery)
    End If
    If igExportSource = 2 Then DoEvents
    If Not rst_Est.EOF Then
        If sArray(lLoop).sPdDayFed = "B" Then
            Select Case rst_Est!eptEstimatedDay
                Case "Mo"
                  ilestDay = 0
                Case "Tu"
                 ilestDay = 1
                Case "We"
                 ilestDay = 2
                Case "Th"
                 ilestDay = 3
                Case "Fr"
                 ilestDay = 4
                Case "Sa"
                 ilestDay = 5
                Case "Su"
                 ilestDay = 6
            End Select
            If ilestDay > ilDayNum Then
              ilestDay = ilestDay - 7
            End If
            slEstDate = DateAdd("d", ilestDay - ilDayNum, sArray(lLoop).sPledgeDate)
            smEstimatedDate = Trim$(slEstDate)
            smEstimatedStartTime = Format$(rst_Est!eptEstimatedTime, "h:mm:ssa/p")
            slResults = """" & smEstimatedDate & """" & "," & smEstimatedStartTime & ","
        Else
            Select Case rst_Est!eptEstimatedDay
                Case "Mo"
                  ilestDay = 0
                Case "Tu"
                 ilestDay = 1
                Case "We"
                 ilestDay = 2
                Case "Th"
                 ilestDay = 3
                Case "Fr"
                 ilestDay = 4
                Case "Sa"
                 ilestDay = 5
                Case "Su"
                 ilestDay = 6
            End Select
            If ilestDay < ilDayNum Then
              ilestDay = ilestDay + 7
            End If
            slEstDate = DateAdd("d", ilestDay - ilDayNum, sArray(lLoop).sPledgeDate)
            smEstimatedDate = Trim$(slEstDate)
            smEstimatedStartTime = Format$(rst_Est!eptEstimatedTime, "h:mm:ssa/p")
            slResults = """" & smEstimatedDate & """" & "," & smEstimatedStartTime & ","
        End If
    Else
        slResults = ""
    End If
    mEstimatedDayAndTime = slResults
    If igExportSource = 2 Then DoEvents
End Function

Private Function mGetMissedReasons() As Integer

    Dim ilRet As Integer
    Dim rst As ADODB.Recordset
    Dim slResults As String
    Dim slStr As String
    Dim llLineNum As Long
    
    On Error GoTo Err_Handler
    mGetMissedReasons = False
    lmTtlMultiUse = 0
    Print #hmToMultiUse, "[MissedReasons]"
    Print #hmToMultiUse, "Code, Reason, IsDefault"
    SQLQuery = "select mnfCode, mnfName, mnfUnitType from MNF_Multi_Names where mnftype = 'M' and (mnfCodeStn = 'A' or mnfCodeStn = 'B')"
    Set rst = gSQLSelectCall(SQLQuery)
    llLineNum = 1
    While Not rst.EOF
        slStr = Trim$(rst!mnfCode) & "," & """" & Trim$(rst!mnfName) & """" & "," & """" & Trim$(rst!mnfUnitType) & """"
        Print #hmToMultiUse, gRemoveIllegalCharsAndLog(slStr, smToMultiUse, llLineNum, False)
        lmTtlMultiUse = lmTtlMultiUse + 1
        llLineNum = llLineNum + 1
        lmFileMultiUseCount = lmFileMultiUseCount + 1
        rst.MoveNext
    Wend
    ilRet = ilRet
    mGetMissedReasons = True
    rst.Close
    Exit Function
    
Err_Handler:
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in mGetMissedReasons: "
            gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mGetMissedReasons: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mGetMissedReasons = False
    rst.Close
    Exit Function

End Function

Private Function mGetReplacementReasons() As Integer

    Dim ilRet As Integer
    Dim rst As ADODB.Recordset
    Dim slResults As String
    Dim slStr As String
    Dim llLineNum As Long
    
    On Error GoTo Err_Handler
    mGetReplacementReasons = False
    lmTtlMultiUse = 0
    Print #hmToMultiUse, "[ReplacementReasons]"
    Print #hmToMultiUse, "Code, Reason"
    SQLQuery = "select mnfCode, mnfName from MNF_Multi_Names where mnftype = 'M' and (mnfCodeStn = 'R')"
    Set rst = gSQLSelectCall(SQLQuery)
    llLineNum = 1
    While Not rst.EOF
        slStr = Trim$(rst!mnfCode) & "," & """" & Trim$(rst!mnfName) & """"
        Print #hmToMultiUse, gRemoveIllegalCharsAndLog(slStr, smToMultiUse, llLineNum, False)
        lmTtlMultiUse = lmTtlMultiUse + 1
        lmFileMultiUseCount = lmFileMultiUseCount + 1
        llLineNum = llLineNum + 1
        rst.MoveNext
    Wend
    ilRet = ilRet
    mGetReplacementReasons = True
    rst.Close
    Exit Function
    
Err_Handler:
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in mGetReplacementReasons: "
            gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mGetReplacementReasons: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mGetReplacementReasons = False
    rst.Close
    Exit Function

End Function

Private Sub mLogTimingResults()

    Dim llTime As Long
    
    lgETime25 = timeGetTime
    lgTtlTime25 = lgTtlTime25 + lgETime25 - lgSTime25
    gLogMsg "*** mInitiateExport - mExport time  = " & gTimeString(lgTtlTime24 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** cmdExport click = " & gTimeString(lgTtlTime24 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** mExportSpots = " & gTimeString(lgTtlTime23 / 1000, True), "WebExpSummary.Txt", False
    llTime = lgTtlTime4 + lgTtlTime2 + lgTtlTime5
    llTime = lgTtlTime23 - llTime
    gLogMsg "     mExportSpots without ggAstInfo, BuildDetail Recs or buildHeaders  = " & gTimeString(llTime / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    Header Recs = " & gTimeString(lgTtlTime5 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    BuildDetail Recs = " & gTimeString(lgTtlTime4 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       mEstimateDateAndTime " & lgCount11, "WebExpSummary.Txt", False
    gLogMsg "       ObtainRotEndDete = " & gTimeString(lgTtlTime15 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Number of Calls to ObtainRotEndDete = " & CStr(lgCount8), "WebExpSummary.Txt", False
    gLogMsg "       Format slRotCpyEndDate = " & gTimeString(lgTtlTime19 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       GetCFS = " & gTimeString(lgTtlTime16 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Build Game Inf = " & gTimeString(lgTtlTime17 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Print string to File = " & gTimeString(lgTtlTime18 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    GGetAstInfo Time = " & gTimeString(lgTtlTime2 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Regional Copy Call = " & gTimeString(lgTtlTime8 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           RC was called " & lgCount1 & " times out of a possible " & lgCount2 & " times.", "WebExpSummary.Txt", False
    gLogMsg "           gSeparateRegions+gRegionTestDefinition = " & gTimeString(lgTtlTime14 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           Binary Search " & gTimeString(lgTtlTime20 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           Get Copy " & gTimeString(lgTtlTime21 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "               Get Copy Found in Array " & lgCount9, "WebExpSummary.Txt", False
    gLogMsg "               Get Copy NOT Found in Array " & lgCount10, "WebExpSummary.Txt", False
    gLogMsg "               SQL in Get Copy " & gTimeString(lgTtlTime22 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** FTP Time = " & gTimeString(lgTtlTime3 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** Web Import Processing Time = " & gTimeString(lgTtlTime6 / 1000, True), "WebExpSummary.Txt", False
    
    lgETime1 = timeGetTime
    lgTtlTime1 = lgETime1 - lgSTime1
    gLogMsg "Total Export Time = " & gTimeString(lgTtlTime1 / 1000, True), "WebExpSummary.Txt", False
End Sub
'ttp 5333
Private Function mIsRemoveForIDC() As Boolean
    Dim blRet As Boolean
    Dim rst As ADODB.Recordset
    
    blRet = False
    If udcCriteria.CRemoveISCI = vbChecked And tgCPPosting(0).lAttCode > 0 Then
        SQLQuery = "select count(*) as mycount  from att where attcode = " & tgCPPosting(0).lAttCode & " AND RTrim(attIDCReceiverID) <> ''"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If rst.Fields("mycount").Value > 0 Then
                blRet = True
            End If
        End If
    End If
    If Not rst Is Nothing Then
        If (rst.State And adStateOpen) <> 0 Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    mIsRemoveForIDC = blRet
End Function

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mSetLogPgmSplitColumns
    If bmInStationFill Then
        gSetMousePointer grdVeh, grdVeh, vbHourglass
    End If
End Sub

Private Sub tmcFilterDelay_Timer()
    tmcFilterDelay.Enabled = False
'    If chkAllStation.Value <> vbChecked Then
        mFillStations
'    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmWebExportSchdSpot
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
        For ilLoop = 1 To grdVeh.Rows - 1 Step 1
            If Trim(grdVeh.TextMatrix(ilLoop, VEHINDEX)) <> "" Then
                If grdVeh.TextMatrix(ilLoop, SELECTEDINDEX) = "1" Then
                    ilVefCode(UBound(ilVefCode)) = grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
            End If
        Next ilLoop
        'D.S. 09/16/14 added test not to do stations if all stations were seleted
        If chkAllStation.Value = vbUnchecked Then
            For ilLoop = 0 To lbcStation.ListCount - 1
                If lbcStation.Selected(ilLoop) Then
                    ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                    ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
                End If
            Next ilLoop
        End If
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("A", "Counterpoint Affidavit", "3", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub


Private Function mSetVefExpDate(sFileName As String) As Boolean

    'Written by: Doug Smith 07/17/12
    Dim ilRet As Integer
    Dim ilLen As Integer
    Dim slFileName As String
    Dim slFromFile As Integer
    Dim hlFrom As Integer
    Dim slTemp As String
    Dim ilVefCode As Integer
    '**** Important Note: If you change the below fields you must change all other occurances of it.   ****
    '**** Currently this includes gBuildWebHeaders, gBuildwebHeaderDetail and mSetVefDate in 3 places. ****
    'FYM 02/06/19 added Market and Rank
    Dim attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime, Mode, LogStartDate, LogEndDate, MonthlyPosting, InterfaceType, UseActual, SuppressLog As String, PledgeByEvent As String, altVehname As String, MGsOnWeb As String, ReplacementsOnWeb As String, WebSiteVersion As String, Market As String, Rank As Integer, ShowCart As String
    
    
    mSetVefExpDate = False
    
    'Build the file name
    ilRet = ilRet
    slFileName = Trim$(sFileName)
    ilLen = Len(slFileName) - 8
    slFileName = right(slFileName, ilLen)
    slFileName = "WebHeaders" & slFileName
    slFileName = smWebExports & slFileName
    
    'Open the file
    On Error GoTo ImportHeadersErr_1
    ilRet = 0
    hlFrom = FreeFile
    Open slFileName For Input Access Read As hlFrom
    If ilRet <> 0 Then
        gLogMsg "Error: frmWebExportSchdSpot-mSetVefExpDate was unable to open the file: " & sFileName, "WebExportLog.Txt", False
        Exit Function
    End If

    On Error GoTo ImportHeadersErr_2
    ilRet = 0
    ' Skip past the header definition record.
    '**** Important Note: If you change the below fields you must change all other occurances of it.   ****
    '**** Currently this includes gBuildWebHeaders, gBuildwebHeaderDetail and mSetVefDate in 3 places. ****
    'FYM 02/06/19 added Market and Rank
    Input #hlFrom, attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime, Mode, LogStartDate, LogEndDate, MonthlyPosting, InterfaceType, UseActual, SuppressLog, PledgeByEvent, altVehname, MGsOnWeb, ReplacementsOnWeb, WebSiteVersion, Market, Rank, ShowCart
    'provide a basic sanity check
    If Len(attCode) < 1 Or attCode <> "attcode" Then
        Exit Function
    End If
    
    'Loop through the file and update vehicles last export date
    slTemp = ""
    Do While Not EOF(hlFrom)
        '**** Important Note: If you change the below fields you must change all other occurances of it.   ****
        '**** Currently this includes gBuildWebHeaders, gBuildwebHeaderDetail and mSetVefDate in 3 places. ****
        'FYM 02/06/19 added Market and Rank
        Input #hlFrom, attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime, Mode, LogStartDate, LogEndDate, MonthlyPosting, InterfaceType, UseActual, SuppressLog, PledgeByEvent, altVehname, MGsOnWeb, ReplacementsOnWeb, WebSiteVersion, Market, Rank, ShowCart
        If VehicleName <> slTemp Then
            slTemp = VehicleName
            ilVefCode = gGetVehCodeFromAttCode(CStr(attCode))
            ilRet = gUpdateLastExportDate(ilVefCode, smEndDate)
        End If
    Loop
    
    Close hlFrom
    mSetVefExpDate = True
    Exit Function
    
ImportHeadersErr_1:
    ilRet = 1
    Resume Next

ImportHeadersErr_2:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in mSetVefExpDate: "
            gLogMsg "Error: " & gMsg & gErrSQL.Description & " Error #" & gErrSQL.NativeError & "; Line #" & Erl, "WebImportLog.Txt", False
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mSetVefExpDate: "
        gLogMsg "Error: " & gMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "WebExportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If

    Exit Function

End Function

Private Sub mCreateJelliRecord(tlAstInfo As ASTINFO, slISCI As String, slCreative As String, slAdvt As String)
    Dim slRecord As String
    Dim ilPos As Integer
    Dim slCallLetters As String
    Dim slBand As String
    Dim slStr As String
    Dim llVef As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    ReDim ilAllowedDays(0 To 6) As Integer
    '8/1/14: Compliant not required
    'Dim ilCompliant As Integer
    Dim ilDay As Integer
    Dim ilAgf As Integer
    Dim slAgency As String
    Dim ilVff As Integer
    Dim slShowVehName As String
    
    If smJelliExport <> "Y" Then
        Exit Sub
    End If
    If smVehicleExportJelli <> "Y" Then
        Exit Sub
    End If
    If smAttExportToJelli <> "Y" Then
        Exit Sub
    End If
    If hmJelli <= 0 Then
        Exit Sub
    End If
    slRecord = """" & slISCI & """"
    'Feed date and time
    slRecord = slRecord & "," & Format(tlAstInfo.sFeedDate, "yyyy-mm-dd") & " " & Format(tlAstInfo.sFeedTime, "hh:mm:ss")
    'Spot length
    slRecord = slRecord & "," & tlAstInfo.iLen
    'Creative Title
    slRecord = slRecord & "," & """" & slCreative & """"
    'Station Call Letters
    slStr = Trim$(txtCallLetters.Text)
    ilPos = InStr(1, slStr, "-", vbBinaryCompare)
    If ilPos > 0 Then
        slCallLetters = Left(slStr, ilPos - 1)
        slBand = Mid(slStr, ilPos + 1)
    Else
        slCallLetters = slStr
        slBand = ""
    End If
    slRecord = slRecord & "," & """" & slCallLetters & """"
    slRecord = slRecord & "," & """" & slBand & """"
    'llVef = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
    'Output alternate if exist
    ilVff = gBinarySearchVff(imVefCode)
    If ilVff = -1 Then
        ilVff = gPopVff()
        ilVff = gBinarySearchVff(imVefCode)
    End If
    slShowVehName = ""
    If ilVff <> -1 Then
        slShowVehName = Trim$(tgVffInfo(ilVff).sWebName)
    End If
    If Trim$(slShowVehName) <> "" Then
        slRecord = slRecord & "," & """" & Trim$(slShowVehName) & """"
    Else
        llVef = gBinarySearchVef(CLng(imVefCode))
        If llVef = -1 Then
            llVef = gPopSellingVehicles()
            llVef = gBinarySearchVef(CLng(imVefCode))
        End If
        If llVef <> -1 Then
            slRecord = slRecord & "," & """" & Trim$(tgVehicleInfo(llVef).sVehicle) & """"
        Else
            slRecord = slRecord & "," & """" & Trim$(slShowVehName) & """"
        End If
    End If
    slRecord = slRecord & "," & imVefCode   'tlAstInfo.iVefCode
    gGetLineParameters False, tlAstInfo, slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays()
    slRecord = slRecord & "," & Format(slStartTime, "hh:mm:ss")
    slRecord = slRecord & "," & Format(slEndTime, "hh:mm:ss")
    If tlAstInfo.iLstMon = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstTue = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstWed = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstThu = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstFri = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstSat = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    If tlAstInfo.iLstSun = 0 Then
        slRecord = slRecord & "," & """" & "N" & """"
    Else
        slRecord = slRecord & "," & """" & "Y" & """"
    End If
    slRecord = slRecord & "," & tlAstInfo.lCntrNo
    slRecord = slRecord & "," & """" & slAdvt & """"
    slAgency = ""
    ilAgf = gBinarySearchAgency(CLng(tlAstInfo.iAgfCode))
    If ilAgf = -1 Then
        ilAgf = gPopAgencies()
        ilAgf = gBinarySearchAgency(CLng(tlAstInfo.iAgfCode))
    End If
    If ilAgf <> -1 Then
        slAgency = Trim$(tgAgencyInfo(ilAgf).sAgencyName)
    End If
    slRecord = slRecord & "," & """" & slAgency & """"
    slRecord = slRecord & "," & """" & tlAstInfo.lCode & """"
    'Write Record
    Print #hmJelli, slRecord
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenJelliFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenJelliFile() As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slLetter As String

    On Error GoTo mOpenJelliFileErr:
    slLetter = ""
    Do
        ilRet = 0
        smJelliFileName = sgExportDirectory & "Jelli-" & Format$(smStartDate, "mmddyy") & slLetter & ".txt"
        ilRet = gFileExist(smJelliFileName)
        If ilRet = 0 Then
            If slLetter = "" Then
                slLetter = "A"
            Else
                slLetter = Chr$(Asc(slLetter) + 1)
            End If
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenJelliFileErr:
    hmJelli = FreeFile
    Open smJelliFileName For Output As hmJelli
    If ilRet <> 0 Then
        Close hmJelli
        hmJelli = -2
        gMsgBox "Open File " & smJelliFileName & " error#" & Str$(Err.Number), vbOKOnly
        mOpenJelliFile = False
        Exit Function
    End If
    On Error GoTo 0
    mOpenJelliFile = True
    Exit Function
mOpenJelliFileErr:
    ilRet = 1
    Resume Next
End Function


Private Sub grdVeh_GotFocus()
    cmdCancel.Caption = "&Cancel"
    smGridTypeAhead = ""
    'plcTme.Visible = False
    'plcCalendar.Visible = False
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
'    ilFound = gGrid_GetRowCol(grdVeh, X, Y, llCurrentRow, llCol)
    'D.S. 07-28-17
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

Private Sub grdVeh_Scroll()
    cmdCancel.Caption = "&Cancel"
    lmScrollTop = grdVeh.TopRow
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

Private Sub grdVeh_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
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
End Sub
Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    imBypassAll = True
    chkAll.Value = vbUnchecked
    imBypassAll = False
    gGrid_Clear grdVeh, True
    For llRow = grdVeh.FixedRows To grdVeh.Rows - 1 Step 1
        grdVeh.Row = llRow
        For llCol = 0 To VEHCODEINDEX Step 1
            grdVeh.Col = llCol
            grdVeh.CellBackColor = vbWhite
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
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mClearGrid"
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdVeh.Row = llRow
    For llCol = VEHINDEX To VEHINDEX Step 1
        grdVeh.Col = llCol
        If grdVeh.TextMatrix(llRow, LOGSORTINDEX) = "A" And (llCol = LOGINDEX) Then
            If grdVeh.CellBackColor <> vbRed Then
                grdVeh.CellBackColor = vbRed
                grdVeh.CellForeColor = vbRed
            End If
        ElseIf grdVeh.TextMatrix(llRow, PGMSORTINDEX) = "A" And (llCol = PGMINDEX) Then
            If grdVeh.CellBackColor <> vbRed Then
                grdVeh.CellBackColor = vbRed
                grdVeh.CellForeColor = vbRed
            End If
        ElseIf grdVeh.TextMatrix(llRow, SPLITSORTINDEX) = "A" And (llCol = SPLITINDEX) Then
            If grdVeh.CellBackColor <> vbRed Then
                grdVeh.CellBackColor = vbRed
                grdVeh.CellForeColor = vbRed
            End If
        Else
            If grdVeh.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
                If grdVeh.CellBackColor <> vbWhite Then
                    grdVeh.CellBackColor = vbWhite
                    grdVeh.CellForeColor = vbWindowText
                End If
            Else
                If grdVeh.CellBackColor <> vbHighlight Then
                    grdVeh.CellBackColor = vbHighlight
                    grdVeh.CellForeColor = vbWhite
                End If
            End If
        End If
    Next llCol
End Sub

Private Sub mTerminate()
    
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gSetMousePointer grdVeh, grdVeh, vbDefault
    igManUnload = Yes
    Unload frmWebExportSchdSpot
    igManUnload = No
End Sub

Private Sub mAddToGrid(llRow As Long, llVeh As Long)

    Dim llCol As Long
    
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
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mAddToGrid"
End Sub
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
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mFindDuplVeh"
End Function
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
                'D.S. 11/8/19 added exit for. we only need to know if at least one exist
                Exit For
            End If
        End If
    Next llRow
    mGetGrdSelCount = llCount
    Exit Function
ErrHand:
    gSetMousePointer grdVeh, grdVeh, vbDefault
    gHandleError "AffErrorLog.txt", "frmWebExportSchdSpot-mGetGrdSelCount"
End Function

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
    Dim iCount As Integer
    Dim blHasAlert As Boolean
    Dim slDates As String
    
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
    slDates = ""
    'TTP 10243 - 7/15/21 - JW - Build Comma separated list of all Monday dates in Start - End Range
    Do
        If slDates <> "" Then slDates = slDates & ","
        slDates = slDates & "'" & Format$(slMoWeekDate, sgSQLDateForm) & "'"
        slMoWeekDate = DateAdd("d", 7, slMoWeekDate)
    Loop While DateValue(gAdjYear(slMoWeekDate)) < DateValue(gAdjYear(smEndDate))
    slMoWeekDate = gObtainPrevMonday(edcDate.Text)
    
    
    'TTP 10243 - 7/15/21 - JW - Get Array of Vehicles that have RePrint Alerts for all Vehicles for Each Week in Comma Separated list of Monday Dates in Range
    mGetVefReprintAlerts slDates
    
    'TTP 10243 - 7/15/21 - JW - Get Array of Vehicles that have Required Affiliate Alerts
    mGetVefRequiredAlerts smEndDate
    
'TTP 10243 - 7/15/21 - JW - Dont Loop for each Week, Get that data from mGetVefReprintAlerts() above....
'    Do
'Debug.Print "Loop Start"
        For ilLoop = grdVeh.FixedRows To grdVeh.Rows - 1
            If Trim(grdVeh.TextMatrix(ilLoop, VEHCODEINDEX)) <> "" Then
                ilVehCode = Trim(grdVeh.TextMatrix(ilLoop, VEHCODEINDEX))
                smVefName = grdVeh.TextMatrix(ilLoop, VEHINDEX)
                blNeedsLogGened = False
                
'TTP 10243 - 7/15/21 - JW - get RePrint Alerts for Vehicle from Array
                '*** Check for log alerts ***
                blHasAlert = mCheckVefOtherAlerts(ilVehCode)
                If Not blHasAlert Then
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
                        'If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "1" Then
                        If .CellBackColor <> vbWhite Then
                            .CellBackColor = vbWhite
                            .CellForeColor = vbWhite
                        End If
                        'Else
                        '.CellBackColor = vbHighlight
                        '.CellForeColor = vbHighlight
                        'End If
                    End With
                End If

'TTP 10243 - 7/15/21 - JW - Dont Query in Loops if possible
'                '*** Check for log alerts ***
'                SQLQuery = "Select * from AUF_Alert_User where "
'                SQLQuery = SQLQuery & "aufStatus = 'R' "
'                SQLQuery = SQLQuery & "and aufType = 'L' "
'                SQLQuery = SQLQuery & "and aufSubType <> 'M' "
'                SQLQuery = SQLQuery & "and aufSubType <> '' "
'                SQLQuery = SQLQuery & "and aufMoWeekDate = " & "'" & Format$(slMoWeekDate, sgSQLDateForm) & "' "
'                SQLQuery = SQLQuery & "and aufVefCode = " & ilVehCode
'                Set rst = gSQLSelectCall(SQLQuery)
'                blRstSetUsed = True
'                If rst.EOF Then
'                    'No alerts found; now check to see if the log for the given week has been generated
'                    'If ilVehCode = 3 Then
'                    '    blNeedsLogGened = blNeedsLogGened
'                    'End If
'                    blNeedsLogGened = mLogNeedsToBeGenerated(ilVehCode, slStartDate, slEndDate)
'                Else
'                    blNeedsLogGened = True
'                End If
'                If blNeedsLogGened Then
'                    grdVeh.TextMatrix(0, LOGINDEX) = "Gen *"
'                    With grdVeh
'                        .Row = ilLoop
'                        .Col = LOGINDEX
'                        '.CellFontName = "Monotype Sorts"
'                        '.TextMatrix(ilLoop, LOGINDEX) = ""
'                        .TextMatrix(ilLoop, LOGSORTINDEX) = "A"
'                        If .CellBackColor <> vbRed Then
'                            .CellBackColor = vbRed
'                            .CellForeColor = vbBlack
'                        End If
'                    End With
'                    slLogNeeded = "Y"
'                Else
'                    With grdVeh
'                        .Row = ilLoop
'                        .Col = LOGINDEX
'                        '.CellFontName = "Monotype Sorts"
'                        '.TextMatrix(ilLoop, LOGINDEX) = ""
'                        .TextMatrix(ilLoop, LOGSORTINDEX) = "B"
'                        'If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "1" Then
'                        If .CellBackColor <> vbWhite Then
'                            .CellBackColor = vbWhite
'                            .CellForeColor = vbBlack
'                        End If
'                        'Else
'                        '.CellBackColor = vbHighlight
'                        '.CellForeColor = vbHighlight
'                        'End If
'                    End With
'                End If

'TTP 10243 - 7/15/21 - JW - get Alerts for Vehicle from Array
                '*** Check for program change alerts ***
                blHasAlert = mCheckVefRequiredAlerts(ilVehCode)
                If blHasAlert Then
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
                        'If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "0" Then
                            If .CellBackColor <> vbWhite Then
                                .CellBackColor = vbWhite
                                .CellForeColor = vbWhite
                            End If
                        'End If
                    End With
                End If

'TTP 10243 - 7/15/21 - JW - Dont Query in Loops if possible
'                '*** Check for program change alerts ***
'                SQLQuery = "Select * from AUF_Alert_User where "
'                SQLQuery = SQLQuery & "aufStatus = 'R' "
'                SQLQuery = SQLQuery & "and aufType = 'P' "
'                '2/15/18: A= Agreement changed
'                'SQLQuery = SQLQuery & "and aufSubType <> '' "
'                SQLQuery = SQLQuery & "and aufSubType = 'A' "
'                SQLQuery = SQLQuery & "and aufMoWeekDate <= " & "'" & Format$(smEndDate, sgSQLDateForm) & "'"
'                SQLQuery = SQLQuery & "and aufVefCode = " & ilVehCode
''    Debug.Print "mFindAlertsForGrdVeh: " & SQLQuery
'                Set rst = gSQLSelectCall(SQLQuery)
'                If Not rst.EOF Then
'                    grdVeh.TextMatrix(0, PGMINDEX) = "Pgm *"
'                    With grdVeh
'                        .Row = ilLoop
'                        .Col = PGMINDEX
'                        '.CellFontName = "Monotype Sorts"
'                        '.TextMatrix(ilLoop, PGMINDEX) = ""
'                        .TextMatrix(ilLoop, PGMSORTINDEX) = "A"
'                        If .CellBackColor <> vbRed Then
'                            .CellBackColor = vbRed 'Red
'                            .CellForeColor = vbBlack 'Red
'                        End If
'                    End With
'                    slPGMNeeded = "Y"
'                Else
'                   With grdVeh
'                        .Row = ilLoop
'                        .Col = PGMINDEX
'                        '.CellFontName = "Monotype Sorts"
'                        '.TextMatrix(ilLoop, PGMINDEX) = ""
'                        .TextMatrix(ilLoop, PGMSORTINDEX) = "B"
'                        If .TextMatrix(grdVeh.Row, SELECTEDINDEX) = "0" Then
'                            If .CellBackColor <> vbWhite Then
'                                .CellBackColor = vbWhite
'                                .CellForeColor = vbBlack
'                            End If
'                        End If
'                    End With
'                End If


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
                                .CellForeColor = vbBlack 'Red
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
                                .CellForeColor = vbBlack
                            End If
                        End If
                    End With
                End If
            End If
        End If

        Next
        
'        slMoWeekDate = DateAdd("d", 7, slMoWeekDate)
'    Loop While DateValue(gAdjYear(slMoWeekDate)) < DateValue(gAdjYear(smEndDate))

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
    llVpf = gBinarySearchVpf(CLng(iVefCode))
    If llVpf = -1 Then
        llVpf = gPopVehicleOptions()
        llVpf = gBinarySearchVpf(CLng(iVefCode))
    End If
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
'    lbcStation.Clear
'    If mGetGrdSelCount() = 1 Then
'        edcTitle3.Visible = True
'        chkAllStation.Visible = True
        lbcStation.Visible = True
'        mFillStations
'    Else
'        edcTitle3.Visible = False
'        chkAllStation.Visible = False
'        lbcStation.Visible = False
'    End If
'    imBypassAll = True
'    chkAll.Value = vbUnchecked
'    imBypassAll = False
    mFillStations
End Sub

Private Sub mSetLogPgmSplitColumns()
    'D.S. 12/27/17
    If imTerminate Then
        Exit Sub
    End If
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
                    If rbcFilter(4).Value = True Then
                        If chkAllStation.Value = vbUnchecked Then
                            If lbcStation.SelCount > 0 Then
                                ilEnable = True
                            End If
                        Else
                            ilEnable = True
                        End If
                    Else
                        If lbcFilter.SelCount > 0 Then
                            If chkAllStation.Value = vbUnchecked Then
                                If lbcStation.SelCount > 0 Then
                                    ilEnable = True
                                End If
                            Else
                                ilEnable = True
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
        Next llRow
    End If
    
    cmdExport.Enabled = ilEnable
End Sub


Private Sub mPopDMA()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgMarketInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgMarketInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Sub mPopMSA()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgMSAMarketInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgMSAMarketInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Function mGetMSA(lMSACode As Long) As String

    Dim ilLoop As Integer
    
    mGetMSA = ""
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        If tgMSAMarketInfo(ilLoop).lCode = lMSACode Then
            mGetMSA = tgMSAMarketInfo(ilLoop).sName
            Exit For
        End If
    Next ilLoop

End Function

Private Sub mPopFormat()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgFormatInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgFormatInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Function mGetFormat(mFormatCode As Long) As String

    Dim ilLoop As Integer
    
    mGetFormat = ""
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        If mFormatCode = tgFormatInfo(ilLoop).lCode Then
            mGetFormat = tgFormatInfo(ilLoop).sName
        End If
    Next ilLoop

End Function


Private Sub mPopState()
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    lbcFilter.Clear
    ilRet = gPopStates()
    For ilRow = 0 To UBound(tgStateInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgStateInfo(ilRow).sPostalName) & " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        lbcFilter.ItemData(lbcFilter.NewIndex) = ilRow    'tgStateInfo(ilRow).iCode
    Next ilRow
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mPopState"
End Sub


Private Function mGetMarketName(iMarketCode As Integer) As String

    Dim temp_rst As ADODB.Recordset
    
    mGetMarketName = ""
    SQLQuery = "Select mktName, mktRank from Mkt where mktCode = " & rst!shttMktCode
    Set temp_rst = gSQLSelectCall(SQLQuery)
    If Not temp_rst.EOF Then
        mGetMarketName = Trim$(temp_rst!mktName)
    End If

End Function

Public Function mIsServiceAgreement(lAttCode As Long) As Boolean

    Dim rstATT As ADODB.Recordset
    Dim SQLQuery1 As String
    
    mIsServiceAgreement = False
    SQLQuery1 = "Select attserviceagreement from att "
    SQLQuery1 = SQLQuery1 + " WHERE (attCode = " & lAttCode & ")"
    Set rstATT = gSQLSelectCall(SQLQuery1)
    If Not rstATT.EOF Then
        If rstATT!attServiceAgreement = "Y" Then
            mIsServiceAgreement = True
        End If
    End If
    rstATT.Close
    
End Function
Public Sub mStartVendorProgress()
    frWebVendor.Visible = True
    PbcWebVendor.Value = 0
End Sub
Public Sub mUpdateVendorProgress(ilProgress As Integer)
    PbcWebVendor.Value = ilProgress
End Sub
Public Sub mStopVendorProgress()
    frWebVendor.Visible = False
End Sub
'10000
Private Sub mSetWebVendorsToTest()
    Dim slSql As String
    slSql = "update vendorservicecontroller set GenerateFile = 'Y'  where mode = 'E'"
    If gExecWebSQLWithRowsEffected(slSql) <> -1 Then
        slSql = "update vendorservicecontroller set ImportFiles = 'D'  where mode = 'I'"
        gExecWebSQLWithRowsEffected slSql
    End If

End Sub

Sub mGetVefRequiredAlerts(smEndDate As String)
    ReDim tmVefReqAlerts(0) As AUF
    Dim rstAUF As ADODB.Recordset
    Dim SQLQuery1 As String
    
    SQLQuery1 = "Select DISTINCT aufVefCode from AUF_Alert_User where aufStatus = 'R' and aufType = 'P' and aufSubType = 'A' and aufMoWeekDate <= '" & Format$(smEndDate, sgSQLDateForm) & "'"
    Set rstAUF = gSQLSelectCall(SQLQuery1)
    If Not rstAUF.EOF Then
        Do
            tmVefReqAlerts(UBound(tmVefReqAlerts)).iVefCode = rstAUF.Fields(0).Value
            ReDim Preserve tmVefReqAlerts(0 To UBound(tmVefReqAlerts) + 1)
            rstAUF.MoveNext
        Loop While Not rstAUF.EOF
    End If
    rstAUF.Close
End Sub

Function mCheckVefRequiredAlerts(ilVehCode As Integer) As Boolean
    mCheckVefRequiredAlerts = False
    Dim ilLoop As Integer
    For ilLoop = 0 To UBound(tmVefReqAlerts)
        If tmVefReqAlerts(ilLoop).iVefCode = ilVehCode Then
            mCheckVefRequiredAlerts = True
            Exit For
        End If
    Next ilLoop
End Function

Sub mGetVefReprintAlerts(smDates As String)
    ReDim tmVefOtherAlerts(0) As AUF
    Dim rstAUF As ADODB.Recordset
    Dim SQLQuery1 As String
    
    SQLQuery1 = "Select DISTINCT aufVefCode from AUF_Alert_User where aufStatus = 'R' and aufType = 'L' and aufSubType <> 'M' and aufSubType <> '' and aufMoWeekDate in (" & smDates & ")"
    Set rstAUF = gSQLSelectCall(SQLQuery1)
    If Not rstAUF.EOF Then
        Do
            tmVefOtherAlerts(UBound(tmVefOtherAlerts)).iVefCode = rstAUF.Fields(0).Value
            ReDim Preserve tmVefOtherAlerts(0 To UBound(tmVefOtherAlerts) + 1)
            rstAUF.MoveNext
        Loop While Not rstAUF.EOF
    End If
    rstAUF.Close
End Sub

Function mCheckVefOtherAlerts(ilVehCode As Integer) As Boolean
    mCheckVefOtherAlerts = False
    Dim ilLoop As Integer
    For ilLoop = 0 To UBound(tmVefOtherAlerts)
        If tmVefOtherAlerts(ilLoop).iVefCode = ilVehCode Then
            mCheckVefOtherAlerts = True
            Exit For
        End If
    Next ilLoop
End Function

