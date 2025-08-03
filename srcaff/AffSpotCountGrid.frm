VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSpotCountGrid 
   Caption         =   "Spot Count Tie-out"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9360
   Begin VB.TextBox edcGridInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1755
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3195
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "AffSpotCountGrid.frx":0000
      Top             =   210
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Timer tmcFillGrid 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6705
      Top             =   5010
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
      Left            =   7305
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   5070
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4110
      TabIndex        =   0
      Top             =   4890
      Width           =   1245
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7860
      Top             =   4935
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5310
      FormDesignWidth =   9360
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCounts 
      Height          =   4500
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7938
      _Version        =   393216
      Rows            =   4
      Cols            =   27
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   27
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacNote 
      Caption         =   "Multicast Column: * = Station; ** = Master Station"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   5040
      Width           =   3690
   End
   Begin VB.Image imcExport 
      Height          =   480
      Left            =   8535
      Picture         =   "AffSpotCountGrid.frx":0004
      Top             =   4845
      Width           =   480
   End
   Begin VB.Label lacKey 
      Caption         =   "Click on the second Title row to see Column definition.  Click within Cell to see discrepancy information."
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   210
      TabIndex        =   4
      Top             =   0
      Width           =   7425
   End
   Begin VB.Image imcKey 
      Height          =   225
      Left            =   5805
      Picture         =   "AffSpotCountGrid.frx":08CE
      Top             =   5025
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmSpotCountGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Specification vales
Private smFeedStartDate As String
Private smFeedEndDate As String
Private smSQLFeedStartDate As String
Private smSQLFeedEndDate As String
Private bmInBalance As Boolean
Private bmOutBalance As Boolean
Private bmNotCompliant As Boolean
Private bmPartiallyPosted As Boolean
Private bmUserPosting As Boolean
Private bmBreakOutBalance As Boolean
Private imSort As Integer   '0=Major vehicle; 1= Major is station
Private bmIncludeWeb As Boolean
Private bmIncludeCodes As Boolean

'Grid Controls
Private lmHDLastCol As Long
Private lmCellLastCol As Long
Private lmCellLastRow As Long

Private lmTopRow As Long            'Top row when cell clicked or - 1

Private imLastColSorted As Integer
Private imLastSort As Integer

Private lmTotalOkCount As Long
Private lmTotalErrorCount As Long

Private bmTerminate As Boolean

Private tmDat() As DATRST

Private rst_Ast As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Lst As ADODB.Recordset
Private rst_DAT As ADODB.Recordset
Private rst_ent As ADODB.Recordset
Private rst_vat As ADODB.Recordset
Private rst_vef As ADODB.Recordset
Private rst_lcf As ADODB.Recordset

Private lmMajor() As Long
Private lmMinor() As Long
Private imVefCode As Integer
Private imShttCode As Integer
Private lmAttCode As Long

Private smSQLQuery As String
Private smSQLMoDate As String   'Used to get MG info
Private smSQLSuDate As String   'Used to get MG info
Private lmWkStartDate As Long   'start date within week to get spots
Private lmWkEndDate As Long     'end date within week to get spots
Private lmStartDate As Long 'Range start date
Private lmEndDate As Long   'Range end date
Private smSQLStartDate As String
Private smSQLEndDate As String

Private lmRowColor As Long

Private bmRetrieveLst As Boolean ' true=retrieve lst values; false = values stored in fields
'Private smAirLstCount As String
'Private smMissedLstCount As String
'Private smMGLstCount As String
Private smNetworkCount As String
Private smNetworkBreakCount As String
Private smFeedBreakCount As String
Private bmAirPlayConflict As Boolean

Private imVendorID() As Integer


Private Const DATEINDEX = 0
Private Const VEHICLEINDEX = 1
Private Const STATIONINDEX = 2
Private Const MULTICASTINDEX = 3
Private Const POSTMETHODINDEX = 4
Private Const FLOWINDEX = 5
Private Const AIRPLAYSINDEX = 6
Private Const NETWORKBREAKINDEX = 7
Private Const FEEDBREAKINDEX = 8
Private Const NETWORKINDEX = 9
Private Const FEEDSPOTINDEX = 10
Private Const FEEDNCINDEX = 11
Private Const PLEDGESPOTINDEX = 12
Private Const PLEDGENCINDEX = 13
Private Const SPOTINDEX = 14
Private Const NOTCARRIEDINDEX = 15
Private Const POSTBYINDEX = 16
Private Const MGINDEX = 17
Private Const VENDOREXPORTINDEX = 18
Private Const VENDORIMPORTINDEX = 19
Private Const VENDORAPPLIEDINDEX = 20
Private Const AGYCOMPLIANTINDEX = 21
Private Const ADVTCOMPLIANTINDEX = 22
Private Const ATTSEQNOINDEX = 23
Private Const MULTICASTSTATIONINDEX = 24
Private Const CODESINDEX = 25
Private Const SORTINDEX = 26

'Private Const HDDATEINDEX = 0
'Private Const HDVEHICLEINDEX = 1
'Private Const HDSTATIONINDEX = 2
'Private Const HDFLOWINDEX = 3
'Private Const HDNETWORKINDEX = 4
'Private Const HDAIRPLAYSINDEX = 5
'Private Const HDAGREEMENTINDEX = 6
'Private Const HDSPOTINDEX = 7
'Private Const HDNOTCARRIEDINDEX = 8
'Private Const HDMGINDEX = 9
'Private Const HDVENDORINDEX = 10
'Private Const HDCOMPLIANTINDEX = 11

Private Sub cmcDone_Click()
    If (cmcDone.Caption = "Done") Or (cmcDone.Caption = "Close") Then
        Unload frmSpotCountGrid
    Else
        bmTerminate = True
    End If
End Sub



Private Sub cmcDone_GotFocus()
    mClearSelection
End Sub

Private Sub edcGridInfo_Click()
    mClearSelection
End Sub

Private Sub Form_Click()
    mClearSelection
    'lacKey.Visible = False
    cmcDone.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 1.1
    Me.Height = (Screen.Height) / 1.2
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    'gCenterForm Me
End Sub

Private Sub Form_Load()
    Dim ilVef As Integer
    Dim ilShtt As Integer
    
    gSetMousePointer grdCounts, grdCounts, vbHourglass
    
    bmTerminate = False
    imcKey.Picture = frmDirectory!imcKey.Picture
    imcExport.Picture = frmDirectory!imcExport.Picture
    mLoadKeyText
    smFeedStartDate = frmSpotCountSpec.edcFeedStartDate.Text
    smSQLFeedStartDate = Format(smFeedStartDate, sgSQLDateForm)
    smFeedEndDate = frmSpotCountSpec.edcFeedEndDate.Text
    smSQLFeedEndDate = Format(smFeedEndDate, sgSQLDateForm)
    If frmSpotCountSpec.ckcInBalance = vbChecked Then
        bmInBalance = True
    Else
        bmInBalance = False
    End If
    If frmSpotCountSpec.ckcOutBalance = vbChecked Then
        bmOutBalance = True
    Else
        bmOutBalance = False
    End If
    If frmSpotCountSpec.ckcNotCompliant = vbChecked Then
        bmNotCompliant = True
    Else
        bmNotCompliant = False
    End If
    If frmSpotCountSpec.ckcPartiallyPosted = vbChecked Then
        bmPartiallyPosted = True
    Else
        bmPartiallyPosted = False
    End If
    
    If frmSpotCountSpec.ckcUserPosting = vbChecked Then
        bmUserPosting = True
    Else
        bmUserPosting = False
    End If
    
    If frmSpotCountSpec.ckcBreakOutBalance = vbChecked Then
        bmBreakOutBalance = True
    Else
        bmBreakOutBalance = False
    End If
    
    If frmSpotCountSpec.ckcWeb = vbChecked Then
        bmIncludeWeb = True
    Else
        bmIncludeWeb = False
    End If
    If frmSpotCountSpec.ckcCodeRow = vbChecked Then
        bmIncludeCodes = True
    Else
        bmIncludeCodes = False
    End If
    
    If frmSpotCountSpec.rbcSort(0).Value Then
        imSort = 0
    Else
        imSort = 1
    End If
    ReDim lmMajor(0 To 0) As Long
    ReDim lmMinor(0 To 0) As Long
    
    For ilVef = 0 To frmSpotCountSpec.lbcVehicles.ListCount - 1 Step 1
        If frmSpotCountSpec.lbcVehicles.Selected(ilVef) Then
            If imSort = 0 Then
                lmMajor(UBound(lmMajor)) = frmSpotCountSpec.lbcVehicles.ItemData(ilVef)
                ReDim Preserve lmMajor(0 To UBound(lmMajor) + 1) As Long
            Else
                lmMinor(UBound(lmMinor)) = frmSpotCountSpec.lbcVehicles.ItemData(ilVef)
                ReDim Preserve lmMinor(0 To UBound(lmMinor) + 1) As Long
            End If
        End If
    Next ilVef
    For ilShtt = 0 To frmSpotCountSpec.lbcStations.ListCount - 1 Step 1
        If frmSpotCountSpec.lbcStations.Selected(ilShtt) Then
            If imSort = 1 Then
                lmMajor(UBound(lmMajor)) = frmSpotCountSpec.lbcStations.ItemData(ilShtt)
                ReDim Preserve lmMajor(0 To UBound(lmMajor) + 1) As Long
            Else
                lmMinor(UBound(lmMinor)) = frmSpotCountSpec.lbcStations.ItemData(ilShtt)
                ReDim Preserve lmMinor(0 To UBound(lmMinor) + 1) As Long
            End If
        End If
    Next ilShtt

    gPopShttInfo

    lmHDLastCol = -1
    lmCellLastCol = -1
    lmCellLastRow = -1
    
    imLastColSorted = -1
    imLastSort = -1
    tmcFillGrid.Enabled = True
    'gSetMousePointer grdCounts, grdCounts, vbDefault

End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    gSetMousePointer grdCounts, grdCounts, vbHourglass
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdCounts
    gGrid_FillWithRows grdCounts
    'imcKey.Left = grdCounts.Left
    'imcKey.Top = 0
    txtKey.Left = grdCounts.Left
    txtKey.Top = grdCounts.Top + 2 * grdCounts.RowHeight(0)
    lacKey.Top = 0 'imcKey.Top
    lacKey.Left = grdCounts.Left    'imcKey.Left + imcKey.Width / 2
    imcExport.Top = 0
    imcExport.Left = Me.Width - imcExport.Width
    'mPopulate
    gSetMousePointer grdCounts, grdCounts, vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Ast.Close
    rst_Cptt.Close
    rst_att.Close
    rst_Lst.Close
    rst_DAT.Close
    rst_ent.Close
    rst_vat.Close
    rst_vef.Close
    rst_lcf.Close
    
    Erase imVendorID
    Set frmSpotCountGrid = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdCounts.ColWidth(ATTSEQNOINDEX) = 0
    grdCounts.ColWidth(MULTICASTSTATIONINDEX) = 0
    grdCounts.ColWidth(CODESINDEX) = 0
    grdCounts.ColWidth(SORTINDEX) = 0
    grdCounts.ColWidth(DATEINDEX) = grdCounts.Width * 0.06
    grdCounts.ColWidth(STATIONINDEX) = grdCounts.Width * 0.06
    grdCounts.ColWidth(MULTICASTINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(POSTMETHODINDEX) = grdCounts.Width * 0.05
    grdCounts.ColWidth(FLOWINDEX) = grdCounts.Width * 0.06
    grdCounts.ColWidth(AIRPLAYSINDEX) = grdCounts.Width * 0.035
    grdCounts.ColWidth(NETWORKINDEX) = grdCounts.Width * 0.035
    grdCounts.ColWidth(NETWORKBREAKINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(FEEDBREAKINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(FEEDSPOTINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(FEEDNCINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(PLEDGESPOTINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(PLEDGENCINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(SPOTINDEX) = grdCounts.Width * 0.035
    grdCounts.ColWidth(NOTCARRIEDINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(POSTBYINDEX) = grdCounts.Width * 0.02
    grdCounts.ColWidth(MGINDEX) = grdCounts.Width * 0.035
    grdCounts.ColWidth(VENDOREXPORTINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(VENDORIMPORTINDEX) = grdCounts.Width * 0.03
    grdCounts.ColWidth(VENDORAPPLIEDINDEX) = grdCounts.Width * 0.04
    grdCounts.ColWidth(AGYCOMPLIANTINDEX) = grdCounts.Width * 0.04
    grdCounts.ColWidth(ADVTCOMPLIANTINDEX) = grdCounts.Width * 0.04
    grdCounts.ColWidth(VEHICLEINDEX) = grdCounts.Width - GRIDSCROLLWIDTH - 15
    For ilCol = DATEINDEX To SORTINDEX Step 1
        If ilCol <> VEHICLEINDEX Then
            grdCounts.ColWidth(VEHICLEINDEX) = grdCounts.ColWidth(VEHICLEINDEX) - grdCounts.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdCounts
    grdCounts.Row = 0
    grdCounts.Col = NETWORKBREAKINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
    grdCounts.Row = 0
    grdCounts.Col = FEEDSPOTINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
    grdCounts.Row = 0
    grdCounts.Col = PLEDGESPOTINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
    grdCounts.Row = 0
    grdCounts.Col = SPOTINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
    grdCounts.Row = 0
    grdCounts.Col = VENDOREXPORTINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
    grdCounts.Row = 0
    grdCounts.Col = AGYCOMPLIANTINDEX
    grdCounts.CellAlignment = flexAlignCenterTop
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdCounts.TextMatrix(0, DATEINDEX) = "Date"
    grdCounts.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdCounts.TextMatrix(0, STATIONINDEX) = "Station"
    grdCounts.TextMatrix(0, MULTICASTINDEX) = "Multi-"
    grdCounts.TextMatrix(0, POSTMETHODINDEX) = "Post"
    grdCounts.TextMatrix(0, FLOWINDEX) = "Flow"
    grdCounts.TextMatrix(0, AIRPLAYSINDEX) = "Air"
    grdCounts.TextMatrix(0, NETWORKINDEX) = "Netwk"
    grdCounts.TextMatrix(0, NETWORKBREAKINDEX) = "Break"
    grdCounts.TextMatrix(0, FEEDBREAKINDEX) = "Break"
    grdCounts.TextMatrix(0, FEEDSPOTINDEX) = "Feed (dat)"
    grdCounts.TextMatrix(0, FEEDNCINDEX) = "Feed (dat)"
    grdCounts.TextMatrix(0, PLEDGESPOTINDEX) = "Pledge (dat)"
    grdCounts.TextMatrix(0, PLEDGENCINDEX) = "Pledge (dat)"
    grdCounts.TextMatrix(0, SPOTINDEX) = "Spot (ast)"
    grdCounts.TextMatrix(0, NOTCARRIEDINDEX) = "Spot (ast)"
    grdCounts.TextMatrix(0, POSTBYINDEX) = "Spot (ast)"
    grdCounts.TextMatrix(0, MGINDEX) = "MG"
    grdCounts.TextMatrix(0, VENDOREXPORTINDEX) = "Vendor"
    grdCounts.TextMatrix(0, VENDORIMPORTINDEX) = "Vendor"
    grdCounts.TextMatrix(0, VENDORAPPLIEDINDEX) = "Vendor"
    grdCounts.TextMatrix(0, AGYCOMPLIANTINDEX) = "Compliance"
    grdCounts.TextMatrix(0, ADVTCOMPLIANTINDEX) = "Compliance"
    grdCounts.TextMatrix(0, ATTSEQNOINDEX) = "AttCode+SeqNo"
    grdCounts.TextMatrix(0, CODESINDEX) = "AttCode|VefCode|ShttCode"
    grdCounts.TextMatrix(0, SORTINDEX) = "Sort"

    grdCounts.TextMatrix(1, DATEINDEX) = ""
    grdCounts.TextMatrix(1, VEHICLEINDEX) = "Name"
    grdCounts.TextMatrix(1, STATIONINDEX) = ""
    grdCounts.TextMatrix(1, MULTICASTINDEX) = "cast"
    grdCounts.TextMatrix(1, POSTMETHODINDEX) = "Method"
    grdCounts.TextMatrix(1, FLOWINDEX) = ""
    grdCounts.TextMatrix(1, AIRPLAYSINDEX) = "Plays"
    grdCounts.TextMatrix(1, NETWORKINDEX) = "Count"
    grdCounts.TextMatrix(1, NETWORKBREAKINDEX) = "Netwk"
    grdCounts.TextMatrix(1, FEEDBREAKINDEX) = "Feed"
    grdCounts.TextMatrix(1, FEEDSPOTINDEX) = "Spots"
    grdCounts.TextMatrix(1, FEEDNCINDEX) = "NotC"
    grdCounts.TextMatrix(1, PLEDGESPOTINDEX) = "Spots"
    grdCounts.TextMatrix(1, PLEDGENCINDEX) = "NotC"
    grdCounts.TextMatrix(1, SPOTINDEX) = "Count"
    grdCounts.TextMatrix(1, NOTCARRIEDINDEX) = "NotC"
    grdCounts.TextMatrix(1, POSTBYINDEX) = "By"
    grdCounts.TextMatrix(1, MGINDEX) = "Count"
    grdCounts.TextMatrix(1, VENDOREXPORTINDEX) = "Expt"
    grdCounts.TextMatrix(1, VENDORIMPORTINDEX) = "Impt"
    grdCounts.TextMatrix(1, VENDORAPPLIEDINDEX) = "Applied"
    grdCounts.TextMatrix(1, AGYCOMPLIANTINDEX) = "Netwk"
    grdCounts.TextMatrix(1, ADVTCOMPLIANTINDEX) = "Station"
    grdCounts.TextMatrix(1, ATTSEQNOINDEX) = ""
    grdCounts.TextMatrix(1, MULTICASTSTATIONINDEX) = ""
    grdCounts.TextMatrix(1, CODESINDEX) = ""
    grdCounts.TextMatrix(1, SORTINDEX) = ""

    grdCounts.Row = 0
    grdCounts.MergeCells = 2    'flexMergeRestrictColumns
    grdCounts.MergeRow(0) = True

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim llVef As Long
    Dim llShtt As Long
    Dim llMajor As Long
    Dim llMinor As Long

    Dim slDate As String
    Dim llLstCount As Long
    Dim blDatExist As Boolean
    Dim llBaseStation As Long
    Dim llBaseDate As Long
    Dim llBaseVehicle As Long
    Dim llBaseSort As Long
    Dim llVendorRow As Long
    Dim llVendorExportCount As Long
    Dim llVendorImportCount As Long
    
    Dim llUnpostedSpotCount As Long
    Dim llPostedSpotCount As Long
    Dim llWebUnpostedSpotCount As Long
    Dim llWebPostedSpotCount As Long
    Dim llNCUnpostedSpotCount As Long
    Dim llNCPostedSpotCount As Long
    
    Dim llMGInMissedOutCount As Long
    Dim llMGInMissedInCount As Long
 
    Dim ilVendorId As Integer
    Dim llDate As Long
    Dim llColor As Long
    
    Dim blOutBalance As Boolean
    Dim blNotCompliant As Boolean
    Dim blPartiallyPosted As Boolean
    Dim blUserPosting As Boolean
    Dim blBreakOutBalance As Boolean
    Dim llAgreementStartRow As Long
    Dim blAffiliateSpotExist As Boolean
    
    On Error GoTo ErrHand
    lmTotalOkCount = 0
    lmTotalErrorCount = 0
    bmRetrieveLst = True
    llAgreementStartRow = -1
    gGrid_Clear grdCounts, True
    grdCounts.Row = 0
'    For llCol = VEHICLEINDEX To STATIONINDEX Step 1
'        grdCounts.Col = llCol
'        grdCounts.CellBackColor = LIGHTBLUE
'    Next llCol
    llRow = grdCounts.FixedRows
    llBaseDate = llRow - 1
    grdCounts.Redraw = False
    lmRowColor = LIGHTGRAY
    lmStartDate = gDateValue(smFeedStartDate)
    lmEndDate = gDateValue(smFeedEndDate)
    lmWkStartDate = gDateValue(smFeedStartDate)
    lmWkEndDate = gDateValue(gObtainNextSunday(smFeedStartDate))
    If lmWkEndDate > lmEndDate Then
        lmWkEndDate = lmEndDate
    End If
    Do While lmWkStartDate <= lmEndDate
        smSQLStartDate = Format(lmWkStartDate, sgSQLDateForm)
        smSQLEndDate = Format(lmWkEndDate, sgSQLDateForm)
        smSQLMoDate = Format(gObtainPrevMonday(Format(lmWkStartDate, sgShowDateForm)), sgSQLDateForm)
        smSQLSuDate = Format(gObtainNextSunday(Format(lmWkStartDate, sgShowDateForm)), sgSQLDateForm)
        For llMajor = 0 To UBound(lmMajor) - 1 Step 1
            If imSort = 0 Then
                imVefCode = lmMajor(llMajor)
            Else
                imShttCode = lmMajor(llMajor)
            End If
            For llMinor = 0 To UBound(lmMinor) - 1 Step 1
                DoEvents
                If (bmTerminate) Then
                    gSetMousePointer grdCounts, grdCounts, vbDefault
                    grdCounts.Redraw = True
                    cmcDone.Caption = "Close"
                    Exit Sub
                End If
                If imSort = 0 Then
                    imShttCode = lmMinor(llMinor)
                Else
                    imVefCode = lmMinor(llMinor)
                End If
                
                llVef = gBinarySearchVef(CLng(imVefCode))
                llShtt = gBinarySearchStationInfoByCode(imShttCode)
                'Get agreement
                smSQLQuery = "Select attCode, attNoAirPlays, attExportType, attLoad, attMulticast, attPostingType from att where attVefCode = " & imVefCode & " And attShfCode = " & imShttCode
                smSQLQuery = smSQLQuery + " And attOnAir <= '" & smSQLStartDate & "'"
                smSQLQuery = smSQLQuery + " And attOffAir >= '" & smSQLStartDate & "'"
                smSQLQuery = smSQLQuery + " And attDropDate >= '" & smSQLStartDate & "'"
                Set rst_att = gSQLSelectCall(smSQLQuery)
                lmAttCode = -1
                If Not rst_att.EOF Then lmAttCode = rst_att!attCode
                If (lmAttCode > 0) And (llVef <> -1) And (llShtt <> -1) Then
                    If (tgVehicleInfo(llVef).sVehType = "G") Then
                        smSQLQuery = "SELECT Count(1) "
                        smSQLQuery = smSQLQuery + " FROM LCF_Log_Calendar"
                        smSQLQuery = smSQLQuery + " WHERE (lcfVefCode = " & tgVehicleInfo(llVef).iCode
                        smSQLQuery = smSQLQuery + " AND lcfLogDate BETWEEN '" & smSQLStartDate & "' And '" & smSQLEndDate & "')"
                        Set rst_lcf = gSQLSelectCall(smSQLQuery)
                        If rst_lcf.EOF Then
                            lmAttCode = -1
                        ElseIf rst_lcf(0).Value <= 0 Then
                            lmAttCode = -1
                        End If
                    End If
                End If
                If (lmAttCode > 0) And (llVef <> -1) And (llShtt <> -1) Then
                    If llRow >= grdCounts.Rows Then
                        grdCounts.AddItem ""
                    End If
                    grdCounts.Row = llRow
                    If lmWkStartDate = lmWkEndDate Then
                        slDate = Format(lmWkStartDate, "m/d")
                    Else
                        slDate = Format(lmWkStartDate, "m/d") & "-" & Format(lmWkEndDate, "m/d")
                    End If
                    
                    If (grdCounts.TextMatrix(llBaseDate, DATEINDEX) <> slDate) Then
                        blOutBalance = False
                        blNotCompliant = False
                        blPartiallyPosted = False
                        blBreakOutBalance = False
                        llAgreementStartRow = llRow
                        llBaseStation = llRow
                        llBaseDate = llRow
                        llBaseVehicle = llRow
                        lmRowColor = Switch(lmRowColor = vbWhite, LIGHTGRAY, lmRowColor = LIGHTGRAY, vbWhite)
                        bmRetrieveLst = True
                        grdCounts.TextMatrix(llRow, DATEINDEX) = slDate
                        grdCounts.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                        grdCounts.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStationInfoByCode(llShtt).sCallLetters)
                        ReDim tgDat(0 To 0) As DAT
                    Else
                        If (grdCounts.TextMatrix(llBaseVehicle, VEHICLEINDEX) <> Trim$(tgVehicleInfo(llVef).sVehicle)) Then
                            blOutBalance = False
                            blNotCompliant = False
                            blPartiallyPosted = False
                            blBreakOutBalance = False
                            llAgreementStartRow = llRow
                            'llBaseStation = llRow
                            'llBaseDate = llRow
                            llBaseVehicle = llRow
                            lmRowColor = Switch(lmRowColor = vbWhite, LIGHTGRAY, lmRowColor = LIGHTGRAY, vbWhite)
                            bmRetrieveLst = True
                            If imSort = 0 Then
                                llBaseDate = llRow
                                llBaseStation = llRow
                                grdCounts.TextMatrix(llRow, DATEINDEX) = slDate
                            Else
                                If (grdCounts.TextMatrix(llBaseStation, STATIONINDEX) <> Trim$(tgStationInfoByCode(llShtt).sCallLetters)) Then
                                    llBaseDate = llRow
                                    llBaseStation = llRow
                                    grdCounts.TextMatrix(llRow, DATEINDEX) = slDate
                                End If
                            End If
                            grdCounts.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                            grdCounts.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStationInfoByCode(llShtt).sCallLetters)
                            ReDim tgDat(0 To 0) As DAT
                        ElseIf (grdCounts.TextMatrix(llBaseStation, STATIONINDEX) <> Trim$(tgStationInfoByCode(llShtt).sCallLetters)) Then
                            blOutBalance = False
                            blNotCompliant = False
                            blPartiallyPosted = False
                            blBreakOutBalance = False
                            llAgreementStartRow = llRow
                            llBaseStation = llRow
                            If imSort = 1 Then
                                llBaseDate = llRow
                                grdCounts.TextMatrix(llRow, DATEINDEX) = slDate
                            End If
                            lmRowColor = Switch(lmRowColor = vbWhite, LIGHTGRAY, lmRowColor = LIGHTGRAY, vbWhite)
                            grdCounts.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                            grdCounts.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStationInfoByCode(llShtt).sCallLetters)
                            ReDim tgDat(0 To 0) As DAT
                        End If
                    End If
                    grdCounts.TextMatrix(llRow, CODESINDEX) = "AttCode = " & lmAttCode & ", VefCode = " & imVefCode & ", ShttCode = " & imShttCode

                    For llCol = DATEINDEX To ADVTCOMPLIANTINDEX Step 1
                        grdCounts.Col = llCol
                        grdCounts.CellBackColor = lmRowColor
                    Next llCol
                    
                    'Post Method
                    If rst_att!attExportType = 0 Then
                        ReDim imVendorID(0 To 0) As Integer
                        grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Manual"
                    Else
                        'Build vendor table for the agreement
                        mBuildVendorID
                        'Wrong place
                        If UBound(imVendorID) > LBound(imVendorID) Then
                            grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Vendor"
                        Else
                            grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Web"
                        End If
                    End If
                    
                    'Multicast: must be after Posting method
                    mMulticast llRow
                                        
                    'Air Plays
                    If rst_att!attNoAirPlays > 1 Then
                        grdCounts.TextMatrix(llRow, AIRPLAYSINDEX) = rst_att!attNoAirPlays
                    Else
                        If rst_att!attLoad > 1 Then
                            grdCounts.TextMatrix(llRow, AIRPLAYSINDEX) = rst_att!attLoad
                        Else
                            grdCounts.TextMatrix(llRow, AIRPLAYSINDEX) = "1"
                        End If
                    End If
                    
                    
                    'Get Network count
                    grdCounts.TextMatrix(llRow, NETWORKINDEX) = mGetNetworkCount()
                    If grdCounts.TextMatrix(llRow, NETWORKINDEX) = "0" Then
                        grdCounts.Col = NETWORKINDEX
                        grdCounts.CellForeColor = vbRed
                        blOutBalance = True
                    End If
                    
                    'Test if affiliate spots exist to determine Flow
                    blAffiliateSpotExist = mAffiliateSpotExist()
                    If Not blAffiliateSpotExist Then
                        If Val(grdCounts.TextMatrix(llRow, NETWORKINDEX)) > 0 Then
                            grdCounts.TextMatrix(llRow, FLOWINDEX) = "T->A"
                            If rst_att!attExportType = 1 Then
                                grdCounts.Col = FLOWINDEX
                                grdCounts.CellForeColor = vbRed
                                blOutBalance = True
                            End If
                        Else
                            grdCounts.TextMatrix(llRow, FLOWINDEX) = "T->L"
                            grdCounts.Col = VEHICLEINDEX
                            grdCounts.CellForeColor = vbRed
                            blOutBalance = True
                        End If
                    Else
                        'check if spots sent to web to determine if manual or web source. Look at ent
                        'Check if exporting to vendor
                        If rst_att!attExportType = 0 Then
                            grdCounts.TextMatrix(llRow, FLOWINDEX) = "A->M"
                        Else
                            grdCounts.TextMatrix(llRow, FLOWINDEX) = "A->W"
                            'Determine if not exported to web
                            If mSpotsExportedToWeb() = False Then
                                grdCounts.Col = FLOWINDEX
                                grdCounts.CellForeColor = vbRed
                                blOutBalance = True
                            End If
                        End If
                    End If
                                        
                    smSQLQuery = "Select * From cptt Where cpttAtfCode = " & lmAttCode
                    smSQLQuery = smSQLQuery & " And cpttStartDate = '" & smSQLMoDate & "'"
                    Set rst_Cptt = gSQLSelectCall(smSQLQuery)
                    
                                        
                    'determine if pledge defined
                    'blDatExist = mDatExist()
                    blDatExist = mGetPledgeInfo()
                                        
                    'Network Break count
                    grdCounts.TextMatrix(llRow, NETWORKBREAKINDEX) = mGetNetworkBreakCount()
                    
                    If rst_att!attPostingType <> 0 Then
                        'Feed Break count
                        grdCounts.TextMatrix(llRow, FEEDBREAKINDEX) = mGetFeedBreakCount(blDatExist)
                        If bmAirPlayConflict Then
                            grdCounts.Col = AIRPLAYSINDEX
                            grdCounts.CellForeColor = vbMagenta
                        End If
                        
                        If Val(grdCounts.TextMatrix(llRow, NETWORKBREAKINDEX)) <> Val(grdCounts.TextMatrix(llRow, FEEDBREAKINDEX)) Then
                            grdCounts.Col = NETWORKBREAKINDEX
                            grdCounts.CellForeColor = vbRed
                            grdCounts.Col = FEEDBREAKINDEX
                            grdCounts.CellForeColor = vbRed
                            blBreakOutBalance = True
                        End If
                        'Agreement Feed count: Network times Air Plays
                        mGetFeedSpotCount llRow, blDatExist   'Val(grdCounts.TextMatrix(llRow, AIRPLAYSINDEX)) * Val(grdCounts.TextMatrix(llRow, NETWORKINDEX))
                        
                        'Agreement Pledge count
                        grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX) = mGetPledgeSpotCount(blDatExist)
                        
                        'Pledge Not Carried. If Dat does not exist, then all spots are treated as if the pledge was set to Live. Not carry will not exist when generating the affiliate spots
                        If (blDatExist) Then
                            grdCounts.TextMatrix(llRow, PLEDGENCINDEX) = mGetPledgeNotCarried()
                        Else
                            grdCounts.TextMatrix(llRow, PLEDGENCINDEX) = "0"
                        End If
                        
                        'If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) <> Val(grdCounts.TextMatrix(llRow, FEEDSPOTINDEX)) Then
                        '    grdCounts.Col = PLEDGESPOTINDEX
                        '    grdCounts.CellForeColor = vbRed
                        '    blOutBalance = True
                        'End If
                        If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, PLEDGENCINDEX)) <> Val(grdCounts.TextMatrix(llRow, FEEDSPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, FEEDNCINDEX)) Then
                            grdCounts.Col = FEEDSPOTINDEX
                            grdCounts.CellForeColor = vbRed
                            grdCounts.Col = FEEDNCINDEX
                            grdCounts.CellForeColor = vbRed
                            grdCounts.Col = PLEDGESPOTINDEX
                            grdCounts.CellForeColor = vbRed
                            grdCounts.Col = PLEDGENCINDEX
                            grdCounts.CellForeColor = vbRed
                            blOutBalance = True
                        End If
                        
                        grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX) = mGetPostedNotCarried() + mGetUnpostedNotCarried()
                        
                        'Get spots
                        llUnpostedSpotCount = mGetUnpostedSpotCount()
                        llPostedSpotCount = mGetPostedSpotCount()
                    Else
                        llUnpostedSpotCount = 0
                        llPostedSpotCount = 0
                    End If
                    
                    If grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Manual" Then
                        If (Not rst_Cptt.EOF) Then
                            If rst_att!attPostingType <> 0 Then
                                If rst_Cptt!cpttPostingStatus = 0 Then
                                    grdCounts.TextMatrix(llRow, SPOTINDEX) = llUnpostedSpotCount + llPostedSpotCount
                                    If (llUnpostedSpotCount > 0) And (llPostedSpotCount > 0) Then
                                        grdCounts.Col = SPOTINDEX
                                        grdCounts.CellForeColor = vbMagenta
                                        blPartiallyPosted = True
                                    ElseIf (llPostedSpotCount > 0) Then
                                        grdCounts.Col = SPOTINDEX
                                        grdCounts.CellForeColor = vbMagenta
                                        blPartiallyPosted = True
                                    End If
                                    If (llUnpostedSpotCount + llPostedSpotCount + Val(grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX)) <> Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, PLEDGENCINDEX))) Then
                                        grdCounts.Col = PLEDGESPOTINDEX
                                        grdCounts.CellForeColor = vbRed
                                        grdCounts.Col = PLEDGENCINDEX
                                        grdCounts.CellForeColor = vbRed
                                        grdCounts.Col = SPOTINDEX
                                        grdCounts.CellForeColor = vbRed
                                        grdCounts.Col = NOTCARRIEDINDEX
                                        grdCounts.CellForeColor = vbRed
                                        blOutBalance = True
                                    End If
                                End If
                            Else
                                If rst_Cptt!cpttPostingStatus = 0 Then
                                    grdCounts.TextMatrix(llRow, FLOWINDEX) = "NotRec'd"
                                Else
                                    grdCounts.TextMatrix(llRow, FLOWINDEX) = "Rec'd"
                                End If
                            End If
                        End If
                    Else
                        If bmIncludeWeb Then
                            llWebUnpostedSpotCount = mGetWebUnpostedSpotCount()
                            llWebPostedSpotCount = mGetWebPostedSpotCount()
                            If (llWebUnpostedSpotCount > 0) Or (llWebPostedSpotCount > 0) Then
                                grdCounts.TextMatrix(llRow, SPOTINDEX) = llWebUnpostedSpotCount + llWebPostedSpotCount
                                If (llWebUnpostedSpotCount > 0) And (llWebPostedSpotCount > 0) Then
                                    grdCounts.Col = SPOTINDEX
                                    grdCounts.CellForeColor = vbMagenta
                                    blPartiallyPosted = True
                                End If
                                'Not Carried not sent to Web
                                If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) <> llWebUnpostedSpotCount + llWebPostedSpotCount Then
                                    grdCounts.Col = PLEDGESPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    grdCounts.Col = SPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blOutBalance = True
                                End If
                            Else
                                If (grdCounts.TextMatrix(llRow, FLOWINDEX) = "A->W") Then
                                    grdCounts.TextMatrix(llRow, SPOTINDEX) = "0"
                                    grdCounts.Col = SPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blOutBalance = True
                                ElseIf (Not blAffiliateSpotExist) Then
                                    grdCounts.TextMatrix(llRow, SPOTINDEX) = "0"
                                End If
                            End If
                        Else
                            If grdCounts.TextMatrix(llRow, FLOWINDEX) = "A->W" Then
                                grdCounts.TextMatrix(llRow, SPOTINDEX) = "-"
                                llWebUnpostedSpotCount = 0
                                llWebPostedSpotCount = 0
                            End If
                        End If
                    End If
                    
                    'Loop on vendors and obtain vendor posting info
                    If imSort = 0 Then
                        llBaseSort = llBaseStation
                    Else
                        llBaseSort = llBaseVehicle
                    End If
                    
                    If UBound(imVendorID) > LBound(imVendorID) Then
                        For ilVendorId = 0 To UBound(imVendorID) - 1 Step 1
                            If ilVendorId = 0 Then
                                If imSort = 0 Then
                                    llVendorRow = llBaseStation
                                Else
                                    llVendorRow = llBaseVehicle
                                End If
                            Else
                                llRow = llRow + 1
                                If llRow >= grdCounts.Rows Then
                                    grdCounts.AddItem ""
                                End If
                                grdCounts.Row = llRow
                                For llCol = DATEINDEX To ADVTCOMPLIANTINDEX Step 1
                                    grdCounts.Col = llCol
                                    grdCounts.CellBackColor = lmRowColor
                                Next llCol
                                llVendorRow = llRow

                                ''Air Play
                                'grdCounts.TextMatrix(llVendorRow, AIRPLAYSINDEX) = grdCounts.TextMatrix(llBaseSort, AIRPLAYSINDEX)
    
                                'Network
                                grdCounts.TextMatrix(llVendorRow, NETWORKINDEX) = grdCounts.TextMatrix(llBaseSort, NETWORKINDEX)
                                mCopyColor llBaseSort, llVendorRow, NETWORKINDEX
                                
                                'Agreement
                                grdCounts.TextMatrix(llVendorRow, FEEDSPOTINDEX) = grdCounts.TextMatrix(llBaseSort, FEEDSPOTINDEX)
                                grdCounts.TextMatrix(llVendorRow, FEEDNCINDEX) = grdCounts.TextMatrix(llBaseSort, FEEDNCINDEX)
                                grdCounts.TextMatrix(llVendorRow, PLEDGESPOTINDEX) = grdCounts.TextMatrix(llBaseSort, PLEDGESPOTINDEX)
                                grdCounts.TextMatrix(llVendorRow, PLEDGENCINDEX) = grdCounts.TextMatrix(llBaseSort, PLEDGENCINDEX)
                                mCopyColor llBaseSort, llVendorRow, FEEDSPOTINDEX
                                mCopyColor llBaseSort, llVendorRow, FEEDNCINDEX
                                mCopyColor llBaseSort, llVendorRow, PLEDGESPOTINDEX
                                mCopyColor llBaseSort, llVendorRow, PLEDGENCINDEX

                                'Not Carried
                                grdCounts.TextMatrix(llVendorRow, NOTCARRIEDINDEX) = grdCounts.TextMatrix(llBaseSort, NOTCARRIEDINDEX)
                                
                                If Val(grdCounts.TextMatrix(llVendorRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llVendorRow, PLEDGENCINDEX)) <> Val(grdCounts.TextMatrix(llVendorRow, FEEDSPOTINDEX)) + Val(grdCounts.TextMatrix(llVendorRow, FEEDNCINDEX)) Then
                                    grdCounts.Col = FEEDSPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    grdCounts.Col = FEEDNCINDEX
                                    grdCounts.CellForeColor = vbRed
                                    grdCounts.Col = PLEDGESPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    grdCounts.Col = PLEDGENCINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blOutBalance = True
                                End If
                                
                                'Spots
                                'Note: grdCounts.TextMatrix(llRow, SPOTINDEX) = llWebUnpostedSpotCount + llWebPostedSpotCount
                                grdCounts.TextMatrix(llVendorRow, SPOTINDEX) = grdCounts.TextMatrix(llBaseSort, SPOTINDEX)
                                If (llWebUnpostedSpotCount > 0) And (llWebPostedSpotCount > 0) Then
                                    grdCounts.Col = SPOTINDEX
                                    grdCounts.CellForeColor = vbMagenta
                                    blPartiallyPosted = True
                                End If
                                If (Val(grdCounts.TextMatrix(llVendorRow, SPOTINDEX)) <> Val(grdCounts.TextMatrix(llVendorRow, PLEDGESPOTINDEX))) Then
                                    grdCounts.Col = PLEDGESPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    grdCounts.Col = SPOTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blOutBalance = True
                                End If
                                grdCounts.Row = llBaseSort
                                grdCounts.Col = FLOWINDEX
                                llColor = grdCounts.CellForeColor
                                grdCounts.Row = llVendorRow
                                grdCounts.Col = FLOWINDEX
                                grdCounts.CellForeColor = llColor
                            End If
                            ''Network
                            'grdCounts.TextMatrix(llVendorRow, NETWORKINDEX) = grdCounts.TextMatrix(llBaseSort, NETWORKINDEX)
                            
                            ''Air Play
                            'grdCounts.TextMatrix(llVendorRow, AIRPLAYSINDEX) = grdCounts.TextMatrix(llBaseSort, AIRPLAYSINDEX)

                            ''Agreement
                            'grdCounts.TextMatrix(llVendorRow, FEEDSPOTINDEX) = grdCounts.TextMatrix(llBaseSort, FEEDSPOTINDEX)
                            'grdCounts.TextMatrix(llVendorRow, PLEDGESPOTINDEX) = grdCounts.TextMatrix(llBaseSort, PLEDGESPOTINDEX)

                            'Determine if export or export/import or import
                            If bmIncludeWeb Then
                                llVendorExportCount = mGetVendorExportCount(imVendorID(ilVendorId))
                                llVendorImportCount = mGetVendorImportCount(imVendorID(ilVendorId))
                                grdCounts.TextMatrix(llVendorRow, VENDOREXPORTINDEX) = llVendorExportCount
                                grdCounts.TextMatrix(llVendorRow, VENDORIMPORTINDEX) = llVendorImportCount
                                If (llVendorExportCount > 0) And (llWebUnpostedSpotCount + llWebPostedSpotCount <> llVendorExportCount) Then
                                    grdCounts.Col = VENDOREXPORTINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blOutBalance = True
                                End If
                                'If imported, get spot count from Spots matching on Vendor Source (WO, NC)
                                If llVendorExportCount > 0 And llVendorImportCount > 0 Then
                                    grdCounts.TextMatrix(llVendorRow, FLOWINDEX) = "A->W<->" & gVendorInitials(imVendorID(ilVendorId))
                                ElseIf llVendorExportCount > 0 And llVendorImportCount = 0 Then
                                    grdCounts.TextMatrix(llVendorRow, FLOWINDEX) = "A->W->" & gVendorInitials(imVendorID(ilVendorId))
                                    grdCounts.TextMatrix(llVendorRow, VENDORIMPORTINDEX) = ""
                                ElseIf llVendorExportCount = 0 And llVendorImportCount > 0 Then
                                    grdCounts.TextMatrix(llVendorRow, FLOWINDEX) = "A->W<-" & gVendorInitials(imVendorID(ilVendorId))
                                    grdCounts.TextMatrix(llVendorRow, VENDOREXPORTINDEX) = ""
                                Else
                                    grdCounts.TextMatrix(llVendorRow, FLOWINDEX) = "A->W->" & gVendorInitials(imVendorID(ilVendorId))
                                    grdCounts.TextMatrix(llVendorRow, VENDOREXPORTINDEX) = ""
                                    grdCounts.TextMatrix(llVendorRow, VENDORIMPORTINDEX) = ""
                                End If
                                grdCounts.TextMatrix(llVendorRow, VENDORAPPLIEDINDEX) = mGetWebByVendorSpotCount(imVendorID(ilVendorId))
                            Else
                                grdCounts.TextMatrix(llVendorRow, FLOWINDEX) = "A->W->" & gVendorInitials(imVendorID(ilVendorId))
                                grdCounts.TextMatrix(llVendorRow, VENDOREXPORTINDEX) = "-"
                                grdCounts.TextMatrix(llVendorRow, VENDORIMPORTINDEX) = "-"
                                grdCounts.TextMatrix(llVendorRow, VENDORAPPLIEDINDEX) = "-"
                                llVendorExportCount = 0
                                llVendorImportCount = 0
                            End If
                        Next ilVendorId
                    End If
                    
                    '**** Posted ****
                    If (Not rst_Cptt.EOF) Then
                        If rst_Cptt!cpttPostingStatus <> 0 Then
                            If grdCounts.TextMatrix(llBaseVehicle, POSTMETHODINDEX) <> "Manual" Then
                                llRow = llRow + 1
                                If llRow >= grdCounts.Rows Then
                                    grdCounts.AddItem ""
                                End If
                                grdCounts.Row = llRow
                                For llCol = DATEINDEX To ADVTCOMPLIANTINDEX Step 1
                                    grdCounts.Col = llCol
                                    grdCounts.CellBackColor = lmRowColor
                                Next llCol
    
                                grdCounts.TextMatrix(llRow, FLOWINDEX) = "W->A"
                                                                
                                ''Air Play
                                'grdCounts.TextMatrix(llRow, AIRPLAYSINDEX) = grdCounts.TextMatrix(llBaseSort, AIRPLAYSINDEX)
                                
                                'Network
                                grdCounts.TextMatrix(llRow, NETWORKINDEX) = grdCounts.TextMatrix(llBaseSort, NETWORKINDEX)
                                mCopyColor llBaseSort, llRow, NETWORKINDEX
    
                                'Agreement
                                grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = grdCounts.TextMatrix(llBaseSort, FEEDSPOTINDEX)
                                grdCounts.TextMatrix(llRow, FEEDNCINDEX) = grdCounts.TextMatrix(llBaseSort, FEEDNCINDEX)
                                grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX) = grdCounts.TextMatrix(llBaseSort, PLEDGESPOTINDEX)
                                grdCounts.TextMatrix(llRow, PLEDGENCINDEX) = grdCounts.TextMatrix(llBaseSort, PLEDGENCINDEX)
                                If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, PLEDGENCINDEX)) <> Val(grdCounts.TextMatrix(llRow, FEEDSPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, FEEDNCINDEX)) Then
                                    mCopyColor llBaseSort, llRow, FEEDSPOTINDEX
                                    mCopyColor llBaseSort, llRow, FEEDNCINDEX
                                    mCopyColor llBaseSort, llRow, PLEDGESPOTINDEX
                                    mCopyColor llBaseSort, llRow, PLEDGENCINDEX
                                End If
    
                                'Not Carried
                                llNCUnpostedSpotCount = mGetUnpostedNotCarried()
                                llNCPostedSpotCount = mGetPostedNotCarried()
                                If llNCUnpostedSpotCount + llNCPostedSpotCount > 0 Then
                                    grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX) = llNCUnpostedSpotCount + llNCPostedSpotCount
                                    If llNCPostedSpotCount > 0 And llNCUnpostedSpotCount > 0 Then
                                        grdCounts.Col = NOTCARRIEDINDEX
                                        grdCounts.CellForeColor = vbMagenta
                                        blPartiallyPosted = True
                                    ElseIf llNCPostedSpotCount = 0 And llNCUnpostedSpotCount > 0 Then
                                        grdCounts.Col = NOTCARRIEDINDEX
                                        grdCounts.CellForeColor = vbMagenta
                                        blPartiallyPosted = True
                                    End If

                                Else
                                    grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX) = "0"
                                End If
                                
                                'Posted By
                                grdCounts.TextMatrix(llRow, POSTBYINDEX) = mGetPostedBy(llBaseSort)
                                    
                                'If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) <> Val(grdCounts.TextMatrix(llRow, SPOTINDEX)) Then
                                '    grdCounts.Col = PLEDGESPOTINDEX
                                '    grdCounts.CellForeColor = vbRed
                                '    blOutBalance = True
                                'End If
                                
                                'If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, PLEDGENCINDEX)) <> Val(grdCounts.TextMatrix(llRow, SPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX)) Then
                                '    grdCounts.Col = PLEDGESPOTINDEX
                                '    grdCounts.CellForeColor = vbRed
                                '    grdCounts.Col = PLEDGENCINDEX
                                '    grdCounts.CellForeColor = vbRed
                                '    grdCounts.Col = SPOTINDEX
                                '    grdCounts.CellForeColor = vbRed
                                '    grdCounts.Col = NOTCARRIEDINDEX
                                '    grdCounts.CellForeColor = vbRed
                                '    blOutBalance = True
                                'End If
                            End If
                            
                            'Spot Count
                            If rst_att!attPostingType <> 0 Then
                                llUnpostedSpotCount = mGetUnpostedSpotCount()
                                llPostedSpotCount = mGetPostedSpotCount()
                                grdCounts.TextMatrix(llRow, SPOTINDEX) = llUnpostedSpotCount + llPostedSpotCount
                            Else
                                llUnpostedSpotCount = 0
                                llPostedSpotCount = 0
                            End If
                            If (llUnpostedSpotCount > 0) And (llPostedSpotCount > 0) Then
                                grdCounts.Col = SPOTINDEX
                                grdCounts.CellForeColor = vbMagenta
                                blPartiallyPosted = True
                            ElseIf (llUnpostedSpotCount > 0) Then
                                grdCounts.Col = SPOTINDEX
                                grdCounts.CellForeColor = vbMagenta
                                blPartiallyPosted = True
                            End If
                            'If (llUnpostedSpotCount + llPostedSpotCount <> Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX))) Then
                            '    grdCounts.Col = SPOTINDEX
                            '    grdCounts.CellForeColor = vbRed
                            '    blOutBalance = True
                            'End If
                            If Val(grdCounts.TextMatrix(llRow, PLEDGESPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, PLEDGENCINDEX)) <> Val(grdCounts.TextMatrix(llRow, SPOTINDEX)) + Val(grdCounts.TextMatrix(llRow, NOTCARRIEDINDEX)) Then
                                grdCounts.Col = PLEDGESPOTINDEX
                                grdCounts.CellForeColor = vbRed
                                grdCounts.Col = PLEDGENCINDEX
                                grdCounts.CellForeColor = vbRed
                                grdCounts.Col = SPOTINDEX
                                grdCounts.CellForeColor = vbRed
                                grdCounts.Col = NOTCARRIEDINDEX
                                grdCounts.CellForeColor = vbRed
                                blOutBalance = True
                            End If
                            If grdCounts.TextMatrix(llBaseVehicle, POSTMETHODINDEX) <> "Manual" Then
                                If grdCounts.TextMatrix(llRow, POSTBYINDEX) = "A" Then
                                    grdCounts.Col = POSTBYINDEX
                                    grdCounts.CellForeColor = vbRed
                                    blUserPosting = True
                                End If
                            End If
                            
                            'MG status
                            llMGInMissedOutCount = mGetMGInMissedOutCount()
                            llMGInMissedInCount = mGetMGInMissedInCount()
                            If (llMGInMissedOutCount > 0) Or (llMGInMissedInCount > 0) Then
                                grdCounts.TextMatrix(llRow, MGINDEX) = llMGInMissedOutCount + llMGInMissedInCount
                            End If
                            If (llMGInMissedOutCount > 0) And (llMGInMissedInCount > 0) Then
                                grdCounts.Col = MGINDEX
                                grdCounts.CellForeColor = ORANGE
                            End If
                                                       
                            'Compliance
                            mCompliance llRow, blNotCompliant
                        End If
                    End If
                    'Show agreement
                    mShowAgreement llRow, llAgreementStartRow, blOutBalance, blNotCompliant, blPartiallyPosted, blUserPosting, blBreakOutBalance
                End If
            Next llMinor
            mShowAgreement llRow, llAgreementStartRow, blOutBalance, blNotCompliant, blPartiallyPosted, blUserPosting, blBreakOutBalance
        Next llMajor
        mShowAgreement llRow, llAgreementStartRow, blOutBalance, blNotCompliant, blPartiallyPosted, blUserPosting, blBreakOutBalance

        'Advance to next week
        lmWkStartDate = lmWkStartDate + 7
        lmWkStartDate = gDateValue(gObtainPrevMonday(Format(lmWkStartDate, sgShowDateForm)))
        lmWkEndDate = gDateValue(gObtainNextSunday(Format(lmWkStartDate, sgShowDateForm)))
        smSQLMoDate = Format(gObtainPrevMonday(Format(lmWkStartDate, sgShowDateForm)), sgSQLDateForm)
        smSQLSuDate = Format(gObtainNextSunday(Format(lmWkStartDate, sgShowDateForm)), sgSQLDateForm)
        If lmWkEndDate > lmEndDate Then
            lmWkEndDate = lmEndDate
        End If
    Loop
    mShowAgreement llRow, llAgreementStartRow, blOutBalance, blNotCompliant, blPartiallyPosted, blUserPosting, blBreakOutBalance
    
'
'    SQLQuery = "Select "
'    SQLQuery = SQLQuery & "* "
'    SQLQuery = SQLQuery & "From abf_AST_Build_Queue "
'    SQLQuery = SQLQuery & "Where "
'    SQLQuery = SQLQuery & " abfEnteredDate >= '" & Format(frmSpotCountSpec.edcFeedStartDate.Text, sgSQLDateForm) & "'"
'    SQLQuery = SQLQuery & " And abfEnteredDate <= '" & Format(frmSpotCountSpec.edcFeedEndDate.Text, sgSQLDateForm) & "'"
'    Set rst_abf = gSQLSelectCall(SQLQuery)
'    Do While Not rst_abf.EOF
'        If llRow >= grdCounts.Rows Then
'            grdCounts.AddItem ""
'        End If
'        grdCounts.Row = llRow
'        For llCol = VEHICLEINDEX To COMPLETEDINDEX Step 1
'            grdCounts.Col = llCol
'            grdCounts.CellBackColor = LIGHTYELLOW
'        Next llCol
'        llVef = gBinarySearchVef(CLng(rst_abf!abfVefCode))
'        If llVef <> -1 Then
'            grdCounts.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
'        Else
'            grdCounts.TextMatrix(llRow, VEHICLEINDEX) = "Vehicle Code = " & rst_abf!abfVefCode
'        End If
'        If rst_abf!abfShttCode > 0 Then
'            ilShtt = gBinarySearchStationInfoByCode(rst_abf!abfShttCode)
'            If ilShtt <> -1 Then
'                grdCounts.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
'            Else
'                grdCounts.TextMatrix(llRow, STATIONINDEX) = "Station Code = " & rst_abf!abfShttCode
'            End If
'        End If
'        If rst_abf!abfSource = "A" Then
'            grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = "Agreement"
'        ElseIf rst_abf!abfSource = "P" Then
'            grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = "Post Log"
'        ElseIf rst_abf!abfSource = "F" Then
'            grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = "Fast Add"
'        Else
'            grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = "Log"
'        End If
'        If rst_abf!abfStatus = "P" Then
'            grdCounts.TextMatrix(llRow, STATUSINDEX) = "Processing"
'        ElseIf rst_abf!abfStatus = "H" Then
'            grdCounts.TextMatrix(llRow, STATUSINDEX) = "On Hold"
'        ElseIf rst_abf!abfStatus = "C" Then
'            grdCounts.TextMatrix(llRow, STATUSINDEX) = "Completed"
'        ElseIf rst_abf!abfStatus = "G" Then
'            grdCounts.TextMatrix(llRow, STATUSINDEX) = "Ready"
'        End If
'        grdCounts.TextMatrix(llRow, GENDATEINDEX) = Format(rst_abf!abfGenStartDate, sgShowDateForm) & "-" & Format(rst_abf!abfGenEndDate, sgShowDateForm)
'        grdCounts.TextMatrix(llRow, ENTEREDINDEX) = Format(rst_abf!abfEnteredDate, sgShowDateForm) & " " & Format(rst_abf!abfEnteredTime, sgShowTimeWSecForm)
'        If gDateValue(Format(rst_abf!abfCompletedDate, sgShowDateForm)) <> gDateValue("12/31/2069") Then
'            grdCounts.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_abf!abfCompletedDate, sgShowDateForm) & " " & Format(rst_abf!abfCompletedTime, sgShowTimeWSecForm)
'        End If
'        grdCounts.TextMatrix(llRow, ABFCODEINDEX) = rst_abf!abfCode
'        llRow = llRow + 1
'        rst_abf.MoveNext
'    Loop
'    mStatusSortCol ENTEREDINDEX
'    mStatusSortCol ENTEREDINDEX
'    'mSetStatusGridColor
    gSetMousePointer grdCounts, grdCounts, vbDefault
    grdCounts.Redraw = True
    cmcDone.Caption = "Done"
    Exit Sub
ErrHand:
    gSetMousePointer grdCounts, grdCounts, vbDefault
    gHandleError "StationBuildQueueStatusLog.txt", "StationSpotBuilderQueue-mPopulate"
    grdCounts.Redraw = True
    Resume Next
End Sub


Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdCounts, True
    For llRow = grdCounts.FixedRows To grdCounts.Rows - 1 Step 1
'        For llCol = VEHICLEINDEX To COMPLETEDINDEX Step 1
'            grdCounts.Row = llRow
'            grdCounts.Col = llCol
'            If llCol = STATUSINDEX Then
'                If grdCounts.TextMatrix(llRow, llCol) = "Not Ready" Then
'                    grdCounts.CellBackColor = vbWhite
'                Else
'                    grdCounts.CellBackColor = LIGHTYELLOW
'                End If
'            Else
'                grdCounts.CellBackColor = LIGHTYELLOW
'            End If
'        Next llCol
    Next llRow
End Sub

Private Sub mClearGrid()
    gGrid_Clear grdCounts, True
End Sub



Private Sub grdCounts_Click()

    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim llNumberTextRows As Long
    Dim slCodes As String
    Dim llLoop As Long
    Dim slNetworkBreak As String
    Dim slFeedBreak As String
    Dim slNetworkSpots As String
    Dim slFeedSpots As String
    Dim slFeedNCSpots As String
    Dim slPledgeSpots As String
    Dim slPledgeNCSpots As String
    Dim slSpotCount As String
    Dim slSpotNCCount As String
    
    If grdCounts.MouseRow = 0 Then
        mClearSelection
        Exit Sub
    End If
    
    slNetworkBreak = "Network Break: Number of breaks found from the Traffic spots transferred to the Affiliate system, or if no breaks transferred, it is obtained from the Vehicle Program structure"
    slFeedBreak = "Feed Break: Number of breaks defined on the Agreement Pledge tab, or if no Pledge defined, it is obtained from the Vehicle Program structure"
    slNetworkSpots = "Network Spots: Traffic Log spots transferred to Affiliate system"
    slFeedSpots = "Feed Spots: Determined by applying the Agreement pledge rules (Excluding Not Carried) to the Network Spots"
    slFeedNCSpots = "Feed NotC: Determined by applying the Agreement pledge rule (Not Carried) to the Network Spots"
    slPledgeSpots = "Pledge Spots: Number of Affiliate spots generated by applying the Agreement pledge rules (Excluding Not Carried)"
    slPledgeNCSpots = "Pledge NotC: Number of Affiliate spots generated by applying the Agreement pledge rules (Not Carried)"
    slSpotCount = "Spot Count: Number of Affiliate or Web system spots (Aired and Missed)"
    slSpotNCCount = "Spot NotC: Number of Not Carried Affiliate system spots"
    
    llNumberTextRows = -1
    llNumberTextRows = 1
    edcGridInfo.Visible = False
    llRow = grdCounts.MouseRow
    llCol = grdCounts.MouseCol
    If ((lmHDLastCol = llCol) And (llRow = grdCounts.FixedRows - 1)) Or ((lmCellLastCol = llCol) And (lmCellLastRow = llRow)) Then
        mClearSelection
        Exit Sub
    End If
    mClearSelection
    lmHDLastCol = -1
    lmCellLastCol = -1
    lmCellLastRow = -1
    If (llRow = grdCounts.FixedRows - 1) Then
        Select Case llCol
            Case DATEINDEX:
                slStr = "Date: Week date range requested"
            Case VEHICLEINDEX:
                slStr = "Vehicle: Agreement vehicle name"
            Case STATIONINDEX:
                slStr = "Station: Affiliate Call Letters"
            Case MULTICASTINDEX:
                llNumberTextRows = 2
                slStr = "Multicast: *and ** indicate that station is defined as multicast within an agreement" + sgCRLF
                slStr = slStr + "** indicates that this station is defined as the master station"
            Case POSTMETHODINDEX:
                llNumberTextRows = 2
                slStr = "Post Method: Manner in which affiliate aired dates and times are specified" + sgCRLF
                slStr = slStr + "Manual (Affiliate Affidavit); Web or Vendor"
            Case FLOWINDEX:
                llNumberTextRows = 3
                slStr = "Flow: Indicates the flow of spots between the three systems: Affiliate; Web and Vendor Traffic" + sgCRLF
                slStr = slStr + "T=Traffic; L=Log; A=Affiliate; W=Web; WO=Wide Orbit; NC=Network Connect; MM=Mr Master; XD=X-Digital" + sgCRLF
                slStr = slStr + "Manual Posting Receipt Only: NotRec'd=Not Received; Rec'd=Received"
            Case AIRPLAYSINDEX:
                slStr = "Air Plays: Number of times the Station has agreed to Air the spots"
            Case NETWORKBREAKINDEX:
                llNumberTextRows = 2      'using 2 to avoid not seeing a line because of wrap around
                slStr = slNetworkBreak  '"Network Break: Number of breaks found from the Traffic spots transferred to the Affiliate system, or if no breaks transferred, it is obtained from the Vehicle Program structure"
            Case FEEDBREAKINDEX:
                llNumberTextRows = 2      'using 2 to avoid not seeing a line because of wrap around
                slStr = slFeedBreak  '"Feed Break: Number of breaks defined on the Agreement Pledge tab, or if no Pledge defined, it is obtained from the Vehicle Program structure"
            Case NETWORKINDEX:
                slStr = slNetworkSpots  '"Network Spots: Traffic Log spots transferred to Affiliate system"
            Case FEEDSPOTINDEX:
                slStr = slFeedSpots   'slFeedSpots '"Feed Spots: Determined by applying the Agreement pledge rules (Excluding Not Carried) to the Network Spots"
            Case FEEDNCINDEX:
                slStr = slFeedNCSpots   '"Feed NotC: Determined by applying the Agreement pledge rule (Not Carried) to the Network Spots"
            Case PLEDGESPOTINDEX:
                slStr = slPledgeSpots   '"Pledge Spots: Number of Affiliate spots generated by applying the Agreement pledge rules (Excluding Not Carried)"
            Case PLEDGENCINDEX:
                slStr = slPledgeNCSpots '"Pledge NotC: Number of Affiliate spots generated by applying the Agreement pledge rules (Not Carried)"
            Case SPOTINDEX:
                slStr = slSpotCount '"Spot Count: Number of Affiliate or Web system spots (Aired and Missed)"
            Case NOTCARRIEDINDEX:
                slStr = slSpotNCCount   '"Spot NotC: Number of Not Carried Affiliate system spots"
            Case POSTBYINDEX:
                llNumberTextRows = 2
                slStr = "Posted By: Source of system which posted the dates and times of Affiliate spots" + sgCRLF
                slStr = slStr + "A= Affiliate system; W=Web system; V=Vendor"
            Case MGINDEX:
                slStr = "MG Count: Number of Affiliate spots defined as MG's within the specified dates"
            Case VENDOREXPORTINDEX:
                slStr = "Vendor Export: Number of Web spots sent to Vendor"
            Case VENDORIMPORTINDEX:
                slStr = "Vendor Import: Number of Vendor spots returned to Web"
            Case VENDORAPPLIEDINDEX:
                slStr = "Vendor Applied: Number of Returned Vendor Spots in sync with Web spots"
            Case AGYCOMPLIANTINDEX:
                slStr = "Network Compliance: Number of spots that are contract compliant"
            Case ADVTCOMPLIANTINDEX:
                slStr = "Station Compliance: Number of spots that are Agreement pledge compliant"
        End Select
        If (slStr <> "") Then
            lmHDLastCol = llCol
            lmCellLastCol = -1
            lmCellLastRow = -1
            edcGridInfo.Height = llNumberTextRows * grdCounts.RowHeight(grdCounts.FixedRows) '+ grdCounts.RowHeight(grdCounts.FixedRows)
            edcGridInfo.Top = grdCounts.Top + grdCounts.FixedRows * grdCounts.RowHeight(grdCounts.FixedRows)
            edcGridInfo.Width = grdCounts.Width
            edcGridInfo.Left = grdCounts.Left '+ (grdCounts.Width - edcGridInfo.Width) / 2
            edcGridInfo.Text = slStr
            edcGridInfo.Visible = True
        End If
    Else
        If llRow < grdCounts.FixedRows Then
            Exit Sub
        End If
        If (grdCounts.TextMatrix(llRow, FLOWINDEX) = "") And bmIncludeCodes Then
            Exit Sub
        End If
        If (grdCounts.TextMatrix(llRow, llCol) = "") Then
            Exit Sub
        Else
            'Get cell info
            Select Case llCol
                Case DATEINDEX:
                    slStr = ""
                Case VEHICLEINDEX:
                    llNumberTextRows = 2
                    slStr = "If Red: Affiliate Network spots not generated." & sgCRLF
                    slStr = slStr & "Fix by: Generate the Traffic Log"
                Case STATIONINDEX:
                    slStr = ""
                Case MULTICASTINDEX:
                    llNumberTextRows = 2
                    slStr = "If Red: Master Station required for Vendor support." & sgCRLF
                    slStr = slStr & "Fix by: Define which of the multicast stations is the master in the Station Sister tab"
                Case POSTMETHODINDEX:
                    slStr = ""
                Case FLOWINDEX:
                    slStr = ""
                    If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->A", vbBinaryCompare) > 0 Then
                        llNumberTextRows = 2
                        slStr = "If Red: Affiliate spots not generated. " & sgCRLF
                        slStr = slStr & "Fix by: Either generate Export or run any Affiliate Spot report or Select week on the Affiliate Affidavit" & sgCRLF
                    End If
                    If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->W", vbBinaryCompare) > 0 Then
                        llNumberTextRows = 2
                        slStr = "If Red: Affiliate spots need to be sent to Web" & sgCRLF
                        slStr = slStr + "Fix by: Generate Web Export"
                    End If
                Case AIRPLAYSINDEX:
                    llNumberTextRows = 2
                    slStr = "If Magenta: Duplicate Breaks defined with same Air Play Number." & sgCRLF
                    slStr = slStr & "Suggestion: Each Duplicated Break should be defined with unique Air Play numbers"
                Case NETWORKBREAKINDEX:
                    llNumberTextRows = 3
                    slStr = "If Red: Network and Feed counts not matching." & sgCRLF
                    slStr = slStr & "Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slNetworkBreak, ":", " Definition:")
                Case FEEDBREAKINDEX:
                    llNumberTextRows = 3
                    slStr = "If Red: Network and Feed counts not matching." & sgCRLF
                    slStr = slStr & "Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slFeedBreak, ":", " Definition:")
                Case NETWORKINDEX:
                    llNumberTextRows = 3
                    slStr = "If Red: Affiliate Network spots not generated." & sgCRLF
                    slStr = slStr & "Fix by: Generate Traffic Log"
                    slStr = slStr & sgCRLF & Replace(slNetworkSpots, ":", " Definition:")
                Case FEEDSPOTINDEX:
                    llNumberTextRows = 4
                    slStr = "If Red: Feed and Pledge counts not matching" & sgCRLF
                    slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                    slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slFeedSpots, ":", " Definition:")
                Case FEEDNCINDEX:
                    llNumberTextRows = 4
                    slStr = "If Red: Feed and Pledge counts not matching" & sgCRLF
                    slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                    slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slFeedNCSpots, ":", " Definition:")
                Case PLEDGESPOTINDEX:
                    llNumberTextRows = 5
                    slStr = "If Red: Feed and Pledge Spot counts not matching" & sgCRLF
                    slStr = slStr & "If Red: Pledge and Spot counts not matching" & sgCRLF
                    slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                    slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slPledgeSpots, ":", " Definition:")
                Case PLEDGENCINDEX:
                    llNumberTextRows = 5
                    slStr = "If Red: Feed and Pledge Spot counts not matching" & sgCRLF
                    slStr = slStr & "If Red: Pledge and Spot counts not matching" & sgCRLF
                    slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                    slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure"
                    slStr = slStr & sgCRLF & Replace(slPledgeNCSpots, ":", " Definition:")
                Case SPOTINDEX:
                    llNumberTextRows = 4
                    slStr = "If Red: Pledge and Spot counts not matching" & sgCRLF
                    If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->W", vbBinaryCompare) > 0 Then
                        slStr = slStr & "Fix by: Generate Traffic Log and/or Export to Web" & sgCRLF
                        slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                        slStr = slStr & "If Magenta: Mixture of Web Posted and Unposted Spots, or Unposted in Posted Week exists"
                    Else
                        If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "W->A", vbBinaryCompare) > 0 Then
                            slStr = slStr & "Fix by: Generate Traffic Log or Re Import to Web" & sgCRLF
                            slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                            slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                        Else
                            If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->A", vbBinaryCompare) > 0 Then
                                slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                                slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                                slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                            Else
                                llNumberTextRows = 2
                                slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                            End If
                        End If
                    End If
                    llNumberTextRows = llNumberTextRows + 1
                    slStr = slStr & sgCRLF & Replace(slSpotCount, ":", " Definition:")
                Case NOTCARRIEDINDEX:
                    llNumberTextRows = 4
                    slStr = "If Red: Pledge and Spot counts not matching" & sgCRLF
                    If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->W", vbBinaryCompare) > 0 Then
                        slStr = slStr & "Fix by: Generate Traffic Log and/or Export to Web" & sgCRLF
                        slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                        slStr = slStr & "If Magenta: Mixture of Web Posted and Unposted Spots, or Unposted in Posted Week exists"
                    Else
                        If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "W->A", vbBinaryCompare) > 0 Then
                            slStr = slStr & "Fix by: Generate Traffic Log or Re Import to Web" & sgCRLF
                            slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                            slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                        Else
                            If InStr(1, grdCounts.TextMatrix(llRow, FLOWINDEX), "->A", vbBinaryCompare) > 0 Then
                                slStr = slStr & "Fix by: Generate Traffic Log" & sgCRLF
                                slStr = slStr & "If Log generation failed, Check: Selling/Airing Links, Time zone definition (Station and/or vehicle offset), Unsold breaks, Program structure" & sgCRLF
                                slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                            Else
                                llNumberTextRows = 2
                                slStr = slStr & "If Magenta: Mixture of Posted and Unposted, or Unposted in Posted Week exists"
                            End If
                        End If
                    End If
                    llNumberTextRows = llNumberTextRows + 1
                    slStr = slStr & sgCRLF & Replace(slSpotNCCount, ":", " Definition:")
                Case POSTBYINDEX:
                    llNumberTextRows = 2
                    slStr = "If Red: Spots posted within Affiliate system but Posting not Completed on Web" & sgCRLF
                    slStr = slStr & "Fix by: Complete the posting on Web"
                Case MGINDEX:
                    slStr = ""
                Case VENDOREXPORTINDEX:
                    llNumberTextRows = 2
                    slStr = "If Red: Not all Web Spot Exported to Vendor" & sgCRLF
                    slStr = slStr & "Fix by: Re-export Affiliate spots to Web"
                Case VENDORIMPORTINDEX:
                    slStr = ""
                Case VENDORAPPLIEDINDEX:
                    slStr = ""
                Case AGYCOMPLIANTINDEX:
                    slStr = "If Red: Not all Spots Network compliant"
                Case ADVTCOMPLIANTINDEX:
                    slStr = "If Red: Not all Spots Station compliant"
            End Select
            If (slStr <> "") Then
                If (Not bmIncludeCodes) And ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) Or (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0)) Then
                    For llLoop = llRow To grdCounts.FixedRows Step -1
                        slCodes = grdCounts.TextMatrix(llRow, CODESINDEX)
                        If slCodes <> "" Then
                            llNumberTextRows = llNumberTextRows + 1
                            slStr = slStr & sgCRLF & grdCounts.TextMatrix(llRow, CODESINDEX)
                            Exit For
                        End If
                    Next llLoop
                End If
                grdCounts.Row = llRow
                grdCounts.Col = llCol
                grdCounts.CellBackColor = LIGHTBLUE 'vbCyan
                lmCellLastCol = llCol
                lmCellLastRow = llRow
                lmHDLastCol = -1
                edcGridInfo.Height = llNumberTextRows * grdCounts.RowHeight(grdCounts.FixedRows) '+ grdCounts.RowHeight(grdCounts.FixedRows)
                'If grdCounts.Top + (grdCounts.FixedRows + llRow - grdCounts.TopRow + 1) * grdCounts.RowHeight(grdCounts.FixedRows) + edcGridInfo.Height < cmcDone.Top Then
                If grdCounts.Top + (grdCounts.FixedRows + llRow - grdCounts.TopRow + 1) * grdCounts.RowHeight(grdCounts.FixedRows) < cmcDone.Top Then
                    edcGridInfo.Top = grdCounts.Top + (grdCounts.FixedRows + llRow - grdCounts.TopRow + 1) * grdCounts.RowHeight(grdCounts.FixedRows)
                Else
                    edcGridInfo.Top = grdCounts.Top + (grdCounts.FixedRows + llRow - grdCounts.TopRow + 1) * grdCounts.RowHeight(grdCounts.FixedRows) - edcGridInfo.Height - 2 * grdCounts.RowHeight(grdCounts.FixedRows)
                End If
                edcGridInfo.Width = grdCounts.Width
                edcGridInfo.Left = grdCounts.Left '+ (grdCounts.Width - edcGridInfo.Width) / 2
                edcGridInfo.Text = slStr
                edcGridInfo.Visible = True
            End If
        End If
    End If
    
End Sub

Private Sub grdCounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    If (grdCounts.MouseRow >= grdCounts.FixedRows) And (grdCounts.TextMatrix(grdCounts.MouseRow, grdCounts.MouseCol)) <> "" Then
        If grdCounts.MouseCol = MULTICASTINDEX Then
            If grdCounts.TextMatrix(grdCounts.MouseRow, MULTICASTSTATIONINDEX) <> "" Then
                grdCounts.ToolTipText = Trim$(grdCounts.TextMatrix(grdCounts.MouseRow, MULTICASTSTATIONINDEX))
            Else
                grdCounts.ToolTipText = ""
            End If
        End If
    Else
        grdCounts.ToolTipText = ""
    End If
    
End Sub

Private Sub grdCounts_Scroll()
    If grdCounts.Redraw = False Then
        grdCounts.Redraw = True
        grdCounts.TopRow = lmTopRow
        grdCounts.Refresh
        grdCounts.Redraw = False
    End If
        cmcDone.SetFocus

End Sub





Private Sub mStatusSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    
'    For llRow = grdCounts.FixedRows To grdCounts.Rows - 1 Step 1
'        slStr = Trim$(grdCounts.TextMatrix(llRow, VEHICLEINDEX))
'        If slStr <> "" Then
'            If ilCol = GENDATEINDEX Then
'                slStr = Trim$(grdCounts.TextMatrix(llRow, GENDATEINDEX))
'                ilPos = InStr(1, slStr, "-", vbTextCompare)
'                If ilPos > 0 Then
'                    slDate = Left(slStr, ilPos - 1)
'                    slSort = Trim$(Str$(gDateValue(slDate)))
'                    Do While Len(slSort) < 6
'                        slSort = "0" & slSort
'                    Loop
'                Else
'                    slSort = "      "
'                End If
'            ElseIf ilCol = ENTEREDINDEX Then
'                slStr = Trim$(grdCounts.TextMatrix(llRow, ENTEREDINDEX))
'                ilPos = InStr(1, slStr, " ", vbTextCompare)
'                If ilPos > 0 Then
'                    slDate = Left(slStr, ilPos - 1)
'                    slTime = Mid(slStr, ilPos + 1)
'                    slSort = Trim$(Str$(gDateValue(slDate)))
'                    Do While Len(slSort) < 6
'                        slSort = "0" & slSort
'                    Loop
'                    slStr = Trim$(Str$(gTimeToLong(slTime, False)))
'                    Do While Len(slStr) < 6
'                        slStr = "0" & slStr
'                    Loop
'                    slSort = slSort & slStr
'                Else
'                    slSort = "            "
'                End If
'            ElseIf ilCol = COMPLETEDINDEX Then
'                slStr = Trim$(grdCounts.TextMatrix(llRow, COMPLETEDINDEX))
'                ilPos = InStr(1, slStr, " ", vbTextCompare)
'                If ilPos > 0 Then
'                    slDate = Left(slStr, ilPos - 1)
'                    slTime = Mid(slStr, ilPos + 1)
'                    slSort = Trim$(Str$(gDateValue(slDate)))
'                    Do While Len(slSort) < 6
'                        slSort = "0" & slSort
'                    Loop
'                    slStr = Trim$(Str$(gTimeToLong(slTime, False)))
'                    Do While Len(slStr) < 6
'                        slStr = "0" & slStr
'                    Loop
'                    slSort = slSort & slStr
'                Else
'                    slSort = "            "
'                End If
'            Else
'                slSort = UCase$(Trim$(grdCounts.TextMatrix(llRow, ilCol)))
'                If slSort = "" Then
'                    slSort = Chr(32)
'                End If
'            End If
'            slStr = grdCounts.TextMatrix(llRow, SORTINDEX)
'            ilPos = InStr(1, slStr, "|", vbTextCompare)
'            If ilPos > 1 Then
'                slStr = Left$(slStr, ilPos - 1)
'            End If
'            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
'                slRow = Trim$(Str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdCounts.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            Else
'                slRow = Trim$(Str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdCounts.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            End If
'        End If
'    Next llRow
'    If ilCol = imLastColSorted Then
'        imLastColSorted = SORTINDEX
'    Else
'        imLastColSorted = -1
'        imLastSort = -1
'    End If
'    gGrid_SortByCol grdCounts, VEHICLEINDEX, SORTINDEX, imLastColSorted, imLastSort
'    imLastColSorted = ilCol
End Sub


Private Sub imcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim llRow As Long
    Dim llBaseRow As Long
    Dim llCol As Long
    Dim hlTo As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slFed As String
    Dim ilTimeAdj As Integer
    Dim slZone As String
    Dim slOutStr As String
    Dim llAttCode As Long
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    
    mClearSelection
    
    slToFile = sgExportDirectory & sgClientName & "_SpotCountTie-out" & "_" & Format(Now, "mmddyy") & "_" & Format(Now, "hhmm") & ".csv"
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        ilRet = gMsgBox("Export Previously Created " & slDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Kill slToFile
    End If
    ilRet = gFileOpen(slToFile, "Output", hlTo)
    If ilRet <> 0 Then
        MsgBox "Open file " & slToFile & " failed Error #" & Str$(Err.Numner), vbOKOnly, "Open Failed"
        Exit Sub
    End If
    gSetMousePointer grdCounts, grdCounts, vbHourglass
    llBaseRow = grdCounts.FixedRows
    Print #hlTo, "Date,Vehicle,Station,Multicast,Posting Method,Flow,Air Play,Network Break,Feed Break,Network Count,Feed Count,Feed Not Carried,Pledge Count,Pledge Not Carried,Spt Count,Not Carried,Posted By,Makegood Count,Vendor Export Count,Vendor Import Count,Vendor Applied Count,Network Compliant,Advertiser Compliant"
    For llRow = grdCounts.FixedRows To grdCounts.Rows - 1 Step 1
        If grdCounts.TextMatrix(llRow, FLOWINDEX) <> "" Then
            slStr = grdCounts.TextMatrix(llRow, CODESINDEX)
            If slStr <> "" Then
                ilRet = gParseItem(slStr, 1, ",", slOutStr)
                ilRet = gParseItem(slOutStr, 2, "=", slOutStr)
                llAttCode = Val(slOutStr)
                ilRet = gParseItem(slStr, 2, ",", slOutStr)
                ilRet = gParseItem(slOutStr, 2, "=", slOutStr)
                ilVefCode = Val(slOutStr)
                ilRet = gParseItem(slStr, 3, ",", slOutStr)
                ilRet = gParseItem(slOutStr, 2, "=", slOutStr)
                ilShttCode = Val(slOutStr)
            End If
            For llCol = DATEINDEX To ADVTCOMPLIANTINDEX Step 1
                If llCol = DATEINDEX Then
                    If grdCounts.TextMatrix(llRow, llCol) <> "" Then
                        slDate = grdCounts.TextMatrix(llRow, llCol)
                    End If
                    slStr = slDate
                Else
                    If llCol = VEHICLEINDEX Then
                        If grdCounts.TextMatrix(llRow, llCol) <> "" Then
                            llBaseRow = llRow
                        End If
                    End If
                    If llCol = VEHICLEINDEX Or llCol = STATIONINDEX Then
                        slStr = slStr & "," & """" & grdCounts.TextMatrix(llBaseRow, llCol) & """"
                    ElseIf (llCol = MULTICASTINDEX) Or (llCol = POSTMETHODINDEX) Or (llCol = AIRPLAYSINDEX) Or (llCol = NETWORKBREAKINDEX) Or (llCol = FEEDBREAKINDEX) Then
                        slStr = slStr & "," & grdCounts.TextMatrix(llBaseRow, llCol)
                    Else
                        slStr = slStr & "," & grdCounts.TextMatrix(llRow, llCol)
                    End If
                End If
            Next llCol
            ilTimeAdj = gGetTimeAdj(ilShttCode, ilVefCode, slFed, slZone)
            If slFed <> "*" Then
                If slFed <> "" Then
                    slZone = slFed & "ST"
                Else
                    slZone = ""
                End If
            End If
            If grdCounts.TextMatrix(llBaseRow, CODESINDEX) <> "" Then
            ilRet = gBinarySearchShtt(ilShttCode)
                If ilRet >= 0 Then
                    slStr = slStr & "," & grdCounts.TextMatrix(llBaseRow, CODESINDEX) & " Zone: " & UCase$(Trim$(tgShttInfo1(ilRet).shttTimeZone)) & " " & ilTimeAdj
                Else
                    slStr = slStr & "," & grdCounts.TextMatrix(llBaseRow, CODESINDEX)
                End If
            Else
                slStr = slStr & "," & "," & "," & ","
            End If
            Print #hlTo, slStr
        End If
    Next llRow
    Close #hlTo
    gSetMousePointer grdCounts, grdCounts, vbDefault
    MsgBox "Export file " & slToFile & " generated", vbOKOnly, "Export"

End Sub

Private Sub imcKey_Click()
    txtKey.Font = "Courier New"

    txtKey.Visible = Not txtKey.Visible
    If txtKey.Visible Then
        lacKey.Caption = "Click key to hide"
    Else
        lacKey.Caption = "Click key to view"
    End If
    txtKey.ZOrder
    DoEvents
End Sub


Private Sub lacKey_Click()
    mClearSelection
End Sub

Private Sub lacNote_Click()
    mClearSelection
End Sub

Private Sub tmcFillGrid_Timer()
    tmcFillGrid.Enabled = False
    gSetMousePointer grdCounts, grdCounts, vbHourglass
    mPopulate
    gSetMousePointer grdCounts, grdCounts, vbDefault
End Sub


Private Function mGetNetworkCount()
    Dim slFed As String
    Dim ilTimeAdj As Integer
    Dim slZone As String
    
    If bmRetrieveLst Then
        bmRetrieveLst = False
'        smSQLQuery = "Select Count(1) from lst where lstLogVefCode = " & imVefCode
'        smSQLQuery = smSQLQuery & " And lstType = 0"
'        smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
'        smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
'        smSQLQuery = smSQLQuery & " And lstLogDate >= '" & smSQLStartDate & "'"
'        smSQLQuery = smSQLQuery & " And lstLogDate <= '" & smSQLEndDate & "'"
'        smSQLQuery = smSQLQuery & " And lstStatus In (0, 1, 9, 10)"
'        Set rst_Lst = gSQLSelectCall(smSQLQuery)
'        If Not rst_Lst.EOF Then
'            smAirLstCount = rst_Lst(0).Value
'        Else
'            smAirLstCount = "0"
'        End If
'        smSQLQuery = "Select Count(1) from lst where lstLogVefCode = " & imVefCode
'        smSQLQuery = smSQLQuery & " And lstType = 0"
'        smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
'        smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
'        smSQLQuery = smSQLQuery & " And lstLogDate >= '" & smSQLStartDate & "'"
'        smSQLQuery = smSQLQuery & " And lstLogDate <= '" & smSQLEndDate & "'"
'        smSQLQuery = smSQLQuery & " And lstStatus In (2, 3, 4, 5, 8)"
'        Set rst_Lst = gSQLSelectCall(smSQLQuery)
'        If Not rst_Lst.EOF Then
'            If rst_Lst(0).Value > 0 Then
'                smMissedLstCount = rst_Lst(0).Value
'            Else
'                smMissedLstCount = ""
'            End If
'        Else
'            smMissedLstCount = ""
'        End If
'        smSQLQuery = "Select Count(1) from lst where lstLogVefCode = " & imVefCode
'        smSQLQuery = smSQLQuery & " And lstType = 0"
'        smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
'        smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
'        smSQLQuery = smSQLQuery & " And lstLogDate >= '" & smSQLStartDate & "'"
'        smSQLQuery = smSQLQuery & " And lstLogDate <= '" & smSQLEndDate & "'"
'        smSQLQuery = smSQLQuery & " And lstStatus = 11"
'        Set rst_Lst = gSQLSelectCall(smSQLQuery)
'        If Not rst_Lst.EOF Then
'            If rst_Lst(0).Value > 0 Then
'                smMGLstCount = rst_Lst(0).Value
'            Else
'                smMGLstCount = ""
'            End If
'        Else
'            smMGLstCount = ""
'        End If
        ilTimeAdj = gGetTimeAdj(imShttCode, imVefCode, slFed, slZone)
        If slFed <> "*" Then
            If slFed <> "" Then
                slZone = slFed & "ST"
            Else
                slZone = ""
            End If
        End If
        smSQLQuery = "Select Count(1) from lst where lstLogVefCode = " & imVefCode
        smSQLQuery = smSQLQuery & " And lstType <> 1"
        smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
        smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
        smSQLQuery = smSQLQuery & " And lstLogDate >= '" & smSQLStartDate & "'"
        smSQLQuery = smSQLQuery & " And lstLogDate <= '" & smSQLEndDate & "'"
        If slZone <> "" Then
            smSQLQuery = smSQLQuery & " And SubString(lstZone, 1, 1) = '" & Left$(slZone, 1) & "'"
        End If
        smSQLQuery = smSQLQuery & " And lstStatus In (0, 1, 9, 10)"
        Set rst_Lst = gSQLSelectCall(smSQLQuery)
        If Not rst_Lst.EOF Then
            smNetworkCount = rst_Lst(0).Value
        Else
            smNetworkCount = "0"
        End If
    End If
    mGetNetworkCount = smNetworkCount
'    grdCounts.TextMatrix(llRow, NETWORKINDEX) = Val(smAirLstCount) + Val(smMissedLstCount)
'    grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = smAirLstCount
'    grdCounts.TextMatrix(llRow, MGINDEX) = smMissedLstCount
'    grdCounts.TextMatrix(llRow, MGINMISSININDEX) = smMGLstCount
End Function

Private Function mDatExist() As Boolean
    smSQLQuery = "Select Count(1) From dat Where datAtfCode = " & lmAttCode
    
    Set rst_DAT = gSQLSelectCall(smSQLQuery)
    If Not rst_DAT.EOF Then
        If rst_DAT(0).Value > 0 Then
            mDatExist = True
        End If
    Else
        mDatExist = False
    End If

End Function

Private Function mGetPledgeSpotCount(blDatExist As Boolean) As String
    smSQLQuery = "Select Count(1) From ast "
    If (blDatExist) Then
        smSQLQuery = smSQLQuery + "Left Outer Join dat On astDatCode = datCode"
    End If
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    If (blDatExist) Then
        'Only 0, 1, 9 and 10 allowed with pledge
        smSQLQuery = smSQLQuery + " And astStatus <> 11"    'Ignore MG's
        smSQLQuery = smSQLQuery + " And datFdStatus In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10)"
    Else
        smSQLQuery = smSQLQuery + " And astStatus In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 14)"
    End If
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        mGetPledgeSpotCount = rst_Ast(0).Value
    Else
        mGetPledgeSpotCount = "0"
    End If
End Function

Private Function mGetPledgeNotCarried() As String
    Dim llNCWithDat As Long
    Dim llNCWODat As Long
    llNCWithDat = 0
    llNCWODat = 0
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery + "Left Outer Join dat On astDatCode = datCode"
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    'Only 4 and 8 allowed with pledge
    smSQLQuery = smSQLQuery & " And datFdStatus = 8"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        llNCWithDat = rst_Ast(0).Value
    End If
    'Catch the break not defined with Dat
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery + "Left Outer Join dat On astDatCode = datCode"
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astStatus = 8"
    smSQLQuery = smSQLQuery & " And DatCode is NUll"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        llNCWODat = rst_Ast(0).Value
    End If
    If llNCWithDat > 0 Or llNCWODat > 0 Then
        mGetPledgeNotCarried = llNCWithDat + llNCWODat
    Else
        mGetPledgeNotCarried = "0"
    End If
End Function

Private Function mGetPostedNotCarried() As Long
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus <> 0"
    'Only 4 and 8 allowed with pledge
    smSQLQuery = smSQLQuery + " And astStatus = 8"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetPostedNotCarried = rst_Ast(0).Value
        Else
            mGetPostedNotCarried = 0
        End If
    Else
        mGetPostedNotCarried = 0
    End If

End Function
Private Function mGetUnpostedNotCarried() As Long
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus = 0"
    'Only 4 and 8 allowed with pledge
    smSQLQuery = smSQLQuery + " And astStatus = 8"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetUnpostedNotCarried = rst_Ast(0).Value
        Else
            mGetUnpostedNotCarried = 0
        End If
    Else
        mGetUnpostedNotCarried = 0
    End If

End Function
Private Function mAffiliateSpotExist() As Boolean
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mAffiliateSpotExist = True
        Else
            mAffiliateSpotExist = False
        End If
    Else
        mAffiliateSpotExist = False
    End If
End Function
Private Function mSpotsExportedToWeb() As Boolean

    smSQLQuery = "Select Count(Distinct astFeedDate) From ast"
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And Mod(astStatus, 100) In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    smSQLQuery = "Select Count(Distinct entFeedDate) From ent"
    smSQLQuery = smSQLQuery & " Where entAttCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And entFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And entFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And entType = 'S'"
    Set rst_ent = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) And (Not rst_ent.EOF) Then
        If rst_ent(0).Value > 0 Then
            If rst_Ast(0).Value <> rst_ent(0).Value Then
                mSpotsExportedToWeb = False
            Else
                mSpotsExportedToWeb = True
            End If
        Else
            mSpotsExportedToWeb = False
        End If
    Else
        mSpotsExportedToWeb = False
    End If
End Function
Private Function mGetUnpostedSpotCount() As Long
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus = 0"
    smSQLQuery = smSQLQuery & " And Mod(astStatus, 100) In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetUnpostedSpotCount = rst_Ast(0).Value
        Else
            mGetUnpostedSpotCount = 0
        End If
    Else
        mGetUnpostedSpotCount = 0
    End If
End Function
Private Function mGetPostedSpotCount() As Long
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus <> 0"
    smSQLQuery = smSQLQuery & " And Mod(astStatus, 100) In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetPostedSpotCount = rst_Ast(0).Value
        Else
            mGetPostedSpotCount = 0
        End If
    Else
        mGetPostedSpotCount = 0
    End If
End Function

Private Function mGetUnpostedMissedCount() As String
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus = 0"
    smSQLQuery = smSQLQuery & " And Mod(astStatus, 100) In (2, 3, 4, 5, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetUnpostedMissedCount = rst_Ast(0).Value
        Else
            mGetUnpostedMissedCount = ""
        End If
    Else
        mGetUnpostedMissedCount = ""
    End If
End Function
Private Function mGetPostedMissedCount() As String
    smSQLQuery = "Select Count(1) From ast "
    smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And astCPStatus <> 0"
    smSQLQuery = smSQLQuery & " And Mod(astStatus, 100) In (2, 3, 4, 5, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetPostedMissedCount = rst_Ast(0).Value
        Else
            mGetPostedMissedCount = ""
        End If
    Else
        mGetPostedMissedCount = ""
    End If
End Function
Private Function mGetMGInMissedInCount() As Integer
    smSQLQuery = "Select Count(1) From ast As A"
    smSQLQuery = smSQLQuery & " Left Outer Join ast As B On A.astLkAstCode = B.astCode"
    smSQLQuery = smSQLQuery & " Where A.astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And A.astAirDate >= '" & smSQLMoDate & "'"
    smSQLQuery = smSQLQuery & " And A.astAirDate <= '" & smSQLSuDate & "'"
    smSQLQuery = smSQLQuery & " And Mod(A.astStatus, 100) = 11"
    smSQLQuery = smSQLQuery & " And B.astFeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And B.astFeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And Mod(B.astStatus, 100) In (2, 3, 4, 5, 14)"
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetMGInMissedInCount = rst_Ast(0).Value
        Else
            mGetMGInMissedInCount = 0
        End If
    Else
        mGetMGInMissedInCount = 0
    End If
End Function
Private Function mGetMGInMissedOutCount() As Integer
    smSQLQuery = "Select Count(1) From ast As A"
    smSQLQuery = smSQLQuery & " Left Outer Join ast As B On A.astLkAstCode = B.astCode"
    smSQLQuery = smSQLQuery & " Where A.astAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And A.astAirDate >= '" & smSQLMoDate & "'"
    smSQLQuery = smSQLQuery & " And A.astAirDate <= '" & smSQLSuDate & "'"
    smSQLQuery = smSQLQuery & " And Mod(A.astStatus, 100) = 11"
    smSQLQuery = smSQLQuery & " And Mod(B.astStatus, 100) In (2, 3, 4, 5, 14)"
    smSQLQuery = smSQLQuery & " And (B.astFeedDate < '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " Or B.astFeedDate > '" & smSQLEndDate & "'" & ")"
    
    Set rst_Ast = gSQLSelectCall(smSQLQuery)
    If (Not rst_Ast.EOF) Then
        If rst_Ast(0).Value > 0 Then
            mGetMGInMissedOutCount = rst_Ast(0).Value
        Else
            mGetMGInMissedOutCount = 0
        End If
    Else
        mGetMGInMissedOutCount = 0
    End If
End Function
Private Sub mCompliance(llRow As Long, blNotCompliant As Boolean)
    
    smSQLQuery = "Select cpttNoSpotsGen, cpttAgyCompliant, cpttNoCompliant, cpttPostingStatus From cptt Where cpttAtfCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And cpttStartDate = '" & smSQLMoDate & "'"
    smSQLQuery = smSQLQuery & " And cpttStatus = 1"
    Set rst_Cptt = gSQLSelectCall(smSQLQuery)
    If (Not rst_Cptt.EOF) Then
        If (rst_Cptt!cpttNoSpotsGen <> 0) Or (rst_Cptt!cpttAgyCompliant <> 0) Or (rst_Cptt!cpttNoCompliant <> 0) Then
            grdCounts.TextMatrix(llRow, AGYCOMPLIANTINDEX) = rst_Cptt!cpttAgyCompliant
            grdCounts.TextMatrix(llRow, ADVTCOMPLIANTINDEX) = rst_Cptt!cpttNoCompliant
            'Set agency color
            If rst_Cptt!cpttNoSpotsGen <> rst_Cptt!cpttAgyCompliant Then
                grdCounts.Col = AGYCOMPLIANTINDEX
                grdCounts.CellForeColor = vbRed
                blNotCompliant = True
            End If
            
            'Set station color
            If rst_Cptt!cpttNoSpotsGen <> rst_Cptt!cpttNoCompliant Then
                grdCounts.Col = ADVTCOMPLIANTINDEX
                grdCounts.CellForeColor = vbRed
                blNotCompliant = True
            End If

        End If
    End If

End Sub

Private Sub mBuildVendorID()
    Dim llDate As Long
    ReDim imVendorID(0 To 0) As Integer
    
    smSQLQuery = "Select vatWvtVendorID From vat_Vendor_Agreement Where vatAttCode = " & lmAttCode
    Set rst_vat = gSQLSelectCall(smSQLQuery)
    Do While Not rst_vat.EOF
        If gIsWebVendor(rst_vat!vatwvtvendorid) Then
            llDate = mDetermineVendorStartDate(rst_vat!vatwvtvendorid)
            If llDate <= lmStartDate Then
                imVendorID(UBound(imVendorID)) = rst_vat!vatwvtvendorid
                ReDim Preserve imVendorID(0 To UBound(imVendorID) + 1) As Integer
            End If
        End If
        rst_vat.MoveNext
    Loop
End Sub
Private Function mDetermineVendorStartDate(ilVendorId As Integer) As Long
    Dim slDataArray() As String
    Dim llTotalRecords As Long

    smSQLQuery = "Select Min(SpotsDate) As MinDate From WebVendorCountArchive "
    smSQLQuery = smSQLQuery & " Where AttCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And VendorIdCode = " & ilVendorId
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        mDetermineVendorStartDate = 99999999
    Else
        mDetermineVendorStartDate = gDateValue(Format(gGetDataNoQuotes(slDataArray(1)), sgShowDateForm))
    End If
End Function

Private Function mGetVendorExportCount(ilVendorId As Integer) As Integer
    Dim slDataArray() As String
    Dim llTotalRecords As Long

    smSQLQuery = "Select Sum(SpotsCount) As SumSpots From WebVendorCountArchive "
    smSQLQuery = smSQLQuery & " Where AttCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And VendorIdCode = " & ilVendorId
    smSQLQuery = smSQLQuery & " And SpotsDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And SpotsDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And ExportOrImport = 'E'"
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        mGetVendorExportCount = 0
    Else
        mGetVendorExportCount = Val(gGetDataNoQuotes(slDataArray(1)))
    End If
    
End Function

Private Function mGetVendorImportCount(ilVendorId As Integer) As Integer
    Dim slDataArray() As String
    Dim llTotalRecords As Long

    smSQLQuery = "Select Sum(SpotsCount) As SumSpots From WebVendorCountArchive "
    smSQLQuery = smSQLQuery & " Where AttCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And VendorIdCode = " & ilVendorId
    smSQLQuery = smSQLQuery & " And SpotsDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And SpotsDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And ExportOrImport = 'I'"
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        mGetVendorImportCount = 0
    Else
        mGetVendorImportCount = Val(gGetDataNoQuotes(slDataArray(1)))
    End If
    
End Function
Private Function mGetWebUnpostedSpotCount() As Long
    Dim llCount As Long
    Dim slStr As String
    Dim slDataArray() As String
    Dim llTotalRecords As Long
    
    smSQLQuery = "Select Count(1) From Spots "
    smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And PostedFlag = 0"
    ''smSQLQuery = smSQLQuery & " And (RecType = ''"
    ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%M%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%R%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%B%')"
    'The % followed by the escape character 25 works but will use the substring instead
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
    smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        slStr = ""
    Else
        slStr = gGetDataNoQuotes(slDataArray(1))
    End If
    If ((slStr = "") Or (slStr = "0")) Then
        smSQLQuery = "Select Count(1) From Spot_History "
        smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
        smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
        smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
        smSQLQuery = smSQLQuery & " And PostedFlag = 0"
        ''smSQLQuery = smSQLQuery & " And (RecType = ''"
        ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
        'smSQLQuery = smSQLQuery & " And (RecType Not Like '%M%'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%R%'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%B%')"
        'The % followed by the escape character 25 works but will use the substring instead
        'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
        smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"

        llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
        If llTotalRecords < 1 Then
            slStr = ""
        Else
            slStr = gGetDataNoQuotes(slDataArray(1))
        End If
    End If
    If slStr = "" Then
        llCount = 0
    Else
        llCount = Val(slStr)
    End If
    mGetWebUnpostedSpotCount = llCount
End Function

Private Function mGetWebByVendorSpotCount(ilVendorId As Integer) As String
    Dim llCount As Long
    Dim slStr As String
    Dim slDataArray() As String
    Dim llTotalRecords As Long
    
    smSQLQuery = "Select Count(1) From Spots "
    smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And Source = '" & gVendorInitials(ilVendorId) & "'"
    ''smSQLQuery = smSQLQuery & " And (RecType = ''"
    ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%M%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%R%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%B%')"
    'The % followed by the escape character 25 works but will use the substring instead
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
    smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
    
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        slStr = ""
    Else
        slStr = gGetDataNoQuotes(slDataArray(1))
    End If
    If ((slStr = "") Or (slStr = "0")) Then
        smSQLQuery = "Select Count(1) From Spot_History "
        smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
        smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
        smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
        smSQLQuery = smSQLQuery & " And Source = '" & gVendorInitials(ilVendorId) & "'"
        ''smSQLQuery = smSQLQuery & " And (RecType = ''"
        ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
        'smSQLQuery = smSQLQuery & " And (RecType Not Like '%M%'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%R%'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%B%')"
        'The % followed by the escape character 25 works but will use the substring instead
        'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
        smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
        
        llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
        If llTotalRecords < 1 Then
            slStr = ""
        Else
            slStr = gGetDataNoQuotes(slDataArray(1))
        End If
    End If
    If slStr = "0" Then
        slStr = ""
    End If
    mGetWebByVendorSpotCount = slStr
End Function
Private Function mGetWebPostedSpotCount() As Long
    Dim llCount As Long
    Dim slStr As String
    Dim slDataArray() As String
    Dim llTotalRecords As Long
    
    smSQLQuery = "Select Count(1) From Spots "
    smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And PostedFlag = 1"
    ''smSQLQuery = smSQLQuery & " And (RecType = ''"
    ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
    ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%M%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%R%'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%B%')"
    'The % followed by the escape character 25 works but will use the substring instead
    'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
    'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
    smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
    
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        slStr = ""
    Else
        slStr = gGetDataNoQuotes(slDataArray(1))
    End If
    If ((slStr = "") Or (slStr = "0")) Then
        smSQLQuery = "Select Count(1) From Spot_History "
        smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
        smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
        smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
        smSQLQuery = smSQLQuery & " And PostedFlag = 1"
        ''smSQLQuery = smSQLQuery & " And (RecType = ''"
        ''smSQLQuery = smSQLQuery & " Or RecType = '0'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'X'"
        ''smSQLQuery = smSQLQuery & " Or RecType = 'G')"
        'The % followed by the escape character 25 works but will use the substring instead
        'smSQLQuery = smSQLQuery & " And (RecType Not Like '%25M%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25R%25'"
        'smSQLQuery = smSQLQuery & " And RecType Not Like '%25B%'25)"
        smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
        
        llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
        If llTotalRecords < 1 Then
            slStr = ""
        Else
            slStr = gGetDataNoQuotes(slDataArray(1))
        End If
    End If
    If slStr = "" Then
        llCount = 0
    Else
        llCount = Val(slStr)
    End If
    mGetWebPostedSpotCount = llCount
End Function


Private Sub mLoadKeyText()

    txtKey.Text = "Vehicle: Airing vehicle name" + sgCRLF
    
    txtKey.Text = txtKey.Text + "Station: Affiliate Call Letters" + sgCRLF
    
    txtKey.Text = txtKey.Text + "Multicast: * and ** indicate that station is defined as multicast within an agreement" + sgCRLF
    txtKey.Text = txtKey.Text + "        ** indicated that this staton is defined as the master station" + sgCRLF
    
    txtKey.Text = txtKey.Text + "Post Method: Method affiliate spots air date/time specified- Affiliate (Manual); Web or Vendor" + sgCRLF
    
    txtKey.Text = txtKey.Text + "Flow :Indicates the flow of spots between the three systems: Affiliate; Web and Vendor Traffic" + modAffiliate.sgCRLF
    txtKey.Text = txtKey.Text + "   T->L Network spots. " + sgCRLF
    txtKey.Text = txtKey.Text + "   T->A Affiliate spots. " + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Affiliate spots missing." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Either export or run spot report" + sgCRLF
    txtKey.Text = txtKey.Text & "                or view week via Affiliate Affidavit" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->M Affiliate spots exist, manually posting" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W Affiliate spots exist. " + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Web spots missing." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Generate Web export" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W->ID Sent Web spots to Vendor ID" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W<->ID Sent Web spots to Vendor and received spots from Vendor ID." + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W<-ID Received spot from Vendor ID" + sgCRLF
    txtKey.Text = txtKey.Text + "   W->A Spots sent from Web to Affiliate" + sgCRLF

    txtKey.Text = txtKey.Text + "Network Count: Traffic Log spots transferred to Affiliate system" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Network spots missing." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Generate Traffic log" + sgCRLF

    txtKey.Text = txtKey.Text + "Air Plays: Number of times the same spot might air" + sgCRLF

    txtKey.Text = txtKey.Text + "Feed Spots: Network spots with the Pledge rules (Aired) applied" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Feed and Pledge out of balance." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Check Pledge and Feed counts" + sgCRLF
    txtKey.Text = txtKey.Text + "Feed NotC: Network spots with the Pledge rules (Not Carried) applied" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Feed and Pledge out of balance." + sgCRLF

    txtKey.Text = txtKey.Text + "Pledge Spots: Affiliate spots with the Pledge rules (Aired) applied" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Pledge and Feed out of balance" + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Check Pledge and Feed counts" + sgCRLF
    txtKey.Text = txtKey.Text + "Pledge NotC: Affiliate spots with the Pledge rules (Not Carried) applied" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Pledge and Feed out of balance" + sgCRLF

    txtKey.Text = txtKey.Text + "Spot Count: Number of Affiliate or Web system spots (Aired and Missed)" + sgCRLF
    txtKey.Text = txtKey.Text + "   T->A Affiliate system spots (Aired and missed)." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Pledge and Affiliate spots out of balance" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W Web system spots (Aired and missed)." + sgCRLF
    txtKey.Text = txtKey.Text + "   A->M Affiliate system spots Posted and Unposted (Aired and Missed)." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Affiliate spots (Posted and Unposted) not equal to Pledge Spots." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Posted spots exist on a non posted week." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Mixture of posted and unposted spots exist on a posted week." + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W Web system spots Posted and Unposted (Aired and missed)" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Web spots and Pledge spots out of balance." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Check spot counts and/or Export spots." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Web spots missing." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Export spots." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Mixture of posted and unposted spots exist on a posted week." + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W->ID or A->W<->ID Affiliate spots on the Web system (Aired and Missed) for Vendor ID" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Web spots and Pledge spots out of balance." + sgCRLF
    txtKey.Text = txtKey.Text + "        To Fix: Check spot counts and/or Export spots." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Mixture of posted and unposted spots exist on a posted week." + sgCRLF
    txtKey.Text = txtKey.Text + "   W->A Affiliate system spots Posted and Unposted (Aired and missed)" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Mixture of posted and unposted spots exist on a posted week." + sgCRLF

    txtKey.Text = txtKey.Text + "Spot NotC: Number of not carried affiliate system spots" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: see Spot Count above." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Unposted spots exist on a posted week." + sgCRLF
    txtKey.Text = txtKey.Text + "        If Magenta: Mixture of posted and unposted spots exist on a posted week." + sgCRLF

    txtKey.Text = txtKey.Text + "MG Count: Number of spots defined as MG's within the specified dates" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Orange: Mixture of 'MG In/Missed In' and 'MG In/Missed Out'." + sgCRLF


    txtKey.Text = txtKey.Text + "Vendor Export: Number of Web spots sent to Vendor" + sgCRLF
    txtKey.Text = txtKey.Text + "   A->W->ID or A->W<->ID Affiliate spots on the Web system(Aired plus missed) for Vendor ID" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Web spots and Vendor Export Spots out of balance." + sgCRLF
    txtKey.Text = txtKey.Text + "Vendor Import: Number of Vendor spots sent back to Web" + sgCRLF
    txtKey.Text = txtKey.Text + "Vendor Applied: Number of returned Vendor spots matched-up with Web spots" + sgCRLF
    txtKey.Text = txtKey.Text + "Compliance Agency: Number of spots that are contract compliant" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Not 100% compliant." + sgCRLF
    txtKey.Text = txtKey.Text + "Compliance Station: Number of spots that are pledge compliant" + sgCRLF
    txtKey.Text = txtKey.Text + "        If Red: Not 100% compliant." + sgCRLF
    
End Sub



Private Sub mMulticast(llRow As Long)
    Dim llGroupID As Long
    Dim slStationList As String
    Dim blMasterFound As Boolean
    Dim rstATT As ADODB.Recordset
    Dim rstShtt As ADODB.Recordset
    slStationList = ""
    blMasterFound = False
    If gIsMulticast(imShttCode) Then
        If rst_att!attMulticast = "Y" Then
            llGroupID = gGetStaMulticastGroupID(imShttCode)
            If llGroupID > 0 Then
                smSQLQuery = "Select shttCode, shttCallLetters, shttMasterCluster FROM shtt where shttMultiCastGroupID = " & llGroupID
                Set rstShtt = gSQLSelectCall(smSQLQuery)
                Do While Not rstShtt.EOF
                    If rstShtt!shttCode = imShttCode Then
                        If rstShtt!shttMasterCluster = "Y" Then
                            grdCounts.TextMatrix(llRow, MULTICASTINDEX) = "**"
                            blMasterFound = True
                        Else
                            grdCounts.TextMatrix(llRow, MULTICASTINDEX) = "*"
                        End If
                    Else
                        smSQLQuery = "Select attCode, attMulticast from att"
                        smSQLQuery = smSQLQuery + " Where attVefCode = " & imVefCode & " And attShfCode = " & rstShtt!shttCode
                        smSQLQuery = smSQLQuery + " And attOnAir <= '" & smSQLStartDate & "'"
                        smSQLQuery = smSQLQuery + " And attOffAir >= '" & smSQLStartDate & "'"
                        smSQLQuery = smSQLQuery + " And attDropDate >= '" & smSQLStartDate & "'"
                        smSQLQuery = smSQLQuery + " And attDropDate >= '" & smSQLStartDate & "'"
                        smSQLQuery = smSQLQuery + " And attMulticast = 'Y'"
                        Set rstATT = gSQLSelectCall(smSQLQuery)
                        If Not rstATT.EOF Then
                            If slStationList = "" Then
                                slStationList = Trim$(rstShtt!shttCallLetters)
                            Else
                                slStationList = slStationList & ", " & Trim$(rstShtt!shttCallLetters)
                            End If
                            If rstShtt!shttMasterCluster = "Y" Then
                                slStationList = slStationList & "**"
                                blMasterFound = True
                            End If
                        End If
                    End If
                    rstShtt.MoveNext
                Loop
                If slStationList <> "" Then
                    If Not blMasterFound Then
                        slStationList = slStationList & " No Master Station Defined"
                        If grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Vendor" Then
                            grdCounts.Col = MULTICASTSTATIONINDEX
                            grdCounts.CellForeColor = vbRed
                        End If
                    End If
                    grdCounts.TextMatrix(llRow, MULTICASTSTATIONINDEX) = slStationList
                End If
            End If
        End If
    End If
End Sub

Private Sub mShowAgreement(llRow As Long, llAgreementStartRow As Long, blOutBalance As Boolean, blNotCompliant As Boolean, blPartiallyPosted As Boolean, blUserPosting As Boolean, blBreakOutBalance As Boolean)
    Dim blShowAgreement As Boolean
    Dim llClearRow As Long
    Dim llClearCol As Long
    Dim slFed As String
    Dim ilTimeAdj As Integer
    Dim slZone As String
    Dim llCol As Long
    Dim ilRet As Integer
    
    If llAgreementStartRow = -1 Then
        Exit Sub
    End If
    If (blOutBalance And bmOutBalance) Or (blNotCompliant And bmNotCompliant) Or (blPartiallyPosted And bmPartiallyPosted) Or (blUserPosting And bmUserPosting) Or (blBreakOutBalance And bmBreakOutBalance) Then
        lmTotalErrorCount = lmTotalErrorCount + 1
    Else
        lmTotalOkCount = lmTotalOkCount + 1
    End If
    blShowAgreement = False
    If blOutBalance And bmOutBalance Then blShowAgreement = True
    If blNotCompliant And bmNotCompliant Then blShowAgreement = True
    If blPartiallyPosted And bmPartiallyPosted Then blShowAgreement = True
    If blUserPosting And bmUserPosting Then blShowAgreement = True
    If blBreakOutBalance And bmBreakOutBalance Then blShowAgreement = True
    If Not blShowAgreement And bmInBalance And Not blOutBalance And Not blNotCompliant And Not blPartiallyPosted And Not blUserPosting And Not blBreakOutBalance Then blShowAgreement = True
    If blShowAgreement Then
        If bmIncludeCodes Then
            llRow = llRow + 1
            If llRow >= grdCounts.Rows Then
                grdCounts.AddItem ""
            End If
            grdCounts.Row = llRow
            For llCol = DATEINDEX To ADVTCOMPLIANTINDEX Step 1
                grdCounts.Col = llCol
                grdCounts.CellBackColor = lmRowColor
            Next llCol
            
            grdCounts.TextMatrix(llRow, VEHICLEINDEX) = "Att: " & lmAttCode & " Vef: " & imVefCode & " Shtt: " & imShttCode
            ilTimeAdj = gGetTimeAdj(imShttCode, imVefCode, slFed, slZone)
            If slFed <> "*" Then
                If slFed <> "" Then
                    slZone = slFed & "ST"
                Else
                    slZone = ""
                End If
            End If
            ilRet = gBinarySearchShtt(imShttCode)
            If ilRet >= 0 Then
                grdCounts.TextMatrix(llRow, STATIONINDEX) = UCase$(Trim$(tgShttInfo1(ilRet).shttTimeZone)) & " " & ilTimeAdj
            End If
            
        End If
        llRow = llRow + 1
    Else
        For llClearRow = llAgreementStartRow To llRow Step 1
            grdCounts.Row = llClearRow
            For llClearCol = DATEINDEX To SORTINDEX Step 1
                grdCounts.Col = llClearCol
                grdCounts.CellBackColor = vbWhite
                grdCounts.CellForeColor = vbBlack
                grdCounts.TextMatrix(llClearRow, llClearCol) = ""
            Next llClearCol
        Next llClearRow
        llRow = llAgreementStartRow
        lmRowColor = Switch(lmRowColor = vbWhite, LIGHTGRAY, lmRowColor = LIGHTGRAY, vbWhite)
    End If
    blOutBalance = False
    blNotCompliant = False
    blPartiallyPosted = False
    blUserPosting = False
    blBreakOutBalance = False
    llAgreementStartRow = -1
End Sub

Private Sub mGetFeedSpotCount(llRow As Long, blDatExist As Boolean)

    'Replace the above with using Lst to find Dat
        
    Dim slLogDate As String
    Dim slLogTime As String
    Dim llDat As Long
    Dim llFeedCount As Long
    Dim llNotFeedCount As Long
    Dim ilDay As Integer
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slFed As String
    Dim ilTimeAdj As Integer
    Dim llLogDate As Long
    Dim ilFdDay As Integer
    Dim blDatFound As Boolean
    Dim slZone As String
    Dim ilAdjDay As String
    
    llFeedCount = 0
    llNotFeedCount = 0
    If Not blDatExist Then
        grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = grdCounts.TextMatrix(llRow, NETWORKINDEX)
        grdCounts.TextMatrix(llRow, FEEDNCINDEX) = 0
        Exit Sub
    End If
    ilTimeAdj = gGetTimeAdj(imShttCode, imVefCode, slFed, slZone)
    If slFed <> "*" Then
        If slFed <> "" Then
            slZone = slFed & "ST"
        Else
            slZone = ""
        End If
    End If
    ReDim tmDat(0 To 30) As DATRST
    llDat = 0
    smSQLQuery = "SELECT * "
    smSQLQuery = smSQLQuery + " FROM dat"
    smSQLQuery = smSQLQuery + " WHERE (datatfCode= " & lmAttCode & ")"
    Set rst_DAT = gSQLSelectCall(smSQLQuery)
    Do While Not rst_DAT.EOF
        gCreateUDTForDat rst_DAT, tmDat(llDat)
        slFdStTime = tmDat(llDat).sFdStTime
        slFdEdTime = tmDat(llDat).sFdEdTime
        If gTimeToLong(slFdEdTime, True) = 86400 Then
            slFdEdTime = "12:59:59AM"
        Else
            If gTimeToLong(slFdStTime, False) <> gTimeToLong(slFdEdTime, True) Then
                slFdEdTime = DateAdd("s", -1, slFdEdTime)
            End If
        End If
        tmDat(llDat).sFdEdTime = slFdEdTime
        llDat = llDat + 1
        If llDat = UBound(tmDat) Then
            ReDim Preserve tmDat(0 To UBound(tmDat) + 30) As DATRST
        End If
        rst_DAT.MoveNext
    Loop
    ReDim Preserve tmDat(0 To llDat) As DATRST
    
    
    ilAdjDay = 0
    If ilTimeAdj <> 0 Then
        ilDay = gWeekDayLong(lmWkEndDate)
        If ilDay = 6 Then
            smSQLQuery = "Select Count(1) From ast "
            smSQLQuery = smSQLQuery & " Where astAtfCode = " & lmAttCode
            smSQLQuery = smSQLQuery & " And astFeedDate = '" & smSQLEndDate & "'"
            If ilTimeAdj = -1 Then
                smSQLQuery = smSQLQuery & " And astFeedTime >= '" & Format$("11:00:00PM", sgSQLTimeForm) & "'"
            ElseIf ilTimeAdj = -2 Then
                smSQLQuery = smSQLQuery & " And astFeedTime >= '" & Format$("10:00:00PM", sgSQLTimeForm) & "'"
            Else
                smSQLQuery = smSQLQuery & " And astFeedTime >= '" & Format$("9:00:00PM", sgSQLTimeForm) & "'"
            End If
            smSQLQuery = smSQLQuery & " And astFeedTime <= '" & Format$("11:59:59PM", sgSQLTimeForm) & "'"
            smSQLQuery = smSQLQuery + " And astStatus In (0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 14)"
            Set rst_Ast = gSQLSelectCall(smSQLQuery)
            If (Not rst_Ast.EOF) Then
                If rst_Ast(0).Value > 0 Then
                    ilAdjDay = 1
                End If
            End If
        Else
            ilAdjDay = 1
        End If
    End If
    
    smSQLQuery = "Select * from lst where lstLogVefCode = " & imVefCode
    smSQLQuery = smSQLQuery & " And lstType <> 1"
    smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
    smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
    smSQLQuery = smSQLQuery & " And lstLogDate >= '" & Format(DateAdd("d", -ilAdjDay, smSQLStartDate), sgSQLDateForm) & "'"
    smSQLQuery = smSQLQuery & " And lstLogDate <= '" & Format(DateAdd("d", ilAdjDay, smSQLEndDate), sgSQLDateForm) & "'"
    If slZone <> "" Then
        smSQLQuery = smSQLQuery & " And SubString(lstZone, 1, 1) = '" & Left$(slZone, 1) & "'"
    End If
    smSQLQuery = smSQLQuery & " And lstStatus In (0, 1, 9, 10)"
    Set rst_Lst = gSQLSelectCall(smSQLQuery)
    Do While Not rst_Lst.EOF
        'Adjust time
        slLogDate = Format(rst_Lst!lstLogDate, sgShowDateForm)
        slLogTime = Format(rst_Lst!lstLogTime, sgShowTimeWSecForm)
        gAdjustEventTime ilTimeAdj, slLogDate, slLogTime
        llLogDate = gDateValue(slLogDate)
        If (llLogDate >= lmWkStartDate) And (llLogDate <= lmWkEndDate) Then
            'Search DAT for match
            If UBound(tmDat) > LBound(tmDat) Then
                blDatFound = False
                ilDay = gWeekDayLong(llLogDate)
                For llDat = 0 To UBound(tmDat) - 1 Step 1
                    ilFdDay = Switch(ilDay = 0, tmDat(llDat).iFdMon, ilDay = 1, tmDat(llDat).iFdTue, ilDay = 2, tmDat(llDat).iFdWed, ilDay = 3, tmDat(llDat).iFdThu, ilDay = 4, tmDat(llDat).iFdFri, ilDay = 5, tmDat(llDat).iFdSat, ilDay = 6, tmDat(llDat).iFdSun)
                    If ilFdDay > 0 Then
                        slFdStTime = tmDat(llDat).sFdStTime
                        slFdEdTime = tmDat(llDat).sFdEdTime
                        If (gTimeToLong(slLogTime, False) >= gTimeToLong(slFdStTime, False)) And (gTimeToLong(slLogTime, False) <= gTimeToLong(slFdEdTime, True)) Then
                            If tmDat(llDat).iFdStatus <> 8 Then
                                llFeedCount = llFeedCount + 1
                            Else
                                llNotFeedCount = llNotFeedCount + 1
                            End If
                            blDatFound = True
                        End If
                    End If
                Next llDat
            Else
                blDatFound = True
            End If
            If Not blDatFound Then
                llNotFeedCount = llNotFeedCount + 1
            End If
        End If
        rst_Lst.MoveNext
    Loop
    grdCounts.TextMatrix(llRow, FEEDSPOTINDEX) = Str(llFeedCount)
    grdCounts.TextMatrix(llRow, FEEDNCINDEX) = Str(llNotFeedCount)
End Sub

Private Sub mCopyColor(llFromRow As Long, llToRow As Long, ilColumnIndex As Integer)
    Dim llColor As Long
    
    grdCounts.Row = llFromRow
    grdCounts.Col = ilColumnIndex
    llColor = grdCounts.CellForeColor
    grdCounts.Row = llToRow
    grdCounts.Col = ilColumnIndex
    grdCounts.CellForeColor = llColor

End Sub

Private Sub mShowCodeRow(llRow As Long)

    If Not bmIncludeCodes Then
        Exit Sub
    End If
    
End Sub

Private Function mGetNetworkBreakCount() As String
    Dim slFed As String
    Dim ilTimeAdj As Integer
    Dim slZone As String
    Dim llCount As Long
    Dim slLogDate As String
    Dim slLogTime As String
    Dim llLogDate As Long
    Dim ilAdjDay As Integer
    Dim ilVefCombo As Integer
    Dim slSDate As String
    Dim ilShttCode As Integer
    Dim llDat As Long
    Dim ilDay As Integer

    llCount = 0
    ilAdjDay = 1
    ilTimeAdj = gGetTimeAdj(imShttCode, imVefCode, slFed, slZone)
    If slFed <> "*" Then
        If slFed <> "" Then
            slZone = slFed & "ST"
        Else
            slZone = ""
        End If
    End If
    smSQLQuery = "Select Distinct lstLogDate, lstLogTime from lst where lstLogVefCode = " & imVefCode
    smSQLQuery = smSQLQuery & " And lstType <= 1"
    smSQLQuery = smSQLQuery & " And lstBkoutLstCode = 0"
    smSQLQuery = smSQLQuery & " And lstSplitNetwork <> 'S'" 'N=Not a split spot; P=Promary; S=Secondary
    'smSQLQuery = smSQLQuery & " And lstLogDate >= '" & smSQLStartDate & "'"
    'smSQLQuery = smSQLQuery & " And lstLogDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And lstLogDate >= '" & Format(DateAdd("d", -ilAdjDay, smSQLStartDate), sgSQLDateForm) & "'"
    smSQLQuery = smSQLQuery & " And lstLogDate <= '" & Format(DateAdd("d", ilAdjDay, smSQLEndDate), sgSQLDateForm) & "'"
    If slZone <> "" Then
        smSQLQuery = smSQLQuery & " And SubString(lstZone, 1, 1) = '" & Left$(slZone, 1) & "'"
    End If
    smSQLQuery = smSQLQuery & " And lstStatus In (0, 1, 9, 10)"
    Set rst_Lst = gSQLSelectCall(smSQLQuery)
    If Not rst_Lst.EOF Then
        Do While Not rst_Lst.EOF
            'Adjust time
            slLogDate = Format(rst_Lst!lstLogDate, sgShowDateForm)
            slLogTime = Format(rst_Lst!lstLogTime, sgShowTimeWSecForm)
            gAdjustEventTime ilTimeAdj, slLogDate, slLogTime
            llLogDate = gDateValue(slLogDate)
            If (llLogDate >= lmWkStartDate) And (llLogDate <= lmWkEndDate) Then
                llCount = llCount + 1
            End If
            rst_Lst.MoveNext
        Loop
    Else
        'Get break information from the program structure
        If UBound(tgDat) <= LBound(tgDat) Then
            mGetBreaks
        End If
        For llDat = 0 To UBound(tgDat) - 1 Step 1
            For ilDay = gWeekDayLong(lmWkStartDate) To gWeekDayLong(lmWkEndDate) Step 1
                If tgDat(llDat).iFdDay(ilDay) > 0 Then
                    llCount = llCount + 1
                End If
            Next ilDay
        Next llDat
    End If
    smNetworkBreakCount = llCount
    mGetNetworkBreakCount = smNetworkBreakCount
End Function

Private Function mGetFeedBreakCount(blDatExist As Boolean) As String
    Dim llOuterLoop As Long
    Dim llInnerLoop As Long
    Dim llCount As Long
    Dim ilDay As Integer
    Dim ilFdDay As Integer
    Dim blMatchFd As Boolean
    Dim llDat As Long
    
    bmAirPlayConflict = False
    llCount = 0
    If Not blDatExist Then
        If UBound(tgDat) <= LBound(tgDat) Then
            mGetBreaks
        End If
        For llDat = 0 To UBound(tgDat) - 1 Step 1
            For ilDay = gWeekDayLong(lmWkStartDate) To gWeekDayLong(lmWkEndDate) Step 1
                If tgDat(llDat).iFdDay(ilDay) > 0 Then
                    llCount = llCount + 1
                End If
            Next ilDay
        Next llDat
    End If
    For ilDay = gWeekDayLong(lmWkStartDate) To gWeekDayLong(lmWkEndDate) Step 1
        For llOuterLoop = 0 To UBound(tmDat) - 1 Step 1
            ilFdDay = Switch(ilDay = 0, tmDat(llOuterLoop).iFdMon, ilDay = 1, tmDat(llOuterLoop).iFdTue, ilDay = 2, tmDat(llOuterLoop).iFdWed, ilDay = 3, tmDat(llOuterLoop).iFdThu, ilDay = 4, tmDat(llOuterLoop).iFdFri, ilDay = 5, tmDat(llOuterLoop).iFdSat, ilDay = 6, tmDat(llOuterLoop).iFdSun)
            If ilFdDay > 0 Then
                blMatchFd = False
                For llInnerLoop = llOuterLoop - 1 To 0 Step -1
                    ilFdDay = Switch(ilDay = 0, tmDat(llInnerLoop).iFdMon, ilDay = 1, tmDat(llInnerLoop).iFdTue, ilDay = 2, tmDat(llInnerLoop).iFdWed, ilDay = 3, tmDat(llInnerLoop).iFdThu, ilDay = 4, tmDat(llInnerLoop).iFdFri, ilDay = 5, tmDat(llInnerLoop).iFdSat, ilDay = 6, tmDat(llInnerLoop).iFdSun)
                    If ilFdDay > 0 Then
                        If gTimeToLong(tmDat(llOuterLoop).sFdStTime, False) = gTimeToLong(tmDat(llInnerLoop).sFdStTime, False) Then
                            blMatchFd = True
                            If tmDat(llOuterLoop).iAirPlayNo = tmDat(llInnerLoop).iAirPlayNo Then
                                bmAirPlayConflict = True
                            End If
                            Exit For
                        End If
                    End If
                Next llInnerLoop
                If Not blMatchFd Then
                    'Check break length and if greater then 5 minutes assume it represents a daypart
                    If gTimeToLong(tmDat(llOuterLoop).sFdEdTime, True) - gTimeToLong(tmDat(llOuterLoop).sFdStTime, False) > 300 Then
                        If UBound(tgDat) <= LBound(tgDat) Then
                            mGetBreaks
                        End If
                        For llDat = 0 To UBound(tgDat) - 1 Step 1
                            If tgDat(llDat).iFdDay(ilDay) > 0 Then
                                If (gTimeToLong(tmDat(llOuterLoop).sFdStTime, False) <= gTimeToLong(tgDat(llDat).sFdSTime, False)) And (gTimeToLong(tgDat(llDat).sFdSTime, False) < gTimeToLong(tmDat(llOuterLoop).sFdEdTime, True)) Then
                                    llCount = llCount + 1
                                End If
                            End If
                        Next llDat
                    Else
                        llCount = llCount + 1
                    End If
                End If
            End If
        Next llOuterLoop
    Next ilDay
    smFeedBreakCount = Str(llCount)
    mGetFeedBreakCount = smFeedBreakCount
End Function

Private Function mGetPledgeInfo() As Boolean
    Dim llDat As Long
    Dim slFdStTime As String
    Dim slFdEdTime As String
    
    llDat = 0
    smSQLQuery = "SELECT * "
    smSQLQuery = smSQLQuery + " FROM dat"
    smSQLQuery = smSQLQuery + " WHERE (datatfCode= " & lmAttCode & ")"
    smSQLQuery = smSQLQuery + " Order by datFdStTime"
    Set rst_DAT = gSQLSelectCall(smSQLQuery)
    If rst_DAT.EOF Then
        ReDim tmDat(0 To 0) As DATRST
        mGetPledgeInfo = False
        Exit Function
    End If
    ReDim tmDat(0 To 30) As DATRST
    Do While Not rst_DAT.EOF
        gCreateUDTForDat rst_DAT, tmDat(llDat)
        slFdStTime = tmDat(llDat).sFdStTime
        slFdEdTime = tmDat(llDat).sFdEdTime
        If gTimeToLong(slFdEdTime, True) = 86400 Then
            slFdEdTime = "12:59:59AM"
        Else
            If gTimeToLong(slFdStTime, False) <> gTimeToLong(slFdEdTime, True) Then
                slFdEdTime = DateAdd("s", -1, slFdEdTime)
            End If
        End If
        tmDat(llDat).sFdEdTime = slFdEdTime
        llDat = llDat + 1
        If llDat = UBound(tmDat) Then
            ReDim Preserve tmDat(0 To UBound(tmDat) + 30) As DATRST
        End If
        rst_DAT.MoveNext
    Loop
    ReDim Preserve tmDat(0 To llDat) As DATRST
    mGetPledgeInfo = True
End Function

Private Function mGetPostedBy(llRow As Long) As String
    Dim slStr As String
    Dim slDataArray() As String
    Dim llTotalRecords As Long
    
    smSQLQuery = "Select Count(1) From Spots "
    smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
    smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
    smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
    smSQLQuery = smSQLQuery & " And PostedFlag = 1"
    smSQLQuery = smSQLQuery & " And ExportedFlag = 1"
    smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
    smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
    
    llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
    If llTotalRecords < 1 Then
        slStr = ""
    Else
        slStr = gGetDataNoQuotes(slDataArray(1))
    End If
    If ((slStr = "") Or (slStr = "0")) Then
        smSQLQuery = "Select Count(1) From Spot_History "
        smSQLQuery = smSQLQuery & " Where attCode = " & lmAttCode
        smSQLQuery = smSQLQuery & " And FeedDate >= '" & smSQLStartDate & "'"
        smSQLQuery = smSQLQuery & " And FeedDate <= '" & smSQLEndDate & "'"
        smSQLQuery = smSQLQuery & " And PostedFlag = 1"
        smSQLQuery = smSQLQuery & " And ExportedFlag = 1"
        smSQLQuery = smSQLQuery & " And CharIndex('M', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('R', RecType) = 0"
        smSQLQuery = smSQLQuery & " And CharIndex('B', RecType) = 0"
        
        llTotalRecords = gExecWebSQLForVendor(slDataArray, smSQLQuery, True)
        If llTotalRecords < 1 Then
            slStr = ""
        Else
            slStr = gGetDataNoQuotes(slDataArray(1))
        End If
    End If
    If slStr = "" Or slStr = "0" Then
        mGetPostedBy = "A"
    Else
        If grdCounts.TextMatrix(llRow, POSTMETHODINDEX) = "Vendor" Then
            mGetPostedBy = "V"
        Else
            mGetPostedBy = "W"
        End If
    End If
    

End Function

Private Sub mGetBreaks()
    Dim ilVefCombo As Integer
    Dim slSDate As String
    Dim ilShttCode As Integer

    'Get break information from the program structure
    smSQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
    Set rst_vef = gSQLSelectCall(smSQLQuery)
    If Not rst_vef.EOF Then
        ilVefCombo = rst_vef!vefCombineVefCode
    Else
        ilVefCombo = 0
    End If
    
    slSDate = smFeedStartDate
    ilShttCode = imShttCode
    gGetAvails lmAttCode, ilShttCode, imVefCode, ilVefCombo, smFeedStartDate, True
End Sub

Private Sub mClearSelection()
    Dim llColor As Long
    If (lmCellLastRow <> -1) And (lmCellLastCol <> -1) Then
        grdCounts.Row = lmCellLastRow
        grdCounts.Col = lmCellLastCol - 1
        llColor = grdCounts.CellBackColor
        grdCounts.Col = lmCellLastCol
        grdCounts.CellBackColor = llColor
    End If
    lmHDLastCol = -1
    lmCellLastCol = -1
    lmCellLastRow = -1
    edcGridInfo.Visible = False
End Sub
