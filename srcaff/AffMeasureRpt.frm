VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMeasureRpt 
   Caption         =   "Affiliate Measurement Report"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "AffMeasureRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7575
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffMeasureRpt.frx":08CA
         Left            =   1080
         List            =   "AffMeasureRpt.frx":08CC
         TabIndex        =   26
         Top             =   810
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   7200
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   4920
         Picture         =   "AffMeasureRpt.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Select Stations from File.."
         Top             =   2040
         Width           =   360
      End
      Begin VB.TextBox edcWksToChart 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "52"
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin V81Affiliate.CSI_Calendar CalWeekDate 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Text            =   "12/12/2022"
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   0
      End
      Begin VB.CheckBox ckcChartIt 
         Caption         =   "Chart It"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox ckcPageSkip 
         Caption         =   "Page Skip"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox ckcDebug 
         Caption         =   "Show Internal Counts"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Frame plctotals 
         Caption         =   "Totals by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   2280
         Begin VB.OptionButton rbctotalsBy 
            Caption         =   "Detail"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton rvcTotalsBy 
            Caption         =   "Summary"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1425
         ItemData        =   "AffMeasureRpt.frx":0E38
         Left            =   3720
         List            =   "AffMeasureRpt.frx":0E3A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   3300
      End
      Begin VB.CheckBox ckcInclNetworkNC 
         Caption         =   "Include Network Non-Compliance"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ListBox lbcStations 
         Height          =   1425
         ItemData        =   "AffMeasureRpt.frx":0E3C
         Left            =   3720
         List            =   "AffMeasureRpt.frx":0E3E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   2400
         Width           =   3300
      End
      Begin VB.ListBox lbcSortChoice 
         Height          =   1425
         ItemData        =   "AffMeasureRpt.frx":0E40
         Left            =   5520
         List            =   "AffMeasureRpt.frx":0E47
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox ckcStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox ckcSort1Selection 
         Caption         =   "All"
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cbcSort1 
         Height          =   315
         Left            =   960
         TabIndex        =   15
         Top             =   1680
         Width           =   1485
      End
      Begin VB.ComboBox cbcSort2 
         Height          =   315
         ItemData        =   "AffMeasureRpt.frx":0E4E
         Left            =   960
         List            =   "AffMeasureRpt.frx":0E50
         TabIndex        =   17
         Top             =   2160
         Width           =   1485
      End
      Begin VB.CheckBox ckcSort2ZtoA 
         Caption         =   "Z to A"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox ckcInclResponse 
         Caption         =   "Include Responsiveness"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Frame frcShowBy 
         Caption         =   "Show"
         Height          =   1020
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3315
         Begin VB.OptionButton optShow 
            Caption         =   "Responsiveness"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.OptionButton optShow 
            Caption         =   "Counts by Year"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.OptionButton optShow 
            Caption         =   "Counts by Aired"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optShow 
            Caption         =   "Pct by Aired"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1365
         End
         Begin VB.OptionButton optShow 
            Caption         =   "Pct by Year"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Label lacWksToChart 
         Caption         =   "Wks"
         Enabled         =   0   'False
         Height          =   225
         Left            =   3240
         TabIndex        =   38
         Top             =   285
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   990
      End
      Begin VB.Label lacSort1 
         Caption         =   "Major Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label lacSort2 
         Caption         =   "Minor Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   945
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5865
      FormDesignWidth =   7575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   37
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   36
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4410
      TabIndex        =   35
      Top             =   225
      Width           =   2685
   End
End
Attribute VB_Name = "frmMeasureRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*  frmMeasureRpt - Reports to measure affiliate delinquency along
'                   and compliance.
'*
'*
'*  Copyright Counterpoint Software, Inc.
'

'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imChkStnListBoxIgnore As Integer
Private imChkVehListBoxignore As Integer
Private imChkSort1SelectionBoxIgnore As Integer
Private imSort1Inx As Integer
Private bmSort1ListTest As Boolean  'the 3rd list box needs to be tested
Private imSort2Inx As Integer
Private imSort2ZtoA As Integer      'true to sort desc, false to sort ascending
Private tmAmr As AMR
Private sm1stWeekDefault As String  'this does not change.  used in case input date changed, to index into the proper week
Private imWeekInx As Integer
Private imFirstWeekInx As Integer
Private imLastWeekInx As Integer

Private rst_Measure As ADODB.Recordset
Private Const SORT1_AUD = 0
Private Const SORT1_FORMAT = 1
Private Const SORT1_MKTNAME = 2
Private Const SORT1_OWNER = 3
Private Const SORT1_SALESREP = 4
Private Const SORT1_SERVICEREP = 5
Private Const SORT1_STATION = 6
Private Const SORT1_VEHICLE = 7

Private Const SORT2_NONE = 0
Private Const SORT2_AUD = 1
Private Const SORT2_MKTNAME = 2
Private Const SORT2_MKTRANK = 3
Private Const SORT2_STATION = 4
Private Const SORT2_VEHICLE = 5
Private Const SORT2_STATION_NC = 6
Private Const SORT2_NETWORK_NC = 7
Private Const SORT2_RESPONSE = 8
Private Const SORT2_WEEKS_MISSING = 9
Private Const SORT2_WEEKS_REPORTED = 10
Private Sub CalWeekDate_CalendarChanged()
    mSetCommand
End Sub

Private Sub CalWeekDate_Validate(Cancel As Boolean)
    mSetCommand
End Sub

Private Sub cbcSort1_Click()
Dim blSortSelectionOK As Boolean
Dim llLoop As Long

    bmSort1ListTest = False
    imSort1Inx = cbcSort1.ListIndex
    blSortSelectionOK = mCheckSortSelection()
    lbcSortChoice.Clear
    ckcSort1Selection.Visible = False
    lbcSortChoice.Visible = False
    lbcStations.Width = lbcVehAff.Width
    Select Case imSort1Inx
        Case SORT1_MKTNAME:      'DMA Market
            ckcSort1Selection.Caption = "All Markets"
            For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                lbcSortChoice.AddItem Trim$(tgMarketInfo(llLoop).sName)
                lbcSortChoice.ItemData(lbcSortChoice.NewIndex) = tgMarketInfo(llLoop).lCode
            Next llLoop
            ckcSort1Selection.Visible = True
            lbcSortChoice.Visible = True
            lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the market list
            bmSort1ListTest = True                              'choice selected whereby 3rd list box needs to be displayed
        Case SORT1_FORMAT:         'Format
            ckcSort1Selection.Caption = "All Formats"
           For llLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
                lbcSortChoice.AddItem Trim$(tgFormatInfo(llLoop).sName)
                lbcSortChoice.ItemData(lbcSortChoice.NewIndex) = tgFormatInfo(llLoop).lCode
            Next llLoop
            ckcSort1Selection.Visible = True
            lbcSortChoice.Visible = True
            lbcStations.Width = (lbcVehAff.Width / 2) - 360 'make station list half width in order to show the format list
            bmSort1ListTest = True                          'choice selected whereby 3rd list box needs to be displayed
        Case SORT1_OWNER:         'Owner
            ckcSort1Selection.Caption = "All Owners"
            For llLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
                lbcSortChoice.AddItem Trim$(tgOwnerInfo(llLoop).sName)
                lbcSortChoice.ItemData(lbcSortChoice.NewIndex) = tgOwnerInfo(llLoop).lCode
            Next llLoop
            ckcSort1Selection.Visible = True
            lbcSortChoice.Visible = True
            lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the owner list
            bmSort1ListTest = True                              'choice selected whereby 3rd list box needs to be displayed
        Case SORT1_SALESREP:         '
            ckcSort1Selection.Caption = "All Sales Reps"
            For llLoop = 0 To UBound(tgMarketRepInfo) - 1 Step 1
                lbcSortChoice.AddItem Trim$(tgMarketRepInfo(llLoop).sName)
                lbcSortChoice.ItemData(lbcSortChoice.NewIndex) = tgMarketRepInfo(llLoop).iUstCode
            Next llLoop
            ckcSort1Selection.Visible = True
            lbcSortChoice.Visible = True
            lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the Sales rep list
            bmSort1ListTest = True                              'choice selected whereby 3rd list box needs to be displayed
        Case SORT1_SERVICEREP:         '
            ckcSort1Selection.Caption = "All Service Reps"
            For llLoop = 0 To UBound(tgServiceRepInfo) - 1 Step 1
                lbcSortChoice.AddItem Trim$(tgServiceRepInfo(llLoop).sName)
                lbcSortChoice.ItemData(lbcSortChoice.NewIndex) = tgServiceRepInfo(llLoop).iUstCode
            Next llLoop
            ckcSort1Selection.Visible = True
            lbcSortChoice.Visible = True
            lbcStations.Width = (lbcVehAff.Width / 2) - 360     'make station list half width in order to show the service rep list
            bmSort1ListTest = True                              'choice selected whereby 3rd list box needs to be displayed
    End Select
    mSetCommand
End Sub

Private Sub cbcSort2_Click()
Dim blSortSelectionOK As Boolean
Dim llLoop As Long

    imSort2Inx = cbcSort2.ListIndex
    If imSort2Inx >= 6 Then
        ckcSort2ZtoA.Enabled = True
    Else
        ckcSort2ZtoA.Enabled = False
        ckcSort2ZtoA.Value = vbUnchecked
    End If
    
    mSetCommand
End Sub

Private Sub chkAllVehicles_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkVehListBoxignore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkVehListBoxignore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    mSetCommand
End Sub

Private Sub ckcChartIt_Click()
    If ckcChartIt.Value = vbChecked Then
        optShow(0).Caption = "Wks Missing"
        optShow(0).Value = True
        optShow(1).Caption = "Station N/C"
        optShow(2).Caption = "Network N/C"
        optShow(3).Caption = "Wks Reported"
        optShow(4).Caption = "Responsiveness"
        optShow(0).Visible = True
        optShow(1).Visible = True
        optShow(2).Visible = True
        optShow(3).Visible = True
        optShow(4).Visible = True
        cbcSort2.ListIndex = 0
        lacSort2.Enabled = False
        cbcSort2.Enabled = False
        ckcSort2ZtoA.Enabled = False
        edcWksToChart.Enabled = True
        edcWksToChart.Text = 52
        lacWksToChart.Enabled = True
    Else
        optShow(0).Caption = "Counts by Aired"
        optShow(0).Value = True
        optShow(2).Caption = "Pct by Aired"
        optShow(3).Caption = "Pct by Year"
        optShow(0).Visible = True
        optShow(0).Value = True
        optShow(1).Visible = False
        optShow(2).Visible = True
        optShow(3).Visible = True
        optShow(4).Visible = False
        lacSort2.Enabled = True
        cbcSort2.Enabled = True
        ckcSort2ZtoA.Enabled = True
        edcWksToChart.Enabled = False
        lacWksToChart.Enabled = False
    End If
End Sub

Private Sub ckcSort1Selection_Click()

    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkSort1SelectionBoxIgnore Then
        Exit Sub
    End If
    If ckcSort1Selection.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSortChoice.ListCount > 0 Then
        imChkSort1SelectionBoxIgnore = True
        lRg = CLng(lbcSortChoice.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSortChoice.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkSort1SelectionBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    mSetCommand
End Sub

Private Sub ckcStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStnListBoxIgnore Then
        Exit Sub
    End If
    If ckcStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStnListBoxIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStnListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    mSetCommand
End Sub

Private Sub cmdDone_Click()
    Unload frmMeasureRpt
End Sub
'
'       Affiliate Measurement - Measure affiliate delinquency and responsiveness
'       records are built for each week in smt from a background program, run as often
'       as on demand by client.  Data is updated for vehicle/station for # weeks aired, # weeks missing
'       # spots posted, # spots posted failing station compliance, # spots posteed failing network compliance,
'       # unique dates that station submitted posting in last 52 weeks
'       All this data will be shown in a variety of different reports, showing either counts or pct of
'       data
Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim sName, sVehicles, sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim slStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slNow As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilInclVehicleCodes As Integer
    Dim ilUseVehicleCodes() As Integer
    Dim ilInclStationCodes As Integer
    Dim ilUseStationCodes() As Integer
    Dim ilInclChoiceCodes As Integer
    Dim llUseChoiceCodes() As Long   'if selecting format, owner, market, sales or service rep, the list of incl/excl is in array
    Dim slWhichChoice As String
    Dim ilLoop As Integer
    Dim blValidMkt As Boolean
    Dim blValidFmt As Boolean
    Dim blValidOwner As Boolean
    Dim blValidSalesRep As Boolean
    Dim blValidServRep As Boolean
    Dim slRptName As String

    Dim llVefInx As Long
    Dim ilShttInx As Integer
    Dim ilMktInx As Integer
    Dim ilFmtInx As Integer
    Dim llOwnerInx As Long
    Dim ilMktRepInx As Integer
    Dim ilServRepInx As Integer
    Dim llTempLong As Long
    Dim llTempLong2 As Long
    Dim ilWeeksAiredInx As Integer
    Dim ilWeeksMissingInx As Integer
    Dim ilSpotPostedInx As Integer
    Dim ilSpotPostedSNCInx As Integer
    Dim ilSpotPostedNNCInx As Integer
    Dim ilDaysSubmittedInx As Integer
    Dim slRunDate As String
    Dim llDate As Long
    Dim slDateOfWeekInfo As String
    Dim ilCount As Integer
    
    On Error GoTo ErrHand
    
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmMeasureRpt, "Aired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    sStartDate = Trim$(CalWeekDate.Text)                        'user input/may or may not have been changed.  if changed, need to index into the proper week based on the 1st weeks date in the record
    If gIsDate(sStartDate) = False Or (Len(Trim$(sStartDate)) = 0) Or Trim$(sm1stWeekDefault) = "" Then
        Beep
        If Trim$(sm1stWeekDefault) = "" Then
             gMsgBox "No affiliate measurement data has been generated", vbCritical
       Else
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        End If
        CalWeekDate.SetFocus
        Exit Sub
    End If
    
    sEndDate = DateAdd("d", 6, sm1stWeekDefault)            'calc end date based on orig current first week
    llStartDate = gDateValue(sm1stWeekDefault)
    llEndDate = gDateValue(sEndDate)
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    
    Do While Weekday(sStartDate, vbSunday) <> vbMonday
        sStartDate = DateAdd("d", -1, sStartDate)
    Loop
    'if start date is greater than the default date, invalid date
    If gDateValue(sStartDate) + 6 > llEndDate Or sm1stWeekDefault = "" Then
        Beep
        gMsgBox "Date entered has not been generated yet", vbCritical
        CalWeekDate.SetFocus
        Exit Sub
    End If
    
    '3-31-17 date entered cannot be more than 52 weeks prior to the default date
    If ((gDateValue(sm1stWeekDefault) - gDateValue(sStartDate)) / 7) + 1 > 52 Then
        Beep
        llTempLong = (gDateValue(sm1stWeekDefault) - (51 * 7)) - 1      '4-7-17 can only go 52 weeks back
        slStr = Trim$(Format(llTempLong, "ddddd"))
        gMsgBox "Date must be later than " + slStr, vbCritical
        CalWeekDate.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'calculate the index into the proper week
    llTempLong = DateValue(sm1stWeekDefault)
    llTempLong2 = DateValue(sStartDate)
    
    'imWeekInx = ((llTempLong - llTempLong2) / 7)
    imFirstWeekInx = ((llTempLong - llTempLong2) / 7) + 1           '3-31-17 adjustment for # weeks to process, always at least 1
    imLastWeekInx = imFirstWeekInx                                   'default to process one week only, or if charting, could be up to 52 weeks
    slRptName = "AfMeasure"
    If ckcChartIt.Value = vbChecked Then
        imLastWeekInx = 51 - imFirstWeekInx           'index 51 (relative to 0) is max to process
        'imLastWeekInx = 51                           '51 weeks (number is relative to 0)
        If (Val(edcWksToChart.Text) + imFirstWeekInx) - 1 < imLastWeekInx Then
            imLastWeekInx = Val(edcWksToChart.Text) + imFirstWeekInx - 1
        End If
        slRptName = "AfMeasureChartIt"
    End If
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)
    gPackDate sgGenDate, tmAmr.iGenDate(0), tmAmr.iGenDate(1)
    tmAmr.lGenTime = gTimeToLong(sgGenTime, False)
    
    dFWeek = CDate(sStartDate)
    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula9 = "'" & sm1stWeekDefault & "'"
 
    sgCrystlFormula2 = Str$(cbcSort1.ListIndex)          'sort fields, major to minor
    sgCrystlFormula3 = Str$(cbcSort2.ListIndex)
    
    sgCrystlFormula4 = "A"                              'ascending or descending
    If ckcSort2ZtoA.Value = vbChecked Then
        sgCrystlFormula4 = "Z"
    End If
    If rbctotalsBy(0).Value = True Then
        sgCrystlFormula5 = "D"                          'show detail
    Else
        sgCrystlFormula5 = "S"                          'summary
    End If
    If ckcInclNetworkNC.Value = vbChecked Then
        sgCrystlFormula6 = "I"
    Else
        sgCrystlFormula6 = "E"
    End If
    If ckcInclResponse.Value = vbChecked Then
        sgCrystlFormula7 = "I"
    Else
        sgCrystlFormula7 = "E"
    End If
    
    If ckcDebug.Value = vbChecked Then           'show internal codes
        sgCrystlFormula8 = "Y"
    Else
        sgCrystlFormula8 = "N"
    End If
    
    If ckcChartIt.Value = vbChecked Then        'for charting, assume counts
        sgCrystlFormula10 = "C"
    Else
        If optShow(0).Value = True Or optShow(1).Value = True Then     'by counts by aired (0) or year which has been removed for now (1)
            sgCrystlFormula10 = "C"
        Else
            sgCrystlFormula10 = "P"
        End If
    End If
    
    If ckcChartIt.Value = vbChecked Then        'for charting, assume counts by aired
        sgCrystlFormula11 = "A"
    Else
        If optShow(0).Value = True Or optShow(2).Value = True Then     'counts by air (0) or pct by air (2)
            sgCrystlFormula11 = "A"
        Else
            sgCrystlFormula11 = "Y"
        End If
    End If
    
    sgCrystlFormula12 = sgClientName
    If ckcPageSkip.Value = vbChecked Then           'skip to new page each major change
        sgCrystlFormula13 = "Y"
    Else
        sgCrystlFormula13 = "N"
    End If
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    
    sVehicles = ""
    sStations = ""
    slWhichChoice = ""          'one of the choices from 3rd list boxes (format, owner, market, sales rep, service rep)
    
    'ReDim ilUseVehicleCodes(1 To 1) As Integer
    ReDim ilUseVehicleCodes(0 To 0) As Integer
    'ReDim ilUseStationCodes(1 To 1) As Integer
    ReDim ilUseStationCodes(0 To 0) As Integer
    'ReDim llUseChoiceCodes(1 To 1) As Long
    ReDim llUseChoiceCodes(0 To 0) As Long
    gObtainCodes lbcVehAff, ilInclVehicleCodes, ilUseVehicleCodes()        'build array of which codes to incl/excl
    For ilLoop = LBound(ilUseVehicleCodes) To UBound(ilUseVehicleCodes) - 1
        If Trim$(sVehicles) = "" Then
            If ilInclVehicleCodes = True Then                          'include the list
                sVehicles = " IN (" & Str(ilUseVehicleCodes(ilLoop))
            Else                                                        'exclude the list
                sVehicles = " Not IN (" & Str(ilUseVehicleCodes(ilLoop))
            End If
        Else
            sVehicles = sVehicles & "," & Str(ilUseVehicleCodes(ilLoop))
        End If
    Next ilLoop
    If sVehicles <> "" Then
        sVehicles = sVehicles & ")"
    End If
    gObtainCodes lbcStations, ilInclStationCodes, ilUseStationCodes()        'build array of which advt codes to incl/excl
    For ilLoop = LBound(ilUseStationCodes) To UBound(ilUseStationCodes) - 1
        If Trim$(sStations) = "" Then
            If ilInclStationCodes = True Then                          'include the list
                sStations = " IN (" & Str(ilUseStationCodes(ilLoop))
            Else                                                        'exclude the list
                sStations = " Not IN (" & Str(ilUseStationCodes(ilLoop))
            End If
        Else
            sStations = sStations & "," & Str(ilUseStationCodes(ilLoop))
        End If
    Next ilLoop
    If sStations <> "" Then
        sStations = sStations & ")"
    End If
    
    If bmSort1ListTest Then                 '3rd list box to test
        gObtainCodesLong lbcSortChoice, ilInclChoiceCodes, llUseChoiceCodes()        'build array of which advt codes to incl/excl
    End If
    
    
'    SQLQuery = "Select count(smtcode) FROM smt  "
'    SQLQuery = SQLQuery & "inner join Vef_vehicles on smtvefcode = vefcode "
'    SQLQuery = SQLQuery & "inner join shtt on smtshttcode = shttcode "
'    SQLQuery = SQLQuery & "WHERE (smtWk1StartDate >= '" & Format$(sm1stWeekDefault, sgSQLDateForm) & "' AND smtWk1StartDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
'    If Trim$(sVehicles) <> "" Then
'        SQLQuery = SQLQuery & " and (vefcode " & sVehicles & ")"
'    End If
'    If Trim$(sStations) <> "" Then
'        SQLQuery = SQLQuery & " and (shttCode " & sStations & ")"
'    End If
'    Set rst_Measure = gSQLSelectCall(SQLQuery)
    
    SQLQuery = "Select * FROM smt  "
    SQLQuery = SQLQuery & "inner join Vef_vehicles on smtvefcode = vefcode "
    SQLQuery = SQLQuery & "inner join shtt on smtshttcode = shttcode "
    SQLQuery = SQLQuery & "WHERE (smtWk1StartDate >= '" & Format$(sm1stWeekDefault, sgSQLDateForm) & "' AND smtWk1StartDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    If Trim$(sVehicles) <> "" Then
        SQLQuery = SQLQuery & " and (vefcode " & sVehicles & ")"
    End If
    If Trim$(sStations) <> "" Then
        SQLQuery = SQLQuery & " and (shttCode " & sStations & ")"
    End If
    
    ilCount = 0
    
    Set rst_Measure = gSQLSelectCall(SQLQuery)
    While Not rst_Measure.EOF
        ilCount = ilCount + 1
        tmAmr.sCallLetters = ""
        tmAmr.sVehicleName = ""
        tmAmr.lAudience = 0
        tmAmr.sFormat = ""
        tmAmr.sMarket = ""
        tmAmr.iRank = 0
        tmAmr.sOwner = ""
        tmAmr.sSalesRep = ""
        tmAmr.sServRep = ""
        
        blValidMkt = True
        blValidFmt = True
        blValidOwner = True
        blValidSalesRep = True
        blValidServRep = True
               
        llVefInx = gBinarySearchVef(CLng(rst_Measure!smtvefcode))
        If llVefInx <> -1 Then
            tmAmr.sVehicleName = Trim$(tgVehicleInfo(llVefInx).sVehicleName)
        
            ilShttInx = gBinarySearchStationInfoByCode(rst_Measure!smtshttcode)
            If ilShttInx <> -1 Then
                tmAmr.sCallLetters = Trim$(tgStationInfoByCode(ilShttInx).sCallLetters)
                tmAmr.lAudience = tgStationInfoByCode(ilShttInx).lAudP12Plus
                ilMktInx = gBinarySearchMkt(CLng(tgStationInfoByCode(ilShttInx).iMktCode))
                If ilMktInx <> -1 Then
                    tmAmr.sMarket = Trim$(tgMarketInfo(ilMktInx).sName)
                    tmAmr.iRank = tgMarketInfo(ilMktInx).iRank
                    blValidMkt = mTestSort1ListTest(SORT1_MKTNAME, (CLng(tgMarketInfo(ilMktInx).lCode)), ilInclChoiceCodes, llUseChoiceCodes())

                End If
                ilFmtInx = gBinarySearchFmt(CLng(tgStationInfoByCode(ilShttInx).iFormatCode))
                If ilFmtInx <> -1 Then
                    tmAmr.sFormat = Trim$(tgFormatInfo(ilFmtInx).sName)
                    blValidFmt = mTestSort1ListTest(SORT1_FORMAT, (CLng(tgStationInfoByCode(ilShttInx).iFormatCode)), ilInclChoiceCodes, llUseChoiceCodes())
                Else
                    If (imSort1Inx = SORT1_FORMAT) And (ckcSort1Selection.Value = vbUnchecked) Then        'if sorting by format and selective formats, do not include the stats without format names
                        blValidFmt = False
                    End If
                End If
       
                llOwnerInx = gBinarySearchOwner(CLng(tgStationInfoByCode(ilShttInx).lOwnerCode))
                If llOwnerInx <> -1 Then
                    tmAmr.sOwner = Trim$(tgOwnerInfo(llOwnerInx).sName)
                    blValidOwner = mTestSort1ListTest(SORT1_OWNER, (tgStationInfoByCode(ilShttInx).lOwnerCode), ilInclChoiceCodes, llUseChoiceCodes())
                Else
                    If (imSort1Inx = SORT1_OWNER) And (ckcSort1Selection.Value = vbUnchecked) Then        'if sorting by owner and selective owners, do not include the stats without owner names
                        blValidOwner = False
                    End If
                End If
                
                ilMktRepInx = gBinarySearchRepInfo(CLng(tgStationInfoByCode(ilShttInx).iMktRepUstCode), tgMarketRepInfo())
                If ilMktRepInx <> -1 Then
                    tmAmr.sSalesRep = Trim$(tgMarketRepInfo(ilMktRepInx).sName)
                    blValidSalesRep = mTestSort1ListTest(SORT1_SALESREP, (CLng(tgStationInfoByCode(ilShttInx).iMktRepUstCode)), ilInclChoiceCodes, llUseChoiceCodes())
                Else
                    If (imSort1Inx = SORT1_SALESREP) And (ckcSort1Selection.Value = vbUnchecked) Then        'if sorting by SalesRep and selective Reps, do not include the stats without rep names
                        blValidSalesRep = False
                    End If
                End If
             
                ilServRepInx = gBinarySearchRepInfo(CLng(tgStationInfoByCode(ilShttInx).iServRepUstCode), tgServiceRepInfo())
                If ilServRepInx <> -1 Then
                    tmAmr.sServRep = Trim$(tgServiceRepInfo(ilServRepInx).sName)
                    blValidServRep = mTestSort1ListTest(SORT1_SERVICEREP, (CLng(tgStationInfoByCode(ilShttInx).iServRepUstCode)), ilInclChoiceCodes, llUseChoiceCodes())
                 Else
                    If (imSort1Inx = SORT1_SERVICEREP) And (ckcSort1Selection.Value = vbUnchecked) Then        'if sorting by service rep and selective reps, do not include the stats without service rep names
                        blValidServRep = False
                    End If
                End If
             
            Else
                blValidMkt = False   'flag at least one filter false, no station found
            End If
        Else
            blValidMkt = False          'flag at least one filter false, no vehicle found
        End If
        If (blValidMkt) And (blValidFmt) And (blValidOwner) And (blValidSalesRep) And (blValidServRep) Then
            For imWeekInx = imFirstWeekInx To imLastWeekInx
                slStr = Format$(rst_Measure!smtGenDate, sgShowDateForm)         'date the data was updated with new info
                gPackDate slStr, tmAmr.iRunDate(0), tmAmr.iRunDate(1)
                'slStr = Format$(rst_Measure!smtWk1StartDate, sgShowDateForm)     'date of 1st week of 52 weeks
                slDateOfWeekInfo = sStartDate                                                'date user requested, could be the 1st date of the record of indexed into a different week of the 52 weeks
                llDate = gDateValue(slDateOfWeekInfo) + (7 * imWeekInx)
                slDateOfWeekInfo = Format$(llDate, "m/d/yy")
                gPackDate slDateOfWeekInfo, tmAmr.iWeekInfoDate(0), tmAmr.iWeekInfoDate(1)
                tmAmr.lCode = 0
                tmAmr.lSmtCode = rst_Measure!smtcode
                'the index into the fields must be hard-coded since theres no way to index into an array using the result sets.
                'the indices are field # offsets, NOT byte offsets (the field offsets should be relative to 0)
                '3-31-17
                tmAmr.iDaysSubmitted = rst_Measure(268 + imWeekInx - 1).Value  'smtDaysSubmitted
                tmAmr.iWeeksAired = rst_Measure(8 + imWeekInx - 1).Value       'smtWeeksAired1
                tmAmr.iWeeksMissing = rst_Measure(60 + imWeekInx - 1).Value     'smtWeeksMissing1
                tmAmr.lSpotsPosted = rst_Measure(112 + imWeekInx - 1).Value     'smtSpotPosted
                tmAmr.lSpotsPostedSNC = rst_Measure(164 + imWeekInx - 1).Value      'smtSpotPostsNC
                tmAmr.lSpotsPostedNNC = rst_Measure(216 + imWeekInx - 1).Value      'smtSpotPostedNNC
                
                tmAmr.iVefCode = rst_Measure!smtvefcode
                tmAmr.iShttCode = rst_Measure!smtshttcode
                gUnpackDate tmAmr.iRunDate(0), tmAmr.iRunDate(1), slRunDate
                slRunDate = Format$(slRunDate, sgShowDateForm)
                tmAmr.sUnused = ""
                tmAmr.sString1 = ""
            
                SQLQuery = "Insert Into amr ( "
                SQLQuery = SQLQuery & "amrCode, "
                SQLQuery = SQLQuery & "amrGenDate, "
                SQLQuery = SQLQuery & "amrGenTime, "
                SQLQuery = SQLQuery & "amrSmtCode, "
                SQLQuery = SQLQuery & "amrAudience, "
                SQLQuery = SQLQuery & "amrWeekInfoDate, "
                SQLQuery = SQLQuery & "amrRunDate, "
                SQLQuery = SQLQuery & "amrVehicleName, "
                SQLQuery = SQLQuery & "amrCallLetters, "
                SQLQuery = SQLQuery & "amrMarket, "
                SQLQuery = SQLQuery & "amrRank, "
                SQLQuery = SQLQuery & "amrFormat, "
                SQLQuery = SQLQuery & "amrOwner, "
                SQLQuery = SQLQuery & "amrSalesRep, "
                SQLQuery = SQLQuery & "amrServRep, "
                SQLQuery = SQLQuery & "amrWeeksAired, "
                SQLQuery = SQLQuery & "amrWeeksMissing, "
                SQLQuery = SQLQuery & "amrSpotsPosted, "
                SQLQuery = SQLQuery & "amrSpotsPostedSNC, "
                SQLQuery = SQLQuery & "amrSpotsPostedNNC, "
                SQLQuery = SQLQuery & "amrDaysSubmitted, "
                SQLQuery = SQLQuery & "amrShttCode, "
                SQLQuery = SQLQuery & "amrVefCode, "
                SQLQuery = SQLQuery & "amrString1, "
                SQLQuery = SQLQuery & "amrUnused "
                SQLQuery = SQLQuery & ") "
                
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & tmAmr.lCode & ", "
                SQLQuery = SQLQuery & "'" & Format$(sgGenDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & ", "
                SQLQuery = SQLQuery & tmAmr.lSmtCode & ", "
                SQLQuery = SQLQuery & tmAmr.lAudience & ", "
                SQLQuery = SQLQuery & "'" & Format$(slDateOfWeekInfo, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(slRunDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sVehicleName)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sCallLetters)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sMarket)) & "', "
                SQLQuery = SQLQuery & tmAmr.iRank & ", "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sFormat)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sOwner)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sSalesRep)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sServRep)) & "', "
                SQLQuery = SQLQuery & tmAmr.iWeeksAired & ", "
                SQLQuery = SQLQuery & tmAmr.iWeeksMissing & ", "
                SQLQuery = SQLQuery & tmAmr.lSpotsPosted & ", "
                SQLQuery = SQLQuery & tmAmr.lSpotsPostedSNC & ", "
                SQLQuery = SQLQuery & tmAmr.lSpotsPostedNNC & ", "
                SQLQuery = SQLQuery & tmAmr.iDaysSubmitted & ", "
                SQLQuery = SQLQuery & tmAmr.iShttCode & ", "
                SQLQuery = SQLQuery & tmAmr.iVefCode & ", "
                SQLQuery = SQLQuery & "''" & ", "
                SQLQuery = SQLQuery & "''" & " "
                SQLQuery = SQLQuery & ") "
    
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "MeasureRpt-cmdReport_Click"
                    Exit Sub
                End If
            Next imWeekInx                          'for imWeekinx
            
        End If
        rst_Measure.MoveNext
        If ilCount Mod 100 = 0 Then
            ilCount = ilCount
        End If
    Wend

    gUserActivityLog "E", sgReportListName & ": Prepass"
    
    SQLQuery = "Select * from amr where (amrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"

    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, Trim$(slRptName) & ".rpt", Trim$(slRptName)
     
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime

    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    'remove all the records just printed
    SQLQuery = "DELETE FROM amr "
    SQLQuery = SQLQuery & " WHERE (amrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "MeasureRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    rst_Measure.Close
    
    cmdReport.Enabled = True               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmPgmClrRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmMeasureRpt
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
Private Sub cmdStationListFile_Click()
    Dim slCurDir As String
    slCurDir = CurDir
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    ' Import from the Selected File
    gSelectiveStationsFromImport lbcStations, ckcStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub edcWksToChart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
        cbcSort1.AddItem "Audience"
        cbcSort1.AddItem "Format"
        cbcSort1.AddItem "Market Name"
        cbcSort1.AddItem "Owner"
        cbcSort1.AddItem "Sales Rep"
        cbcSort1.AddItem "Service Rep"
        cbcSort1.AddItem "Station"
        cbcSort1.AddItem "Vehicle"
        cbcSort1.ListIndex = 7
        
        cbcSort2.AddItem "None"
        cbcSort2.AddItem "Audience"
        cbcSort2.AddItem "Market Name"
        cbcSort2.AddItem "Market Rank"
        cbcSort2.AddItem "Station"
        cbcSort2.AddItem "Vehicle"
        cbcSort2.AddItem "Station Non-Compliant"
        cbcSort2.AddItem "Network Non-Compliant"
        cbcSort2.AddItem "Responsiveness"
        cbcSort2.AddItem "Weeks Missing"
        cbcSort2.AddItem "Weeks Reported"
        cbcSort2.ListIndex = 0
End Sub

Private Sub Form_Initialize()
Dim ilHalf As Integer
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
 '   ckcAdvt.Top = 240
    
'    ilHalf = (Frame2.Height - ckcAdvt.Height - chkAllVehicles.Height - 120) / 2
'    lbcAdvt.Move ckcAdvt.Left, ckcAdvt.Top + ckcAdvt.Height
'    lbcAdvt.Height = ilHalf
'    lbcVehAff.Height = ilHalf
'    lbcStations.Height = ilHalf
'    chkAllVehicles.Top = lbcAdvt.Top + lbcAdvt.Height
'    ckcStations.Top = chkAllVehicles.Top
'    lbcVehAff.Top = chkAllVehicles.Top + chkAllVehicles.Height
'    lbcStations.Top = lbcVehAff.Top
'    lbcVehAff.Height = ilHalf
'    lbcStations.Height = ilHalf

    'lbcStations.Width = 3300

    gSetFonts frmMeasureRpt
    gCenterForm frmMeasureRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim ilRet As Integer
    Dim lRg As Long
    Dim lRet As Long
    Dim slDateDefault As String
    Dim slName As String

    imChkListBoxIgnore = False
    imChkVehListBoxignore = False
    imChkStnListBoxIgnore = False
    frmMeasureRpt.Caption = "Affiliate Measurement Report - " & sgClientName
    
   
    'populate the Stations, Vehicles & Advertisers (currently only advertisers are selectable)
    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
    
    lbcVehAff.Clear

    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    gPopRepInfo "M", tgMarketRepInfo()
    gPopRepInfo "S", tgServiceRepInfo()
        
    sgGenDate = Format$(gNow(), "m/d/yyyy")
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    'backup to monday, then need to default to previous week
    slDateDefault = ""
    sm1stWeekDefault = ""
    SQLQuery = "Select max(smtWk1StartDate) from smt"
    Set rst_Measure = gSQLSelectCall(SQLQuery)
    If Not rst_Measure.EOF Then
       slDateDefault = Format$(rst_Measure(0), sgShowDateForm)
        CalWeekDate.Text = slDateDefault
        sm1stWeekDefault = slDateDefault            'this does not change.  used in case input date changed, to index into the proper week
    End If
        
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

    If igUstCode = 1 Then                       'system csi; for testing
        ckcChartIt.Visible = True
        edcWksToChart.Visible = True
        lacWksToChart.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmMeasureRpt = Nothing
End Sub


Private Sub lbcSortChoice_Click()
    If imChkSort1SelectionBoxIgnore Then
        Exit Sub
    End If
    If ckcSort1Selection.Value = vbChecked Then
        imChkSort1SelectionBoxIgnore = True
        ckcSort1Selection.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkSort1SelectionBoxIgnore = False
    End If
    mSetCommand
End Sub

Private Sub lbcStations_Click()
    If imChkStnListBoxIgnore Then
        Exit Sub
    End If
    If ckcStations.Value = vbChecked Then
        imChkStnListBoxIgnore = True
        ckcStations.Value = 0    'chged from false to 0 10-22-99
        imChkStnListBoxIgnore = False
    End If
    mSetCommand
End Sub
Private Sub lbcVehAff_Click()
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = vbChecked Then
        imChkVehListBoxignore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = 0    'chged from false to 0 10-22-99
        imChkVehListBoxignore = False
    End If
    mSetCommand
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub


'       mTestSortSelection - Test sort selections on Major and Minor sort fields;
'       both selections cannot be the same
Public Function mCheckSortSelection() As Boolean

Dim blSortSelectionOK As Boolean

    blSortSelectionOK = True
    If imSort1Inx = SORT1_AUD And imSort2Inx = SORT2_AUD Then
        blSortSelectionOK = False
    ElseIf imSort1Inx = SORT1_MKTNAME And imSort2Inx = SORT2_MKTNAME Then
        blSortSelectionOK = False
    ElseIf imSort1Inx = SORT1_STATION And imSort2Inx = SORT2_STATION Then
        blSortSelectionOK = False
    ElseIf imSort1Inx = SORT1_VEHICLE And imSort2Inx = SORT2_VEHICLE Then
        blSortSelectionOK = False
    End If
    If Not blSortSelectionOK Then
        MsgBox "Major and minor sorts cannot be the same"
    End If
    
   mCheckSortSelection = blSortSelectionOK

End Function
Private Sub mSetCommand()
Dim blSortSelectionOK As Boolean
Dim blSetGenOK As Boolean

    blSortSelectionOK = mCheckSortSelection()
    blSetGenOK = True
    If lbcVehAff.SelCount = 0 Or lbcStations.SelCount = 0 Then
        blSetGenOK = False
    End If
    If CalWeekDate.Text = "" Then
        blSetGenOK = False
    End If
    
    If (blSetGenOK) And (blSortSelectionOK) Then
        If imSort1Inx = SORT1_MKTNAME Or imSort1Inx = SORT1_FORMAT Or imSort1Inx = SORT1_OWNER Or imSort1Inx = SORT1_SALESREP Or imSort1Inx = SORT1_SERVICEREP Then
            If lbcSortChoice.SelCount = 0 Then
                cmdReport.Enabled = False
            Else
                cmdReport.Enabled = True
            End If
        Else
            cmdReport.Enabled = True
        End If
    Else
        cmdReport.Enabled = False
    End If
End Sub
'
'               mTestSort1ListTest - Test one of the categories (other than vehicle and station) that
'               user has selected for possible filters
'               Format, Market, Owner, Sales Rep or Service Rep
Private Function mTestSort1ListTest(ilWhichList As Integer, llValue As Long, ilIncludeCodes As Integer, llUseCodes() As Long) As Boolean
Dim blFound As Boolean

        blFound = True              'default if no testing
        If ilWhichList = SORT1_MKTNAME And imSort1Inx = SORT1_MKTNAME Then
            blFound = gTestIncludeExcludeLong(llValue, ilIncludeCodes, llUseCodes())
        ElseIf ilWhichList = SORT1_FORMAT And imSort1Inx = SORT1_FORMAT Then
            blFound = gTestIncludeExcludeLong(llValue, ilIncludeCodes, llUseCodes())
        ElseIf ilWhichList = SORT1_OWNER And imSort1Inx = SORT1_OWNER Then
            blFound = gTestIncludeExcludeLong(llValue, ilIncludeCodes, llUseCodes())
        ElseIf ilWhichList = SORT1_SALESREP And imSort1Inx = SORT1_SALESREP Then
            blFound = gTestIncludeExcludeLong(llValue, ilIncludeCodes, llUseCodes())
        ElseIf ilWhichList = SORT1_SERVICEREP And imSort1Inx = SORT1_SERVICEREP Then
            blFound = gTestIncludeExcludeLong(llValue, ilIncludeCodes, llUseCodes())
        End If
        
    mTestSort1ListTest = blFound
End Function

