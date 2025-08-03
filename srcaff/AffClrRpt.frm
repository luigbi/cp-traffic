VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmClearRpt 
   Caption         =   "Advertiser Clearance Report"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "AffClrRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9165
   Begin VB.Timer tmcQueue 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8115
      Top             =   1200
   End
   Begin VB.CommandButton cmdSendToQueue 
      Caption         =   "Send To Queue"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   225
      Visible         =   0   'False
      Width           =   1740
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
      FormDesignHeight=   5775
      FormDesignWidth =   9165
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
      Height          =   3780
      Left            =   195
      TabIndex        =   6
      Top             =   1830
      Width           =   8760
      Begin V81Affiliate.CSI_Calendar calEffWeek 
         Height          =   270
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         Text            =   "10/1/2007"
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
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Not Carried"
         Height          =   255
         Index           =   1
         Left            =   1770
         TabIndex        =   16
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox ckcAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   195
         Left            =   3750
         TabIndex        =   30
         Top             =   1800
         Width           =   1380
      End
      Begin VB.ListBox lbcVehicles 
         Height          =   1425
         ItemData        =   "AffClrRpt.frx":08CA
         Left            =   3720
         List            =   "AffClrRpt.frx":08D1
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   2160
         Width           =   4905
      End
      Begin VB.ListBox lbcStations 
         Height          =   1425
         ItemData        =   "AffClrRpt.frx":08D9
         Left            =   6240
         List            =   "AffClrRpt.frx":08E0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Frame Frame4 
         Caption         =   "Spots pledged to air in daypart"
         Height          =   780
         Left            =   120
         TabIndex        =   24
         Top             =   2670
         Width           =   3075
         Begin VB.OptionButton optSpotAired 
            Caption         =   "Show Exact Times"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   26
            Top             =   480
            Width           =   2340
         End
         Begin VB.OptionButton optSpotAired 
            Caption         =   "Show Daypart"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   2070
         End
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1275
         Width           =   3240
      End
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Include Not Aired"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   990
         Width           =   1935
      End
      Begin VB.TextBox txtNoWeeks 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   615
         Width           =   525
      End
      Begin VB.Frame ShowSpotAired 
         Caption         =   "Sort contracts by"
         Height          =   1020
         Left            =   120
         TabIndex        =   18
         Top             =   1590
         Width           =   3540
         Begin VB.OptionButton optSortby 
            Caption         =   "MSA Market Rank"
            Height          =   255
            Index           =   4
            Left            =   1785
            TabIndex        =   23
            Top             =   480
            Width           =   1650
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "MSA Market Name"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1650
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Call Letters"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   21
            Top             =   720
            Width           =   1440
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "DMA Market Name"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1710
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "DMA Market Rank"
            Height          =   255
            Index           =   1
            Left            =   1785
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All Contracts"
         Height          =   195
         Left            =   6180
         TabIndex        =   28
         Top             =   180
         Width           =   1860
      End
      Begin VB.ListBox lbcContract 
         Height          =   1230
         ItemData        =   "AffClrRpt.frx":08E7
         Left            =   6180
         List            =   "AffClrRpt.frx":08E9
         MultiSelect     =   2  'Extended
         TabIndex        =   29
         Top             =   480
         Width           =   2340
      End
      Begin VB.ListBox lbcAdvertiser 
         Height          =   1230
         ItemData        =   "AffClrRpt.frx":08EB
         Left            =   3750
         List            =   "AffClrRpt.frx":08ED
         TabIndex        =   27
         Top             =   480
         Width           =   2340
      End
      Begin VB.Label Label2 
         Caption         =   "Aired Start Week"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "# Weeks"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   1020
      End
      Begin VB.Label lacTitle1 
         Caption         =   "Advertisers"
         Height          =   255
         Left            =   3750
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4620
      TabIndex        =   8
      Top             =   705
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   7
      Top             =   225
      Width           =   2685
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
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffClrRpt.frx":08EF
         Left            =   840
         List            =   "AffClrRpt.frx":08F1
         TabIndex        =   4
         Top             =   795
         Width           =   1935
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1185
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   840
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmClearRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'* Advertiser Posting (Clearance) Detail
'* 9/9/99 dh
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
 
Private smFWkDate As String
Private smLWkDate As String
Private imAdfCode As Integer
Private imAllClick As Integer
Private imAllVehClick As Integer
Private hmAst As Integer
Private tmStatusOptions As STATUSOPTIONS




Private Sub chkAll_Click()
 Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcContract.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcContract.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcContract.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
End Sub

Private Sub ckcAllVehicles_Click()
 Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllVehClick Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllVehClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllVehClick = False
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmClearRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles, sStations As String
    Dim sStartDate, sEndDate, sDateRange As String
    Dim sContracts As String
    Dim sStatus, sCPStatus As String    'spot status and posting status flags
    Dim sStationType As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'Dim sGenDate As String              'prepass generation date for crystal filtering
    'Dim sGenTime As String              'prepass generation time
    'Dim ilFilterBy As Integer               'required for common rtn, unused for this report
    Dim slAdjustedStartDate As String         'backup the start date to get spots pledged last week that are airing in the 1st week of request
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim ilShowExact As Integer
    ReDim ilSelectedVehicles(0 To 0) As Integer         '5-30-18

    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    'Dim NewForm As New frmViewReport
        
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
    If lbcAdvertiser.ListIndex < 0 Then
        gMsgBox "Advertiser must be specified.", vbOKOnly
        Exit Sub
    End If
    If calEffWeek.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        calEffWeek.SetFocus
        Exit Sub
    End If
    If Trim$(txtNoWeeks.Text) = "" Then
        gMsgBox "# Weeks must be specified.", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    If gIsDate(calEffWeek.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yyyy).", vbCritical
        calEffWeek.SetFocus
    Else
        smFWkDate = Format(calEffWeek.Text, "m/d/yyyy")
    End If
    
    sStartDate = calEffWeek.Text
    'date must be a monday
    If Weekday(sStartDate, vbSunday) <> vbMonday Then
        gMsgBox "Date Must be a Monday", vbOKOnly
        calEffWeek.SetFocus
        Exit Sub
    End If
    'slAdjustedStartDate = DateAdd("d", -7, sStartDate)          'backup the start date to get spots pledged last week that are airing in the 1st week of request

    
    'test for validity if  number entered
    sCode = Trim$(txtNoWeeks.Text)
    If Not IsNumeric(sCode) Then
        gMsgBox "Invalid # weeks", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    
    If lbcContract.SelCount <= 0 Then                               '12-1-12at least 1 contract must be selected
        Screen.MousePointer = vbDefault
        gMsgBox "Contracts must be selected.", vbOKOnly
        lbcContract.SetFocus
        Exit Sub
    End If
    
    dFWeek = CDate(smFWkDate)
    iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
    dLWeek = DateAdd("d", iNoWeeks, dFWeek)
    sLWeek = CStr(dLWeek)
    smLWkDate = Format$(sLWeek, "m/d/yyyy")
    
    Screen.MousePointer = vbHourglass
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)
    
    gInitStatusSelections tmStatusOptions               '3-26-12 set all options to exclude
    'all aired spots are included
    tmStatusOptions.iInclLive0 = True
    tmStatusOptions.iInclDelay1 = True
    tmStatusOptions.iInclAirOutPledge6 = True
    tmStatusOptions.iInclAiredNotPledge7 = True
    tmStatusOptions.iInclDelayCmmlOnly9 = True
    tmStatusOptions.iInclAirCmmlOnly10 = True
    tmStatusOptions.iInclMG11 = True
    tmStatusOptions.iInclBonus12 = True
    tmStatusOptions.iInclRepl13 = True
    tmStatusOptions.iInclResolveMissed = False              'exclude resolved missed spots when mgs shown, only codes shown on this report
    
    If Not ckcIncludeMiss(0).Value = vbChecked Then        'exclude not aired, not aired status already set as false
        If ckcIncludeMiss(1).Value = vbChecked Then      'exclude not aired, include not carried
            tmStatusOptions.iInclNotCarry8 = True
        End If
    Else    'include not aired
        tmStatusOptions.iInclMissed2 = True
        tmStatusOptions.iInclMissed3 = True
        tmStatusOptions.iInclMissed4 = True
        tmStatusOptions.iInclMissed5 = True
        tmStatusOptions.iInclMissedMGBypass14 = True        '4-12-17  if site has mg not allowed, include them if they exists for reported period
        If ckcIncludeMiss(1).Value = vbChecked Then 'include not carried
            tmStatusOptions.iInclNotCarry8 = True
        End If
    End If
    
    'determine option to include non-reported stations
    If ckcInclNotRecd.Value = vbChecked Then     'include non-reported (or cp not received) stations
        tmStatusOptions.iNotReported = True
    End If
    
    If Not ckcIncludeMiss(0).Value = vbChecked Then        'exclude not aired
        If ckcIncludeMiss(1).Value = vbUnchecked Then      'exclude not carried
            ilShowExact = True
        Else
            ilShowExact = False
        End If
        
    Else    'include not aired
        If ckcIncludeMiss(1).Value = vbChecked Then 'include not carried
            ilShowExact = False
        Else
            ilShowExact = True
        End If
    End If

    'mBuildAst
    ilAdvtOption = -1             'gather all advt, filter later
    ilIncludeNonRegionSpots = True      '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
    blFilterAvailNames = False
    blUseAirDAte = True                             'assume to use air date for spot filter
    blIncludePledgeInfo = True
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
                                
    
    'send the options sent into structure for general rtn
    tlSpotRptOptions.sStartDate = smFWkDate
    tlSpotRptOptions.sEndDate = smLWkDate
    tlSpotRptOptions.bUseAirDAte = blUseAirDAte
    tlSpotRptOptions.iAdvtOption = ilAdvtOption
    tlSpotRptOptions.iCreateAstInfo = True
    tlSpotRptOptions.iShowExact = ilShowExact
    tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
    tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
    tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
    tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
    tlSpotRptOptions.lContractNumber = 0            '6-4-18 no single contract option in this report
    
    '4-17-08 always exclude spots not carried, bBuildAstStnClr tests for the status to exclude them.  last parameter is a flag to show exact station feed (exclude not carried if true)
    'gBuildAstStnClr hmAst, smFWkDate, smLWkDate, iType, lbcVehicles, lbcStations, ilAdvt, True, lbcAdvertiser, sGenDate, sGenTime, True, False, ilIncludeNonRegionSpots, ilFilterBy, lbcAdvertiser   '9-26-08 get everything (not carried included), filter later
    '3-26-12 change to use general subroutine to gather and create the prepass spots; only those to be printed are created
    'gBuildAstSpotsByStatus hmAst, smFWkDate, smLWkDate, iType, lbcVehicles, lbcStations, ilAdvt, True, lbcAdvertiser, sGenDate, sGenTime, True, False, ilIncludeNonRegionSpots, ilFilterBy, lbcAdvertiser, tmStatusOptions
    '12-3-12 start date has been backed up 1 week to capture any aired spots that were pledged the previous week or that has been posted into following week
    'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, smLWkDate, iType, lbcVehicles, lbcStations, ilAdvt, True, lbcAdvertiser, sGenDate, sGenTime, True, False, ilIncludeNonRegionSpots, ilFilterBy, lbcAdvertiser, tmStatusOptions
    
    
    '2-18-14 chg to use new filedesign
    'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, smLWkDate, blUseAirDAte, lbcVehicles, lbcStations, True, lbcAdvertiser, True, False, ilIncludeNonRegionSpots, ilFilterBy, lbcAdvertiser, tmStatusOptions, blFilterAvailNames, lbcAdvertiser
    gCopySelectedVehicles lbcVehicles, ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcVehicles, lbcStations, lbcAdvertiser, lbcAdvertiser, lbcAdvertiser
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles, lbcStations, lbcAdvertiser, lbcAdvertiser, lbcAdvertiser       '5-30-18

    'CRpt1.Connect = "DSN = " & sgDatabaseName
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmClearRpt, "AfAdvClr", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex    'select the user input
        ilRptDest = 2
    '    CRpt1.Destination = crptToFile
    '    If cboFileType.ListIndex < 0 Then
    '        iType = 0
    '    Else
    '        iType = cboFileType.ItemData(cboFileType.ListIndex)
    '    End If
    '    Select Case iType
    '        Case 0, 1, 2
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.txt"
    '        Case 3
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.DIF"
    '        Case 4
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.CSV"
    '        Case 7
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.RPT"
    '        Case 10
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.xls"
    '        Case 13
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.wks"
    '        Case 15
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.RTF"
    '        Case 17
    '            CRpt1.PrintFileType = iType    'crptText
    '            CRpt1.PrintFileName = sgReportDirectory + "afAdvClr.Doc"
    '    End Select
    '    sOutput = CRpt1.PrintFileName
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    Screen.MousePointer = vbHourglass
    
    sContracts = ""
    sStatus = ""
    sCPStatus = ""
    If chkAll.Value = vbUnchecked Then
         For i = 0 To lbcContract.ListCount - 1
             If lbcContract.Selected(i) Then
                 If sContracts = "" Then
                    ' sContracts = "((lstCntrNo = " & lbcContract.List(i) & ")"
                    sContracts = " astCntrno IN ( " & Str(lbcContract.List(i))
                 Else
                     'sContracts = sContracts & " OR (lstCntrNo = " & lbcContract.List(i) & ")"
                     sContracts = sContracts & "," & Str(lbcContract.List(i))
                 End If
             End If
         Next i
        
         sContracts = sContracts & ") and"
    End If
    
'    If Not ckcIncludeMiss(0).Value = vbChecked Then        'exclude missed (not aired)
'        '1-30-07 include the new statuses considered aired (9 = pgm/coml delayed, 10 = air coml only)
'        'sStatus = " and (astStatus = 0 or astStatus = 1 or astStatus = 7 or aststatus = 9 or aststatus = 10 or astStatus = 20 or astStatus = 21) and (astStatus <> 22)  "
'        sStatus = " and ((astStatus <= 1) or (astStatus = 7) or (aststatus >= 9 and astStatus <= 21)) "    'aststatus 11-19 doesnt exist
'    Else
'        sStatus = " and ((astStatus <> 22) and (astpledgestatus < 2) or (astpledgestatus > 5 and astpledgestatus <> 8 )) "
'    End If
'
    
        
'    If Not ckcIncludeMiss(0).Value = vbChecked Then        'exclude not aired
'        '1-30-07 include the new statuses considered aired (9 = pgm/coml delayed, 10 = air coml only)
'        'sStatus = " and (astStatus = 0 or astStatus = 1 or astStatus = 7 or aststatus = 9 or aststatus = 10 or astStatus = 20 or astStatus = 21) and  (astStatus <> 22)"
'        '3-10-08 replace OR tests (slow) with AND test (faster
'        If ckcIncludeMiss(1).Value = vbUnchecked Then      'exclude not aired, exclude not carried
'            sStatus = " and ((Mod(astStatus, 100) < 2) or (Mod(astStatus, 100) > 6 ) and ((Mod(astStatus, 100) <> 8) and (astpledgestatus <> 4 and astpledgestatus <> 8)))"
'            'ilShowExact = True
'        Else            'exclude not aired, include not carried
'            sStatus = "and ((Mod(astStatus, 100) < 2) or (Mod(astStatus, 100) > 6 ))"
'        End If
'
'    Else    'include not aired
'        If ckcIncludeMiss(1).Value = vbChecked Then 'include not carried
'            sStatus = " and (Mod(astStatus, 100) <> 22)"
'        Else   'exclude not carried
'            sStatus = " and ((Mod(astStatus, 100) <> 22) and (Mod(astStatus, 100) <> 8 and (astpledgestatus <> 4 and astpledgestatus <> 8)))"
'            'ilShowExact = True
'        End If
'    End If
'
'
'    If ckcAllVehicles.Value = vbUnchecked Then    '= 0 Then                        'User did NOT select all vehicles
'        For i = 0 To lbcVehicles.ListCount - 1 Step 1
'            If lbcVehicles.Selected(i) Then
'                If Len(sVehicles) = 0 Then
'                    sVehicles = " and ((vefCode = " & lbcVehicles.ItemData(i) & ")"
'               Else
'                    sVehicles = sVehicles & " OR (vefCode = " & lbcVehicles.ItemData(i) & ")"
'                End If
'            End If
'        Next i
'        If Len(sVehicles) > 0 Then
'            sVehicles = sVehicles & ")"
'        End If
'    End If
'

     '8-4-06 this report was originally changed to use the prepass and create AFR.  But it was
     'was taken out because there was a dramatic increase in processing time to create records
     'in AFR.
     
'     SQLQuery = "SELECT astcode,astairdate,astairtime,aststatus,astcpstatus,astfeeddate,astfeedtime,astpledgedate,astpledgestarttime,astpledgeendtime, "
'     SQLQuery = SQLQuery & "lstcntrno,mktRank,mktName,ShttCallLetters,ShttTimeZone,ShttStateLic,ShttCityLic, metName,metRank, Vefname, mnfCode, mnfName "
'     SQLQuery = SQLQuery & " From ast INNER JOIN  shtt ON astShfCode = shttCode INNER JOIN  VEF_Vehicles ON astVefCode = vefCode INNER JOIN  lst ON  astLsfCode =lstCode INNER JOIN  mkt ON  shttMktCode = mktCode left outer JOIN  met ON  shttMetCode = metCode INNER JOIN  ADF_Advertisers ON lstAdfCode = adfCode "
'     SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names on AfrMissedMnfCode = mnfcode "
'
'     SQLQuery = SQLQuery & " WHERE (astAirDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
'     SQLQuery = SQLQuery & " AND " & sContracts
'     SQLQuery = SQLQuery & sCPStatus        '3-26-12 remove testings for status, only spots to be printed are now created in prepass      & sStatus      '3-30-04
'     SQLQuery = SQLQuery & sVehicles

    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    If optSortby(0).Value Then         'DMA sort by market name
        sgCrystlFormula2 = "M"
    ElseIf optSortby(1).Value = True Then      'DMA market rank
        sgCrystlFormula2 = "R"
    ElseIf optSortby(3).Value = True Then       'MSA Market name
        sgCrystlFormula2 = "S"
    ElseIf optSortby(4).Value = True Then       'MSA Market Rank
        sgCrystlFormula2 = "K"
    Else                                    'else sort by call letters
        sgCrystlFormula2 = "C"
    End If
    
    
    If optSpotAired(0).Value Then   'show exact time or DP times if "Aired outside pledge", aststatus = 0
        sgCrystlFormula3 = "P"
    Else
        sgCrystlFormula3 = "A"
    End If
    
    '12-15-09 This sql call supercedes the previous set of code to filter the records to print.
    'Changed back to prepass and select only the AFR records.
    'SQLQuery = "Select * from afr where   afrgenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))
    'SQLQuery = SQLQuery & " and (astAirDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
    'dan 8-16-11
    SQLQuery = "Select afrPledgeDate, afrPledgeStarttime, afrPledgeEndTime, astAirDate, astAirTime, astStatus, astCPStatus, astFeedTime,"
    SQLQuery = SQLQuery & "astCntrno, metName, metRank, mktName, mktRank, shttCallLetters, shttTimeZone, shttStateLic, shttCityLic, vefName "
    SQLQuery = SQLQuery & " FROM  afr INNER JOIN ast ON afrAstCode=astCode INNER JOIN shtt ON astShfCode = shttCode "
    SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles ON astVefCode =vefCode INNER JOIN mkt ON shttMktCode= mktCode LEFT OUTER JOIN met ON shttMetCode= metCode INNER JOIN ADF_Advertisers ON astAdfCode =adfCode "
    SQLQuery = SQLQuery & "where " & sContracts & "   afrgenDate = " & "'" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False)))))
    
    gUserActivityLog "E", sgReportListName & ": Prepass"

    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "afAdvClr.Rpt", "AfAdvClr"
    
'    cmdReport.Enabled = True            'give user back control to gen, done buttons
'    cmdDone.Enabled = True
'    cmdReturn.Enabled = True
    
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
    
    gUserActivityLog "S", sgReportListName & ": Clear AFR"

    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "ClearRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
        
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    gUserActivityLog "E", sgReportListName & ": Clear AFR"

    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "ClearRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmClearRpt
End Sub

Private Sub cmdSendToQueue_Click()
    '5/26/13: Report Queue. add button and handle click event
    Dim ilRet As Integer
    
    ilRet = gCreateReportQueue(frmClearRpt)
End Sub

Private Sub Form_Initialize()
    '5/26/13: Report Queue
    If igReportSource = 2 Then
        Me.Left = Screen.Width + 120
    Else
        Me.Width = Screen.Width / 1.3
        Me.Height = Screen.Height / 1.3
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
        gSetFonts frmClearRpt
        gCenterForm frmClearRpt
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim lRg As Long
    Dim lRet As Long
        
    'chkAll.Visible = False
    frmClearRpt.Caption = "Advertiser Clearance Report - " & sgClientName
    smFWkDate = ""
    smLWkDate = ""
    imAllClick = False
    imAllVehClick = False
    
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
    'force all stations selected
    lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, True, lRg)
    
    lbcVehicles.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    'force all vehicles selected
    'lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
    'lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, True, lRg)
    
    mFillAdvt
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
      
    gPopExportTypes cboFileType         '3-15-04
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    '5/26/13: Report Queue, show button
    If (bgReportQueue) Then
        cmdSendToQueue.Width = cmdReport.Width / 2 - 30
        cmdReport.Width = cmdSendToQueue.Width - 30
        cmdSendToQueue.Left = cmdReport.Left + cmdReport.Width + 60
        cmdSendToQueue.Visible = True
        cmdSendToQueue.Enabled = False
    Else
        cmdSendToQueue.Visible = False
        cmdSendToQueue.Enabled = False
    End If
    
    '5/26/13: Report Queue, allow form to finish coming up
    If igReportSource = 2 Then
        tmcQueue.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path

    Set frmClearRpt = Nothing
End Sub



Private Sub lbcAdvertiser_Click()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim sCode As String
    On Error GoTo ErrHand
    
    lbcContract.Clear
    'chkAll.Visible = False
    chkAll.Value = 0        'chged from False to 0 10-22-99

    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    If calEffWeek.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        calEffWeek.SetFocus
        Exit Sub
    End If
    If Trim$(txtNoWeeks.Text) = "" Then
        gMsgBox "# Weeks must be specified.", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    If gIsDate(calEffWeek.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        calEffWeek.SetFocus
    Else
        smFWkDate = Format(calEffWeek.Text, sgShowDateForm)
    End If
    
    'test for validity if  number entered
    sCode = Trim$(txtNoWeeks.Text)
    If Not IsNumeric(sCode) Then
        gMsgBox "Invalid # weeks", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    dFWeek = CDate(smFWkDate)
    iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
    dLWeek = DateAdd("d", iNoWeeks, dFWeek)
    sLWeek = CStr(dLWeek)
    smLWkDate = Format$(sLWeek, sgShowDateForm)
    imAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    
    SQLQuery = "SELECT DISTINCT lstCntrNo from lst"
    'SQLQuery = SQLQuery + " WHERE (adf.adfCode = lst.lstAdfCode"
    SQLQuery = SQLQuery + " WHERE ((lstLogDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery + " AND lstAdfCode = " & imAdfCode & ")"
    SQLQuery = SQLQuery + " ORDER BY lstCntrNo"
    
    Set rst = gSQLSelectCall(SQLQuery)
    'If Not rst.EOF Then
    '    chkAll.Visible = True
    'End If
    While Not rst.EOF
        lbcContract.AddItem rst!lstCntrNo  ', " & rst(1).Value & ""
        rst.MoveNext
    Wend
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "ClearRpt-lbcAdvertiser"
End Sub

Private Sub lbcContract_Click()
  If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0        'chged from False to 0 10-22-99
        imAllClick = False
    End If
End Sub

Private Sub lbcVehicles_Click()
If imAllVehClick Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = 1 Then
        imAllVehClick = True
        ckcAllVehicles.Value = 0        'chged from False to 0 10-22-99
        imAllVehClick = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0           'default to PDF
    Else
        cboFileType.Enabled = False
    End If
    
    '5/26/13: Report Queue, Enable Queue button on for Print and Send to File
    If (Index = 1) Or (Index = 2) Then
        cmdSendToQueue.Enabled = True
    Else
        cmdSendToQueue.Enabled = False
    End If
End Sub



Private Sub mFillAdvt()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvertiser.Clear
    lbcContract.Clear
    'chkAll.Value = False
    chkAll.Value = 0        'chged from False to 0 10-22-99
    'SQLQuery = "SELECT adf.adfName, adf.adfCode from ADF_Advertisers adf"
    SQLQuery = "SELECT adfName, adfCode"
    SQLQuery = SQLQuery & " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"
  
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        iFound = False
    
        If Not iFound Then
            lbcAdvertiser.AddItem rst!adfName '& ", " & rst(1).Value
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = rst!adfCode
        End If
        rst.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Clear Rpt-mFillAdvt"
End Sub


Private Sub tmcQueue_Timer()
    '5/26/13: Report Queue, generate report
    Dim ilRet As Integer
       
    tmcQueue.Enabled = False
    'Set Control values
    ilRet = gSetReportCtrls(frmClearRpt, lgReportRqtCode)
    'Start generation
    cmdReport_Click
    If igReportReturn <> 2 Then
        igReportReturn = 1
    End If
    'Indicate to frmMail (Task Loop) that the report is completed
    igReportModelessStatus = 1
    Unload frmClearRpt
End Sub
