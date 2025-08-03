VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAffiliateRpt 
   Caption         =   "Affiliate Clearance Counts"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "AffiliateRpt.frx":0000
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
      TabIndex        =   6
      Top             =   1725
      Width           =   6960
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6480
         Picture         =   "AffiliateRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Select Stations from File.."
         Top             =   1920
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1410
         TabIndex        =   7
         Top             =   240
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   -1  'True
         csi_defaultdatetype=   1
         csi_caldateformat=   1
         font            =   "AffiliateRpt.frx":0E34
         csi_daynamefont =   "AffiliateRpt.frx":0E60
         csi_monthnamefont=   "AffiliateRpt.frx":0E8E
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   1410
         TabIndex        =   9
         Top             =   615
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   -1  'True
         csi_defaultdatetype=   1
         csi_caldateformat=   1
         font            =   "AffiliateRpt.frx":0EBC
         csi_daynamefont =   "AffiliateRpt.frx":0EE8
         csi_monthnamefont=   "AffiliateRpt.frx":0F16
      End
      Begin VB.CheckBox ckcSkip4 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ckcSkip3 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox ckcSkip2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox ckcSkip1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.ComboBox cbcSort4 
         Height          =   315
         ItemData        =   "AffiliateRpt.frx":0F44
         Left            =   1080
         List            =   "AffiliateRpt.frx":0F46
         TabIndex        =   17
         Top             =   2880
         Width           =   1365
      End
      Begin VB.ComboBox cbcSort3 
         Height          =   315
         ItemData        =   "AffiliateRpt.frx":0F48
         Left            =   1080
         List            =   "AffiliateRpt.frx":0F4A
         TabIndex        =   15
         Top             =   2400
         Width           =   1365
      End
      Begin VB.ComboBox cbcSort2 
         Height          =   315
         ItemData        =   "AffiliateRpt.frx":0F4C
         Left            =   1080
         List            =   "AffiliateRpt.frx":0F4E
         TabIndex        =   13
         Top             =   1920
         Width           =   1365
      End
      Begin VB.ComboBox cbcSort1 
         Height          =   315
         ItemData        =   "AffiliateRpt.frx":0F50
         Left            =   1080
         List            =   "AffiliateRpt.frx":0F52
         TabIndex        =   11
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CheckBox ckcAdvt 
         Caption         =   "All Advertisers"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ckcStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   5280
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ListBox lbcAdvt 
         Height          =   1425
         ItemData        =   "AffiliateRpt.frx":0F54
         Left            =   3480
         List            =   "AffiliateRpt.frx":0F56
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   3315
      End
      Begin VB.ListBox lbcStations 
         Height          =   1425
         ItemData        =   "AffiliateRpt.frx":0F58
         Left            =   5280
         List            =   "AffiliateRpt.frx":0F5A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1425
         ItemData        =   "AffiliateRpt.frx":0F5C
         Left            =   3480
         List            =   "AffiliateRpt.frx":0F5E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Skip"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label lacSkip 
         Alignment       =   1  'Right Justify
         Caption         =   "Page"
         Height          =   255
         Left            =   2400
         TabIndex        =   35
         Top             =   930
         Width           =   465
      End
      Begin VB.Label lacSortSeq2 
         Caption         =   "Major to Minor:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label lacSort4 
         Caption         =   "Sort Field #4"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2940
         Width           =   1065
      End
      Begin VB.Label lacSort3 
         Caption         =   "Sort Field #3"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label lacSort2 
         Caption         =   "Sort Field #2"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label lacSort1 
         Caption         =   "Sort Field #1"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label lacSortSeq 
         Caption         =   "Enter sort sequence-"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   930
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Start Date"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Aired End Date"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   25
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   24
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   23
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
         ItemData        =   "AffiliateRpt.frx":0F60
         Left            =   1050
         List            =   "AffiliateRpt.frx":0F62
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
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
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   2130
      End
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
   End
End
Attribute VB_Name = "frmAffiliateRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmAffiliateRpt - List of spots aired for vehicles and/or stations
'*                if the spot doesnt exist in AST, do not go out to
'*                retrieve it from the LST.  Also, include only those
'*                spots that have been imported or posted
'*
'*  Created 7/30/03 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'
'*      8-11-04 Add option to select by stations
'               Fix selectivity by Advertiser
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imChkStnListBoxIgnore As Integer
Private imChkVehListBoxignore As Integer
Private imFirstTime As Integer
Private hmAst As Integer
Private tmStatusOptions As STATUSOPTIONS



Private Sub chkListBox_Click()
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
End Sub

Private Sub ckcAdvt_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAdvt.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcAdvt.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcAdvt.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcAdvt.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

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
End Sub

Private Sub cmdDone_Click()
    Unload frmAffiliateRpt
End Sub
'
'       Affiliate Clearance Counts - this report shows spot counts for Ordered,
'       Aired, or Not reported categories for one or more of the following fields:
'       Vehicle, Market, Station, Advt/Contract
'       Any combination of the above fields can be selected, and the user can
'       select their sort sequence.

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles, sStations, sAdvt, sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim iType As Integer                '12-5-12 replaced by bluseairdate
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'Dim NewForm As New frmViewReport
    Dim sStartTime As String
    Dim sEndTime As String
    Dim sCPStatus As String         '12-24-03 option to include non-reported stations
    Dim slNow As String
   ' Dim sGenDate As String
    'Dim sGenTime As String
   ' Dim ilAdvt As Integer
    
    'Dim ilFilterBy As Integer               'required for common rtn, unused for this report
    'Dim slAdjustedStartDate As String       'entered start date adjusted to get the previous week in case spots were moved across pledged weeks
    Dim ilShowExact As Integer
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    Dim sInputStartDate As String               '12-7-17
    Dim sInputEndDate As String
    ReDim ilSelectedVehicles(0 To 0) As Integer '5-30-18
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global gen date/time so it doesnt have to be passed
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)



    sStartDate = Trim$(CalOnAirDate.Text)
    sEndDate = Trim$(CalOffAirDate.Text)
    If gIsDate(sStartDate) = False Or (Len(Trim$(sStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Or (Len(Trim$(sEndDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    gUserActivityLog "S", sgReportListName & ": Prepass"
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmAffiliateRpt, "Aired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sInputStartDate = Trim$(sStartDate)             '12-7-17
    sInputEndDate = Trim$(sEndDate)                 '12-7-17

    dFWeek = CDate(sStartDate)
    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dFWeek = CDate(sEndDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"

    'slAdjustedStartDate = sStartDate
    'Do While Weekday(slAdjustedStartDate, vbSunday) <> vbMonday
    '    slAdjustedStartDate = DateAdd("d", -1, slAdjustedStartDate)
    'Loop
    Do While Weekday(sStartDate, vbSunday) <> vbMonday
        sStartDate = DateAdd("d", -1, sStartDate)
    Loop
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    sVehicles = ""
    sStations = ""
    sAdvt = ""
    sStatus = ""
    sCPStatus = ""
     

    '12-7-17 move this code so the report headings do not default to monday
'    dFWeek = CDate(sStartDate)
'    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
'    dFWeek = CDate(sEndDate)
'    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    sgCrystlFormula3 = Str$(cbcSort1.ListIndex + 1)         'sort fields, major to minor
    sgCrystlFormula4 = Str$(cbcSort2.ListIndex)
    sgCrystlFormula5 = Str$(cbcSort3.ListIndex)
    sgCrystlFormula6 = Str$(cbcSort4.ListIndex)
    
    'determine page skips for each grouping
    If ckcSkip1.Value = vbChecked Then
        sgCrystlFormula7 = "'Y'"
    Else
        sgCrystlFormula7 = "'N'"
    End If
    If ckcSkip2.Value = vbChecked Then
        sgCrystlFormula8 = "'Y'"
    Else
        sgCrystlFormula8 = "'N'"
    End If
    If ckcSkip3.Value = vbChecked Then
        sgCrystlFormula9 = "'Y'"
    Else
        sgCrystlFormula9 = "'N'"
    End If
    If ckcSkip4.Value = vbChecked Then
        sgCrystlFormula10 = "'Y'"
    Else
        sgCrystlFormula10 = "'N'"
    End If
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    'if both missed spots and not reported are excluded, only get those spots already marked as aired
    'If ckcIncludeMiss.Value = vbChecked Or ckcInclNotRecd.Value = vbChecked Then
    'get all spots so that ordered, aired and not reported can be shown
        'mBuildAstStnClr sStartDate, sEndDate, iType, lbcVehAff, lbcStations
        
     ilAdvtOption = False                            'all advt
     If ckcAdvt.Value = vbUnchecked Then
        ilAdvtOption = True
     End If
     
    ilIncludeNonRegionSpots = True      '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
    
    tmStatusOptions.iInclLive0 = True
    tmStatusOptions.iInclDelay1 = True
    tmStatusOptions.iInclMissed2 = True
    tmStatusOptions.iInclMissed3 = True
    tmStatusOptions.iInclMissed4 = True
    tmStatusOptions.iInclMissed5 = True
    tmStatusOptions.iInclAirOutPledge6 = True
    tmStatusOptions.iInclAiredNotPledge7 = True
    tmStatusOptions.iInclNotCarry8 = False
    tmStatusOptions.iInclDelayCmmlOnly9 = True
    tmStatusOptions.iInclAirCmmlOnly10 = True
    tmStatusOptions.iInclMG11 = True
    tmStatusOptions.iInclBonus12 = True
    tmStatusOptions.iInclRepl13 = True
    tmStatusOptions.iNotReported = True
    tmStatusOptions.iInclResolveMissed = False      'exclude the resolved missed in counts
    tmStatusOptions.iInclMissedMGBypass14 = True           '4-12-17 default to include the missed mg bypass spots

    blFilterAvailNames = False
    'control sent after blFilterAvailNames is n/a in this report for the general subrtn
    
    blIncludePledgeInfo = False     'no need for any pledge info into this report
    If tmStatusOptions.iNotReported = True Then        '3-26-15 if including not reported, need to see pledge data for those agreements not posted.  The pledge status is tested.
        blIncludePledgeInfo = True
    End If
    ilShowExact = True          '10-25-13 make sure Not Carried is never included
    blUseAirDAte = True        'use air dates vs feed dates

    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    'send the options sent into structure for general rtn
    tlSpotRptOptions.sStartDate = sStartDate
    tlSpotRptOptions.sEndDate = sEndDate
    tlSpotRptOptions.bUseAirDAte = blUseAirDAte
    tlSpotRptOptions.iAdvtOption = ilAdvtOption
    tlSpotRptOptions.iCreateAstInfo = True
    tlSpotRptOptions.iShowExact = ilShowExact
    tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
    tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
    tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
    tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
    tlSpotRptOptions.lContractNumber = 0            '6-4-18 no single contract option in this report

    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, iType, lbcVehAff, lbcStations, ilAdvt, ilAdvtOption, lbcAdvt, sGenDate, sGenTime, True, False, ilIncludeNonRegionSpots, ilFilterBy, lbcVehAff, tmStatusOptions
    'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, blUseAirDAte, lbcVehAff, lbcStations, ilAdvtOption, lbcAdvt, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff
    '2-19-14 change for new design and additional parameters
    gCopySelectedVehicles lbcVehAff, ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcVehAff, lbcStations, lbcAdvt, lbcVehAff, lbcVehAff
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcStations, lbcAdvt, lbcVehAff, lbcVehAff

    'Dan M 8/3/11  use crystal mod, not sql mod
    'Dan M 9/19/11 revert to sql mod for cr11
    'sStatus = " and (Mod(astStatus,100) <> 22 and Mod(astStatus,100) <> 8)"     'ignore the missed portion of a makegood and not spots not carried
    
    'sStatus = " and ( (astStatus = 0 or astStatus = 1 or astStatus >= 9) or ((afrLinkStatus = '' ) and (astStatus >= 2 or aststatus <= 5 )) ) "
    
    'SQLQuery = "SELECT * "
    'SQLQuery = SQLQuery & " FROM VEF_Vehicles, shtt, ast, lst, att, webl, "
    'Dan 8-12-09 att removed from sql call
    'SQLQuery = SQLQuery & " FROM afr, VEF_Vehicles, shtt, ast, cpf, mkt, "
    'SQLQuery = SQLQuery & " ADF_Advertisers "
    
    SQLQuery = "Select astStatus, astCPstatus, astCntrno, astCpfCode, CpfName, adfname, mktName,shttCallLetters, vefName FROM afr  "
    SQLQuery = SQLQuery & " INNER JOIN ast on afrastcode = astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on astvefcode = vefcode "
    SQLQuery = SQLQuery & " INNER JOIN mkt on shttmktcode = mktcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "
    
'    SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(sStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    '12-7-17 use the input start date vs start of week
    SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(sInputStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sInputEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
    'SQLQuery = SQLQuery & sCPStatus & sStatus          '2-19-14 should be filtered out by the gBuildAst general rtn
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    'SQLQuery = SQLQuery & sAdvt
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfCounts.rpt", "AfCounts"
     
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime

    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "AffiliateRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    Screen.MousePointer = vbDefault

        
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffiliateRpt-cmdReport_Click"
    Exit Sub
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAffiliateRpt
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

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
    If imFirstTime = True Then
        mPopSorts cbcSort1, False           'Major group total #1, dont allow NONE for a choice
        mPopSorts cbcSort2, True           ' group total #2, allow  NONE for a choice
        mPopSorts cbcSort3, True           ' group total #3, allow  NONE for a choice
        mPopSorts cbcSort4, True           ' minor group total #4,  allow NONE for a choice
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
Dim ilHalf As Integer
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    ckcAdvt.Top = 240
    
    ilHalf = (Frame2.Height - ckcAdvt.Height - chkAllVehicles.Height - 120) / 2
    lbcAdvt.Move ckcAdvt.Left, ckcAdvt.Top + ckcAdvt.Height
    lbcAdvt.Height = ilHalf
    lbcVehAff.Height = ilHalf
    lbcStations.Height = ilHalf
    chkAllVehicles.Top = lbcAdvt.Top + lbcAdvt.Height
    ckcStations.Top = chkAllVehicles.Top
    lbcVehAff.Top = chkAllVehicles.Top + chkAllVehicles.Height
    lbcStations.Top = lbcVehAff.Top
    lbcVehAff.Height = ilHalf
    lbcStations.Height = ilHalf


    gSetFonts frmAffiliateRpt
    gCenterForm frmAffiliateRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim ilRet As Integer
    Dim lRg As Long
    Dim lRet As Long
    
    imFirstTime = True
    imChkListBoxIgnore = False
    imChkVehListBoxignore = False
    imChkStnListBoxIgnore = False
    frmAffiliateRpt.Caption = "Affiliate Clearance Counts Report - " & sgClientName
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
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
    'lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
    'lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, True, lRg)
    
    lbcAdvt.Clear
    For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
        lbcAdvt.AddItem Trim$(tgAdvtInfo(iLoop).sAdvtName)
        lbcAdvt.ItemData(lbcAdvt.NewIndex) = tgAdvtInfo(iLoop).iCode
    Next iLoop
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmAffiliateRpt = Nothing
End Sub

Private Sub grdVehAff_Click()
    If chkAllVehicles.Value = 1 Then
        imChkVehListBoxignore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = 0    'chged from false to 0 10-22-99
        imChkVehListBoxignore = False
    End If
End Sub

Private Sub lbcAdvt_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAdvt.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        ckcAdvt.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub lbcStations_Click()
    If imChkStnListBoxIgnore Then
        Exit Sub
    End If
    If ckcStations.Value = 1 Then
        imChkStnListBoxIgnore = True
        ckcStations.Value = 0    'chged from false to 0 10-22-99
        imChkStnListBoxIgnore = False
    End If
End Sub
Private Sub lbcVehAff_Click()
    If imChkVehListBoxignore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = 1 Then
        imChkVehListBoxignore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = 0    'chged from false to 0 10-22-99
        imChkVehListBoxignore = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

'
'           Populate the drop down with the valid fields to sort:  Vehicle, Market, station and Advertiser (advt/contr# implied)
'           <input> cbcControl as dropdown control
'                   ilShowNone : true - show None as a choice, else false to default to 1st element
'           DH 7-12-04
Public Sub mPopSorts(cbcControl As control, ilShowNone As Integer)
    If ilShowNone Then
        cbcControl.AddItem "None"
    End If
    cbcControl.AddItem "Vehicle"
    cbcControl.AddItem "DMA Market"
    cbcControl.AddItem "Station"
    cbcControl.AddItem "Advertiser"
    cbcControl.ListIndex = 0
    
   
End Sub
