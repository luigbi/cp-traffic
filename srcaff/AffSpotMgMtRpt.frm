VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmSpotMgmtRpt 
   Caption         =   "Affiliate Spot Management Report"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   7575
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frcSelection 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   6960
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6360
         Picture         =   "AffSpotMgmtRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Select Stations from File.."
         Top             =   1920
         Width           =   360
      End
      Begin VB.CheckBox chkIncCopyChanges 
         Caption         =   "Include Copy Changes"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3120
      End
      Begin VB.CheckBox chkInclMGBypass 
         Caption         =   "Include Bypassed Makegood Spots"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3120
      End
      Begin VB.CheckBox chkNewPage 
         Caption         =   "Skip to New Page Each Station"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2415
         Width           =   2595
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
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
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.CheckBox chkBonus 
         Caption         =   "Include Bonus Spots "
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1665
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox chkReplaceSpots 
         Caption         =   "Include Replacement Spots "
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1410
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox chkResolvMiss 
         Caption         =   "Include Madegood Missed Spots"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1155
         Value           =   1  'Checked
         Width           =   2970
      End
      Begin VB.CheckBox chkUnResolvMiss 
         Caption         =   "Include Unresolved Missed Spots"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   900
         Value           =   1  'Checked
         Width           =   2730
      End
      Begin VB.ListBox lbcStatus 
         Height          =   840
         ItemData        =   "AffSpotMgmtRpt.frx":056A
         Left            =   120
         List            =   "AffSpotMgmtRpt.frx":056C
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   1
         ItemData        =   "AffSpotMgmtRpt.frx":056E
         Left            =   5280
         List            =   "AffSpotMgmtRpt.frx":0570
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   0
         ItemData        =   "AffSpotMgmtRpt.frx":0572
         Left            =   3480
         List            =   "AffSpotMgmtRpt.frx":0574
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   5280
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Include Not Aired"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtEndTime 
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "12M"
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtStartTime 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "12M"
         Top             =   540
         Width           =   855
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1230
         ItemData        =   "AffSpotMgmtRpt.frx":0576
         Left            =   3510
         List            =   "AffSpotMgmtRpt.frx":0578
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   3225
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3510
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"AffSpotMgmtRpt.frx":057A
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   570
         Width           =   345
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Dates- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   210
         Width           =   345
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
      FormDesignHeight=   6435
      FormDesignWidth =   7575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   27
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   26
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   25
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
         Left            =   1050
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
Attribute VB_Name = "FrmSpotMgmtRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'
'   Affiliate Spot Management Report - produces list of spots that are missed (not madegood),
'   Missed and Madegood (showing its missed part and makegood part), replacement spots, and bonus spots.
'   Selectivity is by Station, Vehicle, date, time, and spot type.
'
'   The foundation of this report has been copied from Fed vs Aired report
'
'
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private hmAst As Integer

Private tmStatusOptions As STATUSOPTIONS


'Private igRptIndex As Integer      'move to global module
'
'
'        mGetStationSelection - get all the selected stations from user selection
'       <input> ilCkcAll - 0 = selected station (not all)
'               lbcListBox - list box of station (lbcVehAff or lbcSelection(0)
'       <return> SQL string selected vehicles
Function mGetStationSelection(ilCkcAll As Integer, lbcListBox As control) As String
Dim i As Integer
Dim slStr As String
    slStr = ""
    If ilCkcAll = 0 Then    'User did NOT select all vehicles
        For i = 0 To lbcListBox.ListCount - 1 Step 1
            If lbcListBox.Selected(i) Then
                If Len(slStr) = 0 Then
                    slStr = "(shttCode = " & lbcListBox.ItemData(i) & ")"
                Else
                    slStr = slStr & " OR (shttCode = " & lbcListBox.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    mGetStationSelection = slStr
        
End Function

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub CkcAll_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imckcAllIgnore Then
        Exit Sub
    End If
    If CkcAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSelection(0).ListCount > 0 Then
        imckcAllIgnore = True
        lRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imckcAllStationsIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSelection(0).ListCount > 0 Then
        imckcAllStationsIgnore = True
        lRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllStationsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdDone_Click()
    Unload FrmSpotMgmtRpt
End Sub
'
'               Affiliate Spot Management Report - produces list of spots that are missed (not madegood),
'               Missed and Madegood (showing its missed part and makegood part), replacement spots, and bonus spots.
'               Selectivity is by Station, Vehicle, date, time, and spot type.
'
'               The foundation of this report has been copied from Fed vs Aired report
'
Private Sub cmdReport_Click()
    'Dan 7/20/11 this creates 3 variants and 1 integer
'    Dim i, j, X, Y, iPos As Integer
    Dim i As Integer, j As Integer, X As Integer, Y As Integer, iPos As Integer
    Dim sCode As String
    'Dan 7/20/11 bm and sName not used.  Can't dim this way, first 4 in line become variants
'    Dim bm As Variant
    Dim sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
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
    Dim ilSelected As Integer
    Dim ilNotSelected As Integer
    Dim slStatusSelected As String
    Dim slSelection As String
    'Dim sGenDate As String
    'Dim sGenTime As String
    Dim slInputStartDate As String
    Dim slInputEndDate As String
    Dim ilShowExact As Integer
    Dim ilIncludeNotCarried As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim sUserTimes As String
    'Dim ilFilterBy As Integer           'required for common rtn, this report doesnt use
    Dim ilValidStatus As Boolean
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    Dim slAdjustedStartDate As String
    ReDim ilSelectedVehicles(0 To 0) As Integer     '5-30-18
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
    'TTP 10403 - Affiliate Spot MGMT report showing extra vehicles when run for a single vehicle
    '3/1/2022 - JW - Fix TTP 10403 - Affiliate Spot MGMT report (report now uses RecordSelection formula)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    'Dim slDate As String
    'Dim slTime As String
    'Dim slMonth As String
    'Dim slDay As String
    'Dim slYear As String
    '3/1/2022 - JW - Fix TTP 10403 - Affiliate Spot MGMT report (remove Random Date)
    'gRandomDateTime slDate, slTime, slMonth, slDay, slYear
    'sgGenDate = DateValue(slMonth & "/" & slDay & "/" & slYear)
    'sgGenTime = Format$(slTime, sgShowTimeWSecForm)
    
    'debugging only for timing tests
    'Dim sGenStartTime As String
    'Dim sGenEndTime As String
    'sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    slInputStartDate = Trim$(CalOnAirDate.Text)
    
    slInputEndDate = Trim$(CalOffAirDate.Text)
     
    If gIsDate(slInputStartDate) = False Or (Len(Trim$(slInputStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(slInputEndDate) = False Or (Len(Trim$(slInputEndDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
    

    sStartTime = txtStartTime.Text
    If (gIsTime(sStartTime) = False) Or (Len(Trim$(sStartTime)) = 0) Then   'Time not valid.
        Beep
        gMsgBox "Please enter a valid start time (h:mm:ssA/P)", vbCritical
        txtStartTime.SetFocus
        Exit Sub
    End If
    
    sEndTime = txtEndTime.Text
    If (gIsTime(sEndTime) = False) Or (Len(Trim$(sEndTime)) = 0) Then   'Time not valid.
        Beep
        gMsgBox "Please enter a valid end time (h:mm:ssA/P)", vbCritical
        txtEndTime.SetFocus
        Exit Sub
    End If
    
    
    sStr = gConvertTime(sStartTime)
    If Second(sStr) = 0 Then
        sStr = Format$(sStr, sgShowTimeWOSecForm)
    Else
        sStr = Format$(sStr, sgShowTimeWSecForm)
    End If
    llStartTime = gTimeToLong(sStr, False)
    sUserTimes = Trim$(sStr) & "-"
   
    sStr = gConvertTime(sEndTime)
    If Second(sStr) = 0 Then
        sStr = Format$(sStr, sgShowTimeWOSecForm)
    Else
        sStr = Format$(sStr, sgShowTimeWSecForm)
    End If
    llEndTime = gTimeToLong(sStr, True)
    sUserTimes = sUserTimes & Trim$(sStr)
    sgCrystlFormula1 = sUserTimes
    
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod FrmSpotMgmtRpt, "PledgeAired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
        
    'use air dates vs fed dates, need to backup the week and process extra week spot may air outside the week
   '2-18-14 ast redesign, no need to backup the date when using air dates for retrieval
   ' sStartDate = DateAdd("d", -7, sStartDate)
    blUseAirDAte = True
    blIncludePledgeInfo = True          'get the pledge info for this report
    
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    ilAdvtOption = False                    'ALL advt included, no checking for selective advt
    sStatus = ""
    sCPStatus = ""
    slStatusSelected = ""
    
   
    'A status list box is hidden and the list of statuses included by user option is then set to selected.
    'default the statuses that should be included; turn off others
    'ignore live and air delay, not carried, delay comml/prg, air cmml only
'    tmStatusOptions.iInclBonus = False
'    tmStatusOptions.iInclReplacement = False
'    tmStatusOptions.iInclResolveMissed = False
'    tmStatusOptions.iInclUnresolveMissed = False
    gInitStatusSelections tmStatusOptions
    If chkResolvMiss.Value = vbChecked Then                 'Date: 3/31/2020 check for value "True" doesn't work-needs to be compared with vbChecked
        tmStatusOptions.iInclResolveMissed = True           'include the resolved missed to show the reference when mg/replacement spots are shown
    End If
    If chkUnResolvMiss.Value = vbChecked Then
        ilValidStatus = mFindAndSet("3-Not Aired Tech Diff")
        If Not ilValidStatus Then
            Exit Sub
        End If
        ilValidStatus = mFindAndSet("4-Not Aired Blackout")
        If Not ilValidStatus Then
            Exit Sub
        End If
        ilValidStatus = mFindAndSet("5-Not Aired Other")
        If Not ilValidStatus Then
            Exit Sub
        End If
        ilValidStatus = mFindAndSet("6-Not Aired Product")
        If Not ilValidStatus Then
            Exit Sub
        End If
        If chkResolvMiss.Value = vbChecked Then            'Date: 3/31/2020 check for value "True" doesn't work-needs to be compared with vbChecked
            slStatusSelected = "Include: Resolved Misses(MG)"
        End If
        'tmStatusOptions.iInclUnresolveMissed = True
        tmStatusOptions.iInclMissed2 = True
        tmStatusOptions.iInclMissed3 = True
        tmStatusOptions.iInclMissed4 = True
        tmStatusOptions.iInclMissed5 = True
    End If
    If chkResolvMiss.Value = vbChecked Then
        ilValidStatus = mFindAndSet("12-MG")
        If Not ilValidStatus Then
            Exit Sub
        End If
        If Trim$(slStatusSelected) = "" Then
            slStatusSelected = "Include: Unresolved Misses"
        Else
            slStatusSelected = slStatusSelected & ", Unresolved Misses"
        End If
        'tmStatusOptions.iInclResolveMissed = True
        tmStatusOptions.iInclMG11 = True
    End If
    If chkReplaceSpots.Value = vbChecked Then
        ilValidStatus = mFindAndSet("14-Replacement")
        If Not ilValidStatus Then
            Exit Sub
        End If
         If Trim$(slStatusSelected) = "" Then
            slStatusSelected = "Include: Replacements"
        Else
            slStatusSelected = slStatusSelected & ", Replacements"
        End If
        'tmStatusOptions.iInclReplacement = True
        tmStatusOptions.iInclRepl13 = True
    End If
    If chkBonus.Value = vbChecked Then
        ilValidStatus = mFindAndSet("13-Bonus")
        If Not ilValidStatus Then
            Exit Sub
        End If
        If Trim$(slStatusSelected) = "" Then
            slStatusSelected = "Include: Bonus"
        Else
            slStatusSelected = slStatusSelected & ", Bonus"
        End If
        'tmStatusOptions.iInclBonus = True
        tmStatusOptions.iInclBonus12 = True
    End If
    
    If chkInclMGBypass.Value = vbChecked Then           '4-13-17 include the missed mg spots
        If Trim$(slStatusSelected) = "" Then
            slStatusSelected = "Include: Missed-MG Bypass"
        Else
            slStatusSelected = slStatusSelected & ", Missed-MG Bypass"
        End If
        tmStatusOptions.iInclMissedMGBypass14 = True
    Else
        tmStatusOptions.iInclMissedMGBypass14 = False
    End If

    tmStatusOptions.iNotReported = False            'Not reported is always excluded in this report

    'Date: 2020/3/23 check for copy change flag
    tmStatusOptions.iInclCopyChanges = False        '3/30/2020 initialize Copy Changes flag to false
    If chkIncCopyChanges.Value = vbChecked Then
        If Trim$(slStatusSelected) = "" Then
            slStatusSelected = "Include: Copy Changes"
        Else
            slStatusSelected = slStatusSelected & ", Copy Changes"
        End If
        tmStatusOptions.iInclCopyChanges = True
    End If
    
    sgCrystlFormula5 = "'" & slStatusSelected & "'"
    
    ilIncludeNotCarried = False         'exclude Not Carried Status
    
    If chkNewPage.Value = vbChecked Then
        sgCrystlFormula4 = "'Y'"
    Else
        sgCrystlFormula4 = "'N'"
    End If
       
 
    'exclude non-reported (or cp not received) stations
    '2-18-14 this is set in the status options array
    'sCPStatus = " and (ast.astCPStatus = 1 Or ast.astCPStatus = 2)"

    dFWeek = CDate(slInputStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dFWeek = CDate(slInputEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    slAdjustedStartDate = Format(slInputStartDate, "m/d/yyyy")
    'backup to Monday since all CPTTS are by week
    Do While Weekday(slAdjustedStartDate, vbSunday) <> vbMonday
        slAdjustedStartDate = DateAdd("d", -1, slAdjustedStartDate)
    Loop
    sStartDate = slAdjustedStartDate
    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    'if both missed spots and not reported are excluded, only get those spots already marked as aired
    
    If Not ilIncludeNotCarried Then
       ilShowExact = True
    Else
        ilShowExact = False
    End If
    
    ilIncludeNonRegionSpots = True          '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
                                
    blFilterAvailNames = False
    'control sent after blFilterAvailNames is n/a in this report for the general subrtn
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
    
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcVehAff, lbcSelection(0), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff
    '2-18-14 change for new design of ggetastinfo
    gCopySelectedVehicles lbcVehAff, ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcVehAff, lbcSelection(0), lbcVehAff, lbcVehAff, lbcVehAff
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(0), lbcVehAff, lbcVehAff, lbcVehAff        '5-30-18
    
    SQLQuery = "Select afrastCode,  afrISCI, AfrAdfName, AfrLinkStatus, AfrMissedMnfcode, AfrMissReplDate, AfrMissRepltime, AfrProdName, afrastcode, "
    '12/11/13: Pledge information obtained from astInfo instead of ast
    'SQLQuery = SQLQuery & "ast.astLsfCode, ast.astAirDate, ast.astAirTime, ast.astStatus, ast.astFeedDate, ast.astPledgeDate, ast.astPledgeStartTime, ast.astPledgeEndTime, ast.astStatus, ast.astcode, "
    SQLQuery = SQLQuery & "astAirDate, astAirTime, astStatus, astFeedDate, afrPledgeDate, afrPledgeStartTime, afrPledgeEndTime, astStatus, astcode, "
    SQLQuery = SQLQuery & " adfName, adfCode,"
    SQLQuery = SQLQuery & "shttCallLetters, shttcode, VefName, mktName, mnfName, mnfcode "
    'SQLQuery = SQLQuery & "ast_LinkToMG.astAirDate, ast_LinkToMG.astAirTime, ast_LinkToMiss.astAirDate, ast_LinkToMiss.astAirTime "

    
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast ast on afrastcode = ast.astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on ast.astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on ast.astvefcode = vefcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mkt on shttmktcode = mktcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names on AfrMissedMnfCode = mnfcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN cpf_Copy_Prodct_ISCI on cpfcode = astcpfcode "
    
    'air dates
    SQLQuery = SQLQuery & "WHERE (ast.astAirDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND ast.astAirDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    
    'statuses
    SQLQuery = SQLQuery & sCPStatus '& sStatus
    '12/13/13: Pledge information obtained from astInfo instead of ast
    ''use pledge times instead of air time, the pledge times must span the user entered time span to include
    SQLQuery = SQLQuery & " and (afrPledgeStartTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "' AND afrPledgeEndTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "')"
    'filter by generation date and time created
    SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
        
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfSpotMgmt.rpt", "AfSpotMgmt"
    
    'debugging only for timing tests
    'sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "SpotMgmtRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
 
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSpotDeclareRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload FrmSpotMgmtRpt
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
    gSelectiveStationsFromImport lbcSelection(0), ckcAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts FrmSpotMgmtRpt
    lacStatusDesc.FontSize = 8
    FrmSpotMgmtRpt.Caption = "Affiliate Spot Management Report- " & sgClientName
    
    gCenterForm FrmSpotMgmtRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim lRg As Long
    Dim lRet As Long
    Dim ilRet As Integer
    Dim ilHideNotCarried As Integer
    Dim lg As Long


    imChkListBoxIgnore = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried           'populate list box with hard-coded spot status codes. Do not show the
    lRg = CLng(lbcStatus.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcStatus.hwnd, LB_SELITEMRANGE, False, lRg)

    'list in the box, but default the ones selected that should be included
    
    'determine height of main (top) list box
    lbcVehAff.Height = (frcSelection.Height - chkListBox.Height - CkcAll.Height - 480) / 2
    
    chkListBox.Caption = "All Vehicles"
    chkListBox.Value = 0    'False
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    chkListBox.Value = vbChecked
    
    CkcAll.Visible = False
    lbcSelection(0).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240, lbcVehAff.Width, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240) - 120
    ckcAllStations.Caption = "All Stations"
    ckcAllStations.Value = vbUnchecked
    lbcSelection(0).Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcSelection(0).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
    lbcSelection(0).Visible = True
    ckcAllStations.Move lbcSelection(0).Left, lbcSelection(0).Top - (ckcAllStations.Height + 120)
    'TTP 9943
    cmdStationListFile.Top = lbcSelection(0).Top - (ckcAllStations.Height + 120)
    ckcAllStations.Visible = True
    
    'scan to see if any vef (vpf) are using avail names. Dont show the legend on input screen
    'if there are no vehicles using avail names
'    lacStatusDesc.Visible = False
'    For iLoop = LBound(tgVpfOptions) To UBound(tgVpfOptions) - 1
'        If tgVpfOptions(iLoop).sAvailNameOnWeb = "Y" Then
'            lacStatusDesc.Visible = True
'            Exit For
'        End If
'    Next iLoop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set FrmSpotMgmtRpt = Nothing
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    'station selection
    If imckcAllStationsIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        imckcAllStationsIgnore = True
        ckcAllStations.Value = vbUnchecked
        imckcAllStationsIgnore = False
    End If

 
End Sub
Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
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
'
'        mGetVehicleSelection - get all the selected vehicles from user selection
'       <input> ilCkcAll - 0 = selected vehicle (not all)
'               lbcListBox - list box of vehicles (lbcVehAff or lbcSelection(0)
'       <return> SQL string selected vehicles
Function mGetVehicleSelection(ilCkcAll As Integer, lbcListBox As control) As String
Dim i As Integer
Dim slStr As String
    slStr = ""
    If ilCkcAll = 0 Then    'User did NOT select all vehicles
        For i = 0 To lbcListBox.ListCount - 1 Step 1
            If lbcListBox.Selected(i) Then
                If Len(slStr) = 0 Then
                    slStr = "(vefCode = " & lbcListBox.ItemData(i) & ")"
                Else
                    slStr = slStr & " OR (vefCode = " & lbcListBox.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    mGetVehicleSelection = slStr
        
End Function
'
'           Find the Status that should be included and create the SQL call
'           <input>  slInclStatus: String to find (i.e. 3=NotAired....)
'           <return>  - true if element found in list box, else false
Public Function mFindAndSet(slInclStatus As String) As Boolean
Dim ilStatus As Integer
        mFindAndSet = False
        For ilStatus = 0 To lbcStatus.ListCount - 1 Step 1
            If lbcStatus.List(ilStatus) = Trim$(slInclStatus) Then
                lbcStatus.Selected(ilStatus) = True
                mFindAndSet = True
                Exit For
            End If
        Next ilStatus

        If Not mFindAndSet Then         'error in string, status not found
            gMsgBox "Invalid Status Requested, Call Counterpoint", vbOKOnly
            Screen.MousePointer = vbDefault
        End If
End Function
