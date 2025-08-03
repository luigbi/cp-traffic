VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAiredRpt 
   Caption         =   "Spot Clearance Report"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "AffAiredRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8760
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
      FormDesignWidth =   8760
   End
   Begin VB.Frame frcSelection 
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
      Width           =   8280
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   7560
         Picture         =   "AffAiredRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Select Stations from File.."
         Top             =   2040
         Width           =   360
      End
      Begin VB.CheckBox ckcExcludeMissedIfMG 
         Caption         =   "Exclude Missed if MG Exists"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   3000
         Width           =   2415
      End
      Begin V81Affiliate.CSI_Calendar calOffAirDate 
         Height          =   270
         Left            =   2760
         TabIndex        =   39
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   1320
         TabIndex        =   38
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         Text            =   "12/12/2022"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   915
         Width           =   1215
      End
      Begin VB.TextBox txtContract 
         Height          =   285
         Left            =   1290
         MaxLength       =   8
         TabIndex        =   27
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CheckBox ckcShowCertify 
         Caption         =   "Show Certification"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1425
         Index           =   1
         ItemData        =   "AffAiredRpt.frx":0E34
         Left            =   6360
         List            =   "AffAiredRpt.frx":0E3B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   2400
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1425
         Index           =   0
         ItemData        =   "AffAiredRpt.frx":0E44
         Left            =   4560
         List            =   "AffAiredRpt.frx":0E46
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   6360
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4560
         TabIndex        =   31
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Not Reported"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   930
         Width           =   1335
      End
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Not Aired"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   915
         Width           =   1095
      End
      Begin VB.TextBox txtEndTime 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "12M"
         Top             =   585
         Width           =   855
      End
      Begin VB.TextBox txtStartTime 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "12M"
         Top             =   585
         Width           =   855
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   1635
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3300
         Begin VB.OptionButton optSortby 
            Caption         =   "DMA Mkt Rank, Advt, Vehicle, Station"
            Height          =   255
            Index           =   5
            Left            =   75
            TabIndex        =   21
            Top             =   1425
            Width           =   3105
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "DMA Mkt Rank, Advt, Station"
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   20
            Top             =   1185
            Width           =   2985
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Vehicle, Station, Date, Time"
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   22
            Top             =   945
            Width           =   2985
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Station, Date, Time"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   705
            Width           =   2280
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Vehicle, Station, Date, Time"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   18
            Top             =   465
            Width           =   2565
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Station, Date, Time"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   225
            Value           =   -1  'True
            Width           =   2280
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1230
         ItemData        =   "AffAiredRpt.frx":0E48
         Left            =   4560
         List            =   "AffAiredRpt.frx":0E4A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   3300
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox ckcNewPage 
         Caption         =   "New page each station"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lacIncl 
         Caption         =   "Include"
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   930
         Width           =   615
      End
      Begin VB.Label lacContract 
         Caption         =   "Contract #"
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   3555
         Width           =   930
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   630
         Width           =   345
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times-Start"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Dates-Start"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   300
         Width           =   345
      End
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
      Height          =   375
      Left            =   4410
      TabIndex        =   34
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
         ItemData        =   "AffAiredRpt.frx":0E4C
         Left            =   1050
         List            =   "AffAiredRpt.frx":0E4E
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
Attribute VB_Name = "frmAiredRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'*  frmAiredRpt - List of spots aired for vehicles and/or stations
'*                if the spot doesnt exist in AST, do not go out to
'*                retrieve it from the LST.  Also, include only those
'*                spots that have been imported or posted
'*
'*  Created 7/30/03 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private hmAst As Integer
Private tmStatusOptions As STATUSOPTIONS

'
'
'        mGetStationSelection - get all the selected stations from user selection
'       <input> ilCkcAll - 0 = selected station (not all)
'               lbcListBox - list box of station (lbcVehAff or lbcSelection(1)
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
    If lbcSelection(1).ListCount > 0 Then
        imckcAllStationsIgnore = True
        lRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(1).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllStationsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmAiredRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
'    Dim i, j, X, Y, iPos As Integer
    Dim i As Integer, j As Integer, X As Integer, Y As Integer, iPos As Integer
    Dim sCode As String
 '   Dim bm As Variant
    'dan M 9-20-11 change to strings
    Dim sName As String, sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim iType As Integer
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
    'Dim sGenDate As String          'generation date and time for filtering to Crystal
    'Dim sGenTime As String
    Dim ilAdvt As Integer           'set to -1 to retrieve all advt
    Dim slAdjustedStartDate As String   'backup 1 week because of the pledged spots airing in week following its feed date
    Dim ilShowExact As Integer          'true to show exact spots aired, vs showing not aired & not carried
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim sUserTimes As String
    
    'tlSpotRepoirtOptions:
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim slSelectedAVailName As String
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    ReDim ilSelectedVehicles(0 To 0) As Integer     '5-30-18
    Dim dfEndWeek As Date                           '5-30-18
    Dim ContractRst As ADODB.Recordset
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    
    On Error GoTo ErrHand
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")             '7-10-13 use global gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
    sStartDate = Trim$(CalOnAirDate.Text)
    sEndDate = Trim$(calOffAirDate.Text)
   
    If gIsDate(sStartDate) = False Or (Len(Trim$(sStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Or (Len(Trim$(sEndDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        calOffAirDate.SetFocus
        Exit Sub
    End If
    
    'backup to Monday since all CPTTS are by week
    slAdjustedStartDate = sStartDate
    Do While Weekday(slAdjustedStartDate, vbSunday) <> vbMonday
        slAdjustedStartDate = DateAdd("d", -1, slAdjustedStartDate)
    Loop

    'slAdjustedStartDate = DateAdd("d", -7, slAdjustedStartDate)        'backup the start date to get spots pledged last week that are airing in the 1st week of request
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sStartDate = Format(sStartDate, "m/d/yyyy")
    
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
    
    sCode = Trim$(txtContract.Text)
    If Not IsNumeric(sCode) And (sCode <> "") Then
        gMsgBox "Enter valid contract number", vbOKOnly
        txtContract.SetFocus
        Exit Sub
    End If
    
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
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
    sgCrystlFormula11 = "'" & sUserTimes & "'"
   
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmAiredRpt, "Aired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    blUseAirDAte = True             'use air date vs feed date for filter
    ilAdvtOption = False            'assume to get all advt
    blIncludePledgeInfo = False      'ignore pledge info in this report
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    gInitStatusSelections tmStatusOptions               '3-27-12 set all options to exclude
    sStatus = ""
    sCPStatus = ""
     
    ilAdvtOption = False
   ' If optSortby(0).Value = True Then       'advt option
    '7-30-14 optsortby(2)  is advt option when vehicles are not stations, optsortby(3) is when vehicles are also stations
    If optSortby(2).Value = True Or optSortby(3).Value = True Or optSortby(5).Value = True Then
        If chkListBox.Value = vbUnchecked Then
            ilAdvtOption = True
        End If
    End If
        
    
    'set aired spots to always include
    tmStatusOptions.iInclLive0 = True
    tmStatusOptions.iInclDelay1 = True
    tmStatusOptions.iInclAirOutPledge6 = True
    tmStatusOptions.iInclAiredNotPledge7 = True
    tmStatusOptions.iInclDelayCmmlOnly9 = True
    tmStatusOptions.iInclAirCmmlOnly10 = True
    tmStatusOptions.iInclMG11 = True
    tmStatusOptions.iInclBonus12 = True
    tmStatusOptions.iInclRepl13 = True
    tmStatusOptions.iInclResolveMissed = True       'used when makegoods/replacemnt spots included, and need to show the reference
    
    If Not ckcIncludeMiss(0).Value = vbChecked Then        'exclude not aired
        '1-30-07 include the new statuses considered aired (9 = pgm/coml delayed, 10 = air coml only)
        'sStatus = " and (astStatus = 0 or astStatus = 1 or astStatus = 7 or aststatus = 9 or aststatus = 10 or astStatus = 20 or astStatus = 21) and  (astStatus <> 22)"
        '3-10-08 replace OR tests (slow) with AND test (faster
        If ckcIncludeMiss(1).Value = vbUnchecked Then      'exclude not carried
        '  Dan M 9-20-11 for rollback, mod must look like: Mod(astStatus, 100) = 1
'            sStatus = " and (((astStatus mod 100) < 2) or ((astStatus mod 100) > 6 and (astStatus mod 100) <> 8))"
            sStatus = " and ((Mod(astStatus,100) < 2) or ( Mod(astStatus,100) > 6 and Mod(astStatus,100) <> 8))"
            ilShowExact = True
        Else
'            sStatus = "and (((astStatus mod 100) < 2) or ((astStatus mod 100) > 6 ))"
            sStatus = "and ((Mod(astStatus,100) < 2) or (Mod(astStatus,100) > 6 ))"
            ilShowExact = False
            tmStatusOptions.iInclNotCarry8 = True
        End If
        
    Else    'include not aired (missed)
        tmStatusOptions.iInclMissed2 = True
        tmStatusOptions.iInclMissed3 = True
        tmStatusOptions.iInclMissed4 = True
        tmStatusOptions.iInclMissed5 = True
        tmStatusOptions.iInclMissedMGBypass14 = True                 '4-12-17   include missed mg bypass spots
        If ckcIncludeMiss(1).Value = vbChecked Then 'include not carried
             sStatus = " and (Mod(astStatus,100) <> ASTEXTENDED_BONUS)"
            ilShowExact = False
            tmStatusOptions.iInclNotCarry8 = True
        Else
            sStatus = " and ((Mod(astStatus,100) <> ASTEXTENDED_BONUS) and (Mod(astStatus,100) <> 8))"
            ilShowExact = True
        End If
    End If
    
    'determine option to include non-reported stations
    If Not ckcInclNotRecd.Value = vbChecked Then    'exclude non-reported (or cp not received) stations
        'sCPStatus = " and (astCPStatus = 1 Or astCPStatus = 2)"
        sCPStatus = " and (astCPStatus <> 0)"      '3-10-08
    Else
        tmStatusOptions.iNotReported = True
    End If
        
    If optSortby(0).Value = True Then          'station
        sgCrystlFormula1 = "'S'"
        iType = 0
    ElseIf optSortby(1).Value = True Then       'vehicle, station
        sgCrystlFormula1 = "'V'"
        iType = 1
    ElseIf optSortby(2).Value = True Then       'advertiser & station
        sgCrystlFormula1 = "'D'"
        iType = 2
    ElseIf optSortby(3).Value = True Then
        sgCrystlFormula1 = "'A'"        'advt, vehicle, station
        iType = 2
    ElseIf optSortby(4).Value = True Then        'dma, advt & station, date,time
        sgCrystlFormula1 = "'R'"                   'sort by rank
        iType = 2
    Else                                         'dma, advt, vehicle, station
        sgCrystlFormula1 = "'K'"
        iType = 2
    End If
    
    '6-16-04 New Page option, applicable for Vehicle sort only; to skip to new page after each station
    If ckcNewPage.Value = vbChecked Then
        sgCrystlFormula4 = "'Y'"
    Else
        sgCrystlFormula4 = "'N'"
    End If
    
    '8-2-06 show posting certifications
    If ckcShowCertify.Value = vbChecked Then
        sgCrystlFormula5 = "'Y'"
    Else
        sgCrystlFormula5 = "'N'"
    End If
    
    dFWeek = CDate(sStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    '5-30-18 use different field for end date, needed later
'    dFWeek = CDate(sEndDate)
'    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dfEndWeek = CDate(sEndDate)     '5-30-18
    sgCrystlFormula3 = "Date(" + Format$(dfEndWeek, "yyyy") + "," + Format$(dfEndWeek, "mm") + "," + Format$(dfEndWeek, "dd") + ")"
    
    '9-12-16 Exclude Missed spot if MG scheduled implies not to show the Missed part of the MG, and suppress the MG verbiage
    If ckcExcludeMissedIfMG.Value = vbChecked Then
        sgCrystlFormula12 = "'Y'"
    Else
        sgCrystlFormula12 = "'N'"
    End If

    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    
    blFilterAvailNames = False
    'control sent after blFilterAvailNames is n/a in this report for the general subrtn
    'Must always build the ast to create the prepass
    'If ckcIncludeMiss.Value = vbChecked Or ckcInclNotRecd.Value = vbChecked Then
        'mBuildAstStnClr sStartDate, sEndDate, iType, lbcVehAff
         ilAdvt = -1                'retrieve all advt
         ilIncludeNonRegionSpots = True     '7-22-10  include spots with/without regional copy
         ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
         '4-17-08 always exclude spots not carried, bBuildAstStnClr tests for the status to exclude them.  last parameter is a flag to include/exclude
         
        If tmStatusOptions.iNotReported = True Then        '3-26-15 if including not reported, need to see pledge data for those agreements not posted.  The pledge status is tested.
            blIncludePledgeInfo = True
        End If
                
        tlSpotRptOptions.sStartDate = slAdjustedStartDate 'sStartDate : needs to be monday date for cptt
        tlSpotRptOptions.sEndDate = sEndDate
        tlSpotRptOptions.bUseAirDAte = blUseAirDAte
        tlSpotRptOptions.iAdvtOption = ilAdvtOption
        tlSpotRptOptions.iCreateAstInfo = True
        tlSpotRptOptions.iShowExact = ilShowExact
        tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
        tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
        tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
        tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
        tlSpotRptOptions.lContractNumber = Val(sCode)           '6-4-18
        
        If iType = 2 Then              '7-27-06 always build by vehicle
            'gBuildAstStnClr hmAst, slAdjustedStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff  'True
            '3-27-12 use common subrtn to build all spots required in prepass.  Build only those that will be printed
            'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions 'True
            
            'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff  'True
            '2-7-14
            If Trim$(sCode) <> "" Then
                gSelectiveContract lbcSelection(0), ilSelectedVehicles(), sCode, dFWeek, dfEndWeek
            Else
                gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()         '5-30-18
            End If
'            gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff
            gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff
        ElseIf iType = 1 Then
            'gBuildAstStnClr hmAst, slAdjustedStartDate, sEndDate, iType, lbcVehAff, lbcSelection(1), ilAdvt, False, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff   'True
            '3-27-12 use common subrtn to build all spots required in prepass.  Build only those that will be printed
            'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, iType, lbcVehAff, lbcSelection(1), ilAdvt, False, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions  'True
            
            'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, blUseAirDAte, lbcVehAff, lbcSelection(1), False, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff  'True
            '2-7-14
            If Trim$(sCode) <> "" Then
                gSelectiveContract lbcVehAff, ilSelectedVehicles(), sCode, dFWeek, dfEndWeek
            Else
                gCopySelectedVehicles lbcVehAff, ilSelectedVehicles()       '5-30-18
            End If
'           gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcVehAff, lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff
           gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff
        Else
            'gBuildAstStnClr hmAst, slAdjustedStartDate, sEndDate, iType, lbcSelection(0), lbcVehAff, ilAdvt, False, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff   'True
            '3-27-12 use common subrtn to build all spots required in prepass.  Build only those that will be printed
           ' gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, iType, lbcSelection(0), lbcVehAff, ilAdvt, False, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions  'True
            
            'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcVehAff, False, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff   'True
            '2-7-14
            If Trim$(sCode) Then
                gSelectiveContract lbcSelection(0), ilSelectedVehicles(), sCode, dFWeek, dfEndWeek
            Else
                gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()         '5-30-18
            End If
            gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcVehAff, lbcVehAff, lbcVehAff, lbcVehAff
        End If
    'End If
    SQLQuery = "SELECT  afrastcode, afrISCI, afrProduct, "
    SQLQuery = SQLQuery & "astcode, astairdate, astairtime, aststatus, astcpstatus, "
    SQLQuery = SQLQuery & "vefcode, vefname, adfcode, adfname, shttcode, shttcallletters, "
    SQLQuery = SQLQuery & " vpfAllowSplitCopy, mnfCode, mnfName, cpfISCI, cpfName "
    'Dan M 9-20-11 revise sql for cr11
'    SQLQuery = SQLQuery & " FROM afr, VEF_Vehicles, shtt, ast, lst, att, vpf_Vehicle_Options,  "
'    SQLQuery = SQLQuery & " ADF_Advertisers "
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrAstCode = astCode   " _
    & "INNER JOIN ADF_Advertisers on astAdfCode = adfCode INNER JOIN  shtt on astShfCode = shttCode INNER JOIN VEF_Vehicles on astVefCode = vefCode " _
    & "INNER JOIN VPF_VEHICLE_OPTIONS on vefCode = vpfVefKCode LEFT OUTER JOIN mkt on shttMktCode = mktCode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_names on afrMissedMnfCode = mnfcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "

    'end changes for cr11
    'SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(sStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & "WHERE (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
    'SQLQuery = SQLQuery & " and (astCPStatus = 1 or astCPStatus = 2)" & sStatus
    'Dan M 9-20-11 no longer needed for cr11
    'SQLQuery = SQLQuery & " and afrAstCode = astcode " & sCPStatus & sStatus
    SQLQuery = SQLQuery & " " & sCPStatus           '& status  3-27-12 no longer need to filter the status codes, only records to be printed are created


    SQLQuery = SQLQuery & " and (astAirTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "' AND astAirTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "')"
    'SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    SQLQuery = SQLQuery & " AND (astAirDate >= '" & Format$(sStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    If txtContract <> "" Then
        SQLQuery = SQLQuery & " and (astCntrNo = " & Trim$(txtContract) & ")"
    End If
    
    ' TTP 10067 - Spot Clearance report - date/time filter stopped working
    'the issue is, AfStnClr.rpt was converted from ODBC to use Pervasive datasource (Maybe for performance?).  Which now; using Pervasive datasource Driver, Crystal doesnt support the SQL Query, and all the records aren't filtered.
    'Changes made to AfStnClr.rpt to filter the generated records based on the @StartDate, @EndDate and @UserTimes formula's
    SQLQuery = ""
        


    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfStnClr.rpt", "AfStnClr"
        
'    cmdReport.Enabled = True            'give user back control to gen, done buttons
'    cmdDone.Enabled = True
'    cmdReturn.Enabled = True
    
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
        gHandleError "AffErrorLog.txt", "AiredRpt-cmdReport_Click"
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
    gHandleError "AffErrorLog.txt", "frmAireRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAiredRpt
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
    gSelectiveStationsFromImport lbcSelection(1), ckcAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub


Private Sub Form_Activate()
'    'grdVehAff.Columns(0).Width = grdVehAff.Width
    If optSortby(0).Value Then
        optSortby_Click 0
    ElseIf optSortby(1).Value Then
        optSortby_Click 1
    End If
End Sub

Private Sub Form_Initialize()
Dim ilHalf As Integer
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.4, Screen.Height / 1.4
'    ckcShowCertify.Move ckcNewPage.Left, ckcNewPage.Top + ckcNewPage.Height + 30
'    lacContract.Move ckcShowCertify.Left, ckcShowCertify.Top + ckcShowCertify.Height + 60
'    txtContract.Move 2040, lacContract.Top - 30
    ilHalf = (frcSelection.Height - chkListBox.Height - CkcAll.Height - 120) / 2
    lbcVehAff.Move chkListBox.Left, chkListBox.Top + chkListBox.Height + 30
    lbcVehAff.Height = ilHalf
    
    gSetFonts frmAiredRpt
    gCenterForm frmAiredRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim sVehicleStn As String           '4-9-04
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    
    imChkListBoxIgnore = False
    frmAiredRpt.Caption = "Spot Clearance Report - " & sgClientName
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    End If

    ckcIncludeMiss(0).Move lacIncl.Left + lacIncl.Width, lacIncl.Top - 30
    ckcIncludeMiss(1).Move ckcIncludeMiss(0).Left + ckcIncludeMiss(0).Width, ckcIncludeMiss(0).Top
    ckcInclNotRecd.Move ckcIncludeMiss(1).Left + ckcIncludeMiss(1).Width, ckcIncludeMiss(1).Top
    
    
    '4-9-04 there are 4 different sorts, 2 are for those clients whose vehicles are also stations.  Those reports dont have any
    'subtotals by vehicles:  station/date/time or advt/station/date/time
    'the other two have subtots by vehicle (vehicle/station/date/time or advt/vehicle/station/date/time
    If sVehicleStn = "Y" Then               'vehicles are stations
        frcSortBy.Height = 975
        'ckcNewPage.Top = frcSortBy.Top + frcSortBy.Height + 120
        optSortby(0).Top = 150
        optSortby(2).Top = 390
        optSortby(4).Top = 630
        optSortby(0).Value = True           'Station, Date, Time
        optSortby(0).Visible = True
        optSortby(2).Visible = True           'Advt, Station, Date, Time
        optSortby(1).Visible = False        'vehicle, station, date, time
        optSortby(3).Visible = False        'advt, vehicle, station, date time
        'new sort added 12/28/10
        optSortby(4).Visible = True         'dma mkt rank,advt station
        optSortby(5).Visible = False        'dma mkt rank, advt, vehicle, station
        'default to show the stations
        chkListBox.Caption = "All Stations"
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcVehAff.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            End If
        Next iLoop
        
        lbcSelection(1).Visible = False
        'lbcSelection(0).Move 3315, 2400, 3555, 1425
        ckcAllStations.Caption = "All Vehicles"
        ckcAllStations.Value = vbUnchecked
        lbcSelection(0).Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcSelection(0).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgVehicleInfo(iLoop).iCode
            'End If
        Next iLoop
        lbcSelection(0).Visible = True
        lbcSelection(1).Visible = False
        'CkcAll.Move 3315, 1800
        CkcAll.Visible = True
        ckcShowCertify.Move 120, frcSortBy.Top + frcSortBy.Height + 120
        ckcExcludeMissedIfMG.Move ckcShowCertify.Left + ckcShowCertify.Width + 240, ckcShowCertify.Top
        lacContract.Move 120, ckcExcludeMissedIfMG.Top + ckcExcludeMissedIfMG.Height + 60
        txtContract.Move lacContract.Left + lacContract.Width + 240, lacContract.Top - 30
        ckcNewPage.Move 120, txtContract.Top + txtContract.Height + 30
    Else                        'vehicles are not stations
        frcSortBy.Height = 975
        'ckcNewPage.Top = frcSortBy.Top + frcSortBy.Height + 120
        optSortby(0).Visible = False
        optSortby(0).Enabled = False
        optSortby(2).Visible = False
        optSortby(2).Enabled = False
        optSortby(1).Value = True
        optSortby(1).Visible = True          'vehicle, station, date, time
        optSortby(3).Visible = True           'advt, vehicle, station, date time
        optSortby(4).Visible = False        'dma rank, advt, station
        optSortby(4).Enabled = False
        optSortby(5).Visible = True         'dma rank, advt, vehicle, station
        optSortby(1).Top = 150
        optSortby(3).Top = 390
        optSortby(5).Top = 630
        'difference is question to skip to new page each vehicle isnt shown when stations are also vehicles
        ckcShowCertify.Move 120, frcSortBy.Top + frcSortBy.Height + 120
        ckcExcludeMissedIfMG.Move ckcShowCertify.Left + ckcShowCertify.Width + 240, ckcShowCertify.Top
        lacContract.Move 120, ckcExcludeMissedIfMG.Top + ckcExcludeMissedIfMG.Height + 60
        txtContract.Move lacContract.Left + lacContract.Width + 240, lacContract.Top - 30
        ckcNewPage.Move 120, txtContract.Top + txtContract.Height + 30

    End If
    
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmAiredRpt = Nothing
End Sub

Private Sub lbcSelection_Click(Index As Integer)
 
 If Index = 0 Then                          'more vehicle or station selection
    If imckcAllIgnore Then
        Exit Sub
    End If
    If CkcAll.Value = vbChecked Then
        imckcAllIgnore = True
        CkcAll.Value = vbUnchecked
        imckcAllIgnore = False
    End If
Else                                       'station selection
    If imckcAllStationsIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        imckcAllStationsIgnore = True
        ckcAllStations.Value = vbUnchecked
        imckcAllStationsIgnore = False
    End If
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

Private Sub optSortby_Click(Index As Integer)
Dim iLoop As Integer
Dim iIndex As Integer
    
    Screen.MousePointer = vbHourglass
    
    'determine height of main (top) list box
    lbcVehAff.Height = (frcSelection.Height - chkListBox.Height - CkcAll.Height - 480) / 2
    
    If optSortby(1).Value = True Then
        chkListBox.Caption = "All Vehicles"
        chkListBox.Value = 0    'False
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
            'End If
        Next iLoop
        ckcNewPage.Visible = True       '6-16-04
        
        lbcSelection(0).Visible = False
        CkcAll.Visible = False
        lbcSelection(1).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240, lbcVehAff.Width, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240) - 120
        ckcAllStations.Caption = "All Stations"
        ckcAllStations.Value = vbUnchecked
        lbcSelection(1).Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
                End If
            End If
        Next iLoop
        lbcSelection(1).Visible = True
        ckcAllStations.Move lbcSelection(1).Left, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        'TTP 9943
        cmdStationListFile.Top = lbcSelection(1).Top - (ckcAllStations.Height + 120)
        ckcAllStations.Visible = True

    ElseIf optSortby(0).Value = True Then        'station
        ckcNewPage.Value = vbUnchecked              '6-16-04
        chkListBox.Caption = "All Stations"
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcVehAff.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgStationInfo(iLoop).iCode
                End If
            End If
        Next iLoop
        
        lbcSelection(1).Visible = False
        ckcAllStations.Visible = False
 
        lbcSelection(0).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240, lbcVehAff.Width, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240) - 120
        CkcAll.Caption = "All Vehicles"
        CkcAll.Value = vbUnchecked
        lbcSelection(0).Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcSelection(0).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgVehicleInfo(iLoop).iCode
            'End If
        Next iLoop
        lbcSelection(0).Visible = True
        CkcAll.Move lbcSelection(0).Left, lbcSelection(0).Top - (CkcAll.Height + 120)
        CkcAll.Visible = True
        
    Else
        If optSortby(2).Value = True Then           'advt,station,vehicle
            ckcNewPage.Visible = False          '6-16-04
            ckcNewPage.Value = vbUnchecked      '6-16-04
        Else
            ckcNewPage.Visible = True
        End If
        chkListBox.Caption = "All Advertisers"
        chkListBox.Value = 0    '
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                lbcVehAff.AddItem Trim$(tgAdvtInfo(iLoop).sAdvtName)
                lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgAdvtInfo(iLoop).iCode
        Next iLoop
        
        'populate stations & vehicles
        lbcSelection(0).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240, lbcVehAff.Width / 2 - 120, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240) - 120
        CkcAll.Caption = "All Vehicles"
        CkcAll.Value = vbUnchecked
        lbcSelection(0).Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcSelection(0).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgVehicleInfo(iLoop).iCode
            'End If
        Next iLoop
        lbcSelection(0).Visible = True
        CkcAll.Move lbcSelection(0).Left, lbcSelection(0).Top - (CkcAll.Height + 120) '1800
        CkcAll.Visible = True
        
        lbcSelection(1).Move lbcVehAff.Left + lbcVehAff.Width - lbcSelection(0).Width, lbcSelection(0).Top, lbcSelection(0).Width, frcSelection.Height - lbcSelection(0).Top - 120
        ckcAllStations.Caption = "All Stations"
        ckcAllStations.Value = vbUnchecked
        lbcSelection(1).Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
                End If
            End If
        Next iLoop
        lbcSelection(1).Visible = True
        ckcAllStations.Move lbcSelection(1).Left, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        ckcAllStations.Visible = True
    End If
    Screen.MousePointer = vbDefault
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
