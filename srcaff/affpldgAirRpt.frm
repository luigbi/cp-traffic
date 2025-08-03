VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPldgAirRpt 
   Caption         =   "Pledge vs Aired Report"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
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
      Height          =   4500
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   7200
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   4800
         Picture         =   "affpldgAirRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Select Stations from File.."
         Top             =   1920
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   420
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
         CSI_DefaultDateType=   1
      End
      Begin VB.CheckBox chkShowExact 
         Caption         =   "Show Exact Station Feed"
         Height          =   390
         Left            =   2400
         TabIndex        =   49
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Station"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Top             =   2280
         Width           =   975
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   420
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
         CSI_DefaultDateType=   1
      End
      Begin VB.CheckBox chkStatusDiscrep 
         Caption         =   "Status Discrepancy*"
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   2520
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ListBox lbcAvailNames 
         Height          =   1230
         ItemData        =   "affpldgAirRpt.frx":056A
         Left            =   5400
         List            =   "affpldgAirRpt.frx":056C
         TabIndex        =   42
         Top             =   480
         Width           =   1635
      End
      Begin VB.CheckBox chkSuppressCounts 
         Caption         =   "Suppress Spot Count Totals"
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   2520
         Width           =   1905
      End
      Begin VB.Frame frcUseDates 
         BorderStyle     =   0  'None
         Caption         =   "Time Sort by"
         Height          =   225
         Left            =   1140
         TabIndex        =   38
         Top             =   120
         Width           =   1635
         Begin VB.OptionButton optUseDates 
            Caption         =   "Aired"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   14
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optUseDates 
            Caption         =   "Fed "
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.CheckBox ckcShowStatus 
         Caption         =   "Show Status Codes"
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   2755
         Width           =   1695
      End
      Begin VB.Frame frcTimeSort 
         Caption         =   "Date/Time Sort"
         Height          =   1170
         Left            =   2280
         TabIndex        =   33
         Top             =   1080
         Width           =   1275
         Begin VB.OptionButton optTimeSort 
            Caption         =   "Fed Date/Time"
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   20
            Top             =   180
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optTimeSort 
            Caption         =   "Air Date/Time"
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   21
            Top             =   420
            Width           =   1125
         End
      End
      Begin VB.ListBox lbcStatus 
         Height          =   840
         ItemData        =   "affpldgAirRpt.frx":056E
         Left            =   120
         List            =   "affpldgAirRpt.frx":0570
         MultiSelect     =   2  'Extended
         TabIndex        =   26
         Top             =   3000
         Width           =   3195
      End
      Begin VB.CheckBox ckcSeparate 
         Caption         =   "Separate Status Codes"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2755
         Width           =   2055
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Network"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   23
         Top             =   2280
         Width           =   975
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   1
         ItemData        =   "affpldgAirRpt.frx":0572
         Left            =   5400
         List            =   "affpldgAirRpt.frx":0574
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   0
         ItemData        =   "affpldgAirRpt.frx":0576
         Left            =   3720
         List            =   "affpldgAirRpt.frx":0578
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtEndTime 
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "12M"
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtStartTime 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "12M"
         Top             =   780
         Width           =   780
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   1170
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2085
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Station, Vehicle, Date/Time"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   50
            Top             =   900
            Width           =   1965
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Station, Vehicle,  Date/Time"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   48
            Top             =   660
            Width           =   1995
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Vehicle, Station, Date/Time"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   1995
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Vehicle, Station, Date/Time"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   180
            Value           =   -1  'True
            Width           =   1965
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1230
         ItemData        =   "affpldgAirRpt.frx":057A
         Left            =   3720
         List            =   "affpldgAirRpt.frx":057C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   1635
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lacDiscrep 
         Caption         =   "Non-Compliant Only"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2300
         Width           =   1335
      End
      Begin VB.Label lacStatusDiscrep 
         Caption         =   "*Status Discrepancy includes All Statuses"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   45
         Top             =   4200
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacAvailNames 
         Caption         =   "Avail Names"
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lacSortBy 
         Caption         =   "Sort by-"
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Use Dates"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   150
         Width           =   855
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"affpldgAirRpt.frx":057E
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2220
         TabIndex        =   18
         Top             =   810
         Width           =   465
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Dates- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   2220
         TabIndex        =   10
         Top             =   450
         Width           =   465
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
      FormDesignHeight=   6240
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
      Height          =   375
      Left            =   4410
      TabIndex        =   35
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
Attribute VB_Name = "frmPldgAirRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmPldgAirRpt - Create a report of spots (from AST) for station/
'*  vehicle for requested date & time spans which compares the air time against
'*  the pledged start/end times (Pledge vs Air Clearance), or compares
'*  the fed time against the air time (Fed vs Air Clearance).
'*  Options to show discreps only or all; show status codes or not;
'   select status codes or all; included not reported stations or not.
'*
'   Create a prepass file in AFR which only has a pointer to the AST file
'   All spots for the vehicle are created and the filtering of spots is
'   processed thru the sql call to Crystal.
'
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private hmAst As Integer
Private tmStatusOptions As STATUSOPTIONS
Private imListBoxWidth As Integer

'Private igRptIndex As Integer      'move to global module


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

Private Sub ckcDiscrep_Click(Index As Integer)
    If ckcDiscrep(0).Value = vbChecked Or ckcDiscrep(1).Value = vbChecked Then        'discrepancy option
        lbcStatus.Enabled = False
    Else
        lbcStatus.Enabled = True
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmPldgAirRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    'Dan 7/20/11 this creates 3 variants and 1 integer
'    Dim i, j, X, Y, iPos As Integer
    Dim i As Integer, j As Integer, X As Integer, Y As Integer, iPos As Integer
    Dim sCode As String
    Dim sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    'Dim iType As Integer
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
    'Dim slDiscrepOnly As String      '6-26-06 selectivity for discreps only
    Dim blNetworkDiscrep As Boolean     '8-4-14
    Dim blStationDiscrep As Boolean     '8-4-14
    Dim ilSelected As Integer
    Dim ilNotSelected As Integer
    Dim slStatusSelected As String
    Dim slStatusNotSelected As String
    Dim slSelection As String
    'Dim sGenDate As String
    'Dim sGenTime As String
    Dim llLoopAST As Long
    Dim slInputStartDate As String
    Dim slInputEndDate As String
    Dim slAdjustedStartDate As String
    Dim ilAdvt As Integer               'set to -1 to retrieve all advt
    Dim ilShowExact As Integer
    Dim ilIncludeNotCarried As Integer
    Dim ilIncludeNotReported As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim sUserTimes As String
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim slSelectedAVailName As String
    Dim lRg As Long
    Dim lRet As Long
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    ReDim ilSelectedVehicles(0 To 0) As Integer         '5-30-18
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)

    
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
    sgCrystlFormula11 = sUserTimes
      
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmPldgAirRpt, "PledgeAired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    
    sStartDate = Format(slInputStartDate, "m/d/yyyy")
    'backup to Monday since all CPTTS are by week
    slAdjustedStartDate = sStartDate
    Do While Weekday(slAdjustedStartDate, vbSunday) <> vbMonday
        slAdjustedStartDate = DateAdd("d", -1, slAdjustedStartDate)
    Loop

    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    
    If optUseDates(1).Value = True Then     'use air dates vs fed dates, need to backup the week and process extra week
                                            'spot may air outside the week
         'sStartDate = DateAdd("d", -7, sStartDate)     '2-5-14 with new keys, no need to backup the dates when using air dates
         sgCrystlFormula9 = "'A'"
         blUseAirDAte = True
    Else
        sgCrystlFormula9 = "'F'"
        blUseAirDAte = False
    End If
    
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    sStatus = ""
    sCPStatus = ""
    'slDiscrepOnly = ""
    slStatusSelected = ""
    slStatusNotSelected = ""
    
    'determine option to include non-reported stations
    If Not ckcInclNotRecd.Value = vbChecked Then    'exclude non-reported (or cp not received) stations
        sCPStatus = " and (astCPStatus = 1 Or astCPStatus = 2)"
        'sgCrystlFormula6 = sgCrystlFormula6 & ", Not Reported"
        ilIncludeNotReported = False
    Else
        ilIncludeNotReported = True
        'sgCrystlFormula6 = sgCrystlFormula6 & ", Not Reported"
    End If
    
    gInitStatusSelections tmStatusOptions               '3-14-12 set all options to exclude before interrogating the list box of selections
    If chkStatusDiscrep.Value = vbChecked Then          '12-10-13 show status discreps only (aststatus vs astpledgedstatus)
        tmStatusOptions.bStatusDiscrep = True           'force Not Carried on
        lRg = CLng(lbcStatus.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStatus.hwnd, LB_SELITEMRANGE, True, lRg)
    End If
    
    'format the sql query for the selection of spot statuses
    'get the description of spot statuses (included/excluded) to show on report
    gGetSQLStatusForCrystal lbcStatus, sStatus, slSelection, ilIncludeNotCarried, ilIncludeNotReported
    sgCrystlFormula6 = slSelection
   
    
    ' Detrmine what to sort by
    ilAdvtOption = False            'assume to get all advt
    
    'this report has done away with testing if vehicles are also stations (like the Spot Clearance).
    'So instead of 4 options, there are only 2
    If optSortby(0).Value = True Then       'vehicle
        sgCrystlFormula1 = "'V'"
        'iType = 1
    ElseIf optSortby(2).Value = True Then        '4-13-17 sort by station
        sgCrystlFormula1 = "'S'"
    ElseIf optSortby(3).Value = True Then       '4-13-17  sort by advt,station,vehicle
        sgCrystlFormula1 = "'D'"
        If chkListBox.Value = vbUnchecked Then
            ilAdvtOption = True
        End If
    Else
        sgCrystlFormula1 = "'A'"        'advt & vehicle
       ' iType = 2
        If chkListBox.Value = vbUnchecked Then
            ilAdvtOption = True
        End If
    End If
    
    If optTimeSort(0).Value = True Then       'Pledge End Time, Pledge Start Time
        sgCrystlFormula7 = "'P'"

    Else
        sgCrystlFormula7 = "'A'"        'air time
'        If chkListBox.Value = vbUnchecked Then
'            ilAdvtOption = True
'        End If
    End If
    
'    gInitStatusSelections tmStatusOptions               '3-14-12 set all options to exclude; 12-11-13 move to initialize earlier in code
    gSetStatusOptions lbcStatus, ilIncludeNotReported, tmStatusOptions
    tmStatusOptions.iInclResolveMissed = True           'show the missed reference if including mg/replacements
    
    If ckcSeparate.Value = vbChecked Then            'Separate the statuses in output
        sgCrystlFormula5 = "'Y'"
    Else
        sgCrystlFormula5 = "'N'"
    End If
    
    If ckcShowStatus.Value = vbChecked Then            'show the status codes on the report
        sgCrystlFormula8 = "'Y'"
    Else
        sgCrystlFormula8 = "'N'"
    End If
    
    '5-21-12 show or suppress the spot counts
    If chkSuppressCounts.Value = vbChecked Then
        sgCrystlFormula12 = "'Y'"
    Else
        sgCrystlFormula12 = "'N'"
    End If

    
    'dFWeek = CDate(sStartDate)
    dFWeek = CDate(slInputStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    'dFWeek = CDate(sEndDate)
    dFWeek = CDate(slInputEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    'if both missed spots and not reported are excluded, only get those spots already marked as aired
    
'    If ckcIncludeMiss.Value = vbChecked Or ckcInclNotRecd.Value = vbChecked Then
     ilAdvt = -1                            'all advt
    sgCrystlFormula10 = ""
    '4-17-08 Pledge vs Air and Fed vs Air will always exclude spots not carried.
    '        Question on screen  Show Exact Station Feed has been hidden and defaulted to Yes
    
    '9-18-08 Used to have a question to show exact times which excluded Not Carried spots.
    'But we had to remove that and give user option to see those Not Carried spots.
    'In order to include/exclude any of the statuses, it will use test the selection of
    'statuses from the list box
    If igRptIndex = PLEDGEVSAIR_RPT Then
        blIncludePledgeInfo = True      'pledge vs aired needs to show pledge info
         'ilShowExact = True          '4-17-08
         If ilIncludeNotCarried Then       'user selected Not Carried?
            ilShowExact = False            '9-18-08
        Else
            ilShowExact = True
        End If
    Else
        blIncludePledgeInfo = False     'fed vs aired doesnt need pledge info
        If tmStatusOptions.iNotReported = True Then        '3-26-15 if including not reported, need to see pledge data for those agreements not posted.  The pledge status is tested.
            blIncludePledgeInfo = True
        End If


        '9-18-08 ckcShowExact value has been defaulted unchecked (to allow for the list of statuses to
        'be used instead for selection of Not Carried )
        'If chkShowExact.Value = vbChecked Then
        If Not ilIncludeNotCarried Then
           ilShowExact = True
           sgCrystlFormula10 = "'Y'"
        Else
            ilShowExact = False
            sgCrystlFormula10 = "' '"
        End If
    End If
    
    blNetworkDiscrep = False
    blStationDiscrep = False
    If ckcDiscrep(0).Value = vbChecked Or ckcDiscrep(1).Value = vbChecked Then             'discreps only
        sgCrystlFormula4 = "'Y'"
        '8-4-14 check only flags in common routine to see if non-compliant
        If ckcDiscrep(0).Value = vbChecked Then
            blNetworkDiscrep = True
        End If
        If ckcDiscrep(1).Value = vbChecked Then
            blStationDiscrep = True
        End If
'        If igRptIndex = PLEDGEVSAIR_RPT Then
'             slDiscrepOnly = " and ( (((astAirTime < afrPledgeStartTime) or (astAirTime > afrPledgeEndTime and afrPledgeEndTime <> '00:00:00')) and (mod(astStatus, 100) <= 1 or mod(aststatus ,100) = 6 or mod(aststatus, 100) = 7 or mod(aststatus , 100) = 9 or mod(aststatus, 100) = 10)) or (afrPledgeDate <> astAirDate) or ((mod(aststatus , 100) >= 2 and mod(aststatus ,100) <= 5) or (mod(aststatus ,100) = 8)) )"
'        Else                                        'fed vs aired
'             blIncludePledgeInfo = True      'fed vs aired needs to test pledge info for discrepancy version
'            'ggetastinfo sets the network and station compliant flags.  Reports only test the flags
'            slDiscrepOnly = " and ( ( ((astAirTime < afrPledgeStartTime) or (astAirTime > afrPledgeEndTime and afrPledgeEndTime <> '00:00:00')) and (mod(astStatus, 100) <= 1 or mod(aststatus , 100) = 6 or mod(aststatus , 100) = 7 or mod(aststatus , 100) = 9 or mod(aststatus , 100) = 10)) or (afrPledgeDate <> astAirDate)   or ((mod(aststatus, 100) >= 2 and mod(aststatus , 100) <= 5) or (mod(aststatus , 100) = 8)) )"
'        End If
    Else
        sgCrystlFormula4 = "'N'"                      'All
        
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
                                
    ilIncludeNonRegionSpots = True          '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
    blFilterAvailNames = False
    sgCrystlFormula13 = "'All Avail Names'"
    If lbcAvailNames.ListIndex > 0 Then           '1st element is All Avails, if not selected then need to test for the matching avail name in spot
        blFilterAvailNames = True
        For i = 0 To lbcAvailNames.ListCount - 1 Step 1
            If lbcAvailNames.Selected(i) Then
                sgCrystlFormula13 = "'" & Trim$(lbcAvailNames.List(i)) & " Avail Name'"
                Exit For
            End If
        Next i
    End If
    
    'send the options sent into structure for general rtn
    tlSpotRptOptions.sStartDate = slAdjustedStartDate   'sStartDate, needs to be a Monday for cptt
    tlSpotRptOptions.sEndDate = sEndDate
    tlSpotRptOptions.bUseAirDAte = blUseAirDAte
    tlSpotRptOptions.iAdvtOption = ilAdvtOption
    tlSpotRptOptions.iCreateAstInfo = True
    tlSpotRptOptions.iShowExact = ilShowExact
    tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
    tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
    tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
    tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
    tlSpotRptOptions.bNetworkDiscrep = blNetworkDiscrep
    tlSpotRptOptions.bStationDiscrep = blStationDiscrep
    tlSpotRptOptions.lContractNumber = 0            '6-4-18 no single contract option in this report
    
    Screen.MousePointer = vbHourglass

    If optSortby(0).Value = True Or optSortby(2).Value = True Then      'vehicle or station (4-13-17)
        '7-27-06 change from local build rtn to common rtn
        '4-17-08 always exclude spots not carried, bBuildAstStnClr tests for the status to exclude them.  last parameter is a flag to show exact station feed (exclude not carried if true)
        'gBuildAstStnClr hmAst, sStartDate, sEndDate, iType, lbcVehAff, lbcSelection(1), ilAdvt, False, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff
        '3-14-12
        'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcVehAff, lbcSelection(1), False, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcAvailNames
        '2-5-14 use a structure to pass many rpt option variables
        gCopySelectedVehicles lbcVehAff, ilSelectedVehicles()       '5-30-18
'        gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcVehAff, lbcSelection(1), lbcVehAff, lbcVehAff, lbcAvailNames
        gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcVehAff, lbcAvailNames     '5-30-18
    Else
        'gBuildAstStnClr hmAst, sStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff
        'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcAvailNames
        '2-5-14 use a structure to pass many rpt option variables
        gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()       '5-30-18
'        gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcVehAff, lbcVehAff, lbcAvailNames
        gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcVehAff, lbcAvailNames     '5-30-18
    End If
    
    
    
'    End If
        
    On Error GoTo ErrHand
    
    SQLQuery = "Select afrastCode, "
    SQLQuery = SQLQuery & "ast.astAtfCode,  ast.astAirDate, ast.astAirTime, ast.astStatus, ast.astCPStatus, ast.astlkastcode, ast.astfeeddate, ast.astfeedtime, ast_miss4mg.astcode, ast_miss4mg.astfeeddate, ast_miss4mg.astfeedtime, afrPledgeDate, afrPledgeStartTime, afrPledgeEndTime, afrPledgeStatus, "
    SQLQuery = SQLQuery & "shttCallLetters, adfName, VefName, vpfAllowSplitCopy, "
    SQLQuery = SQLQuery & " mnfName, mnfcode, cpfISCI "

    
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrastcode = ast.astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on ast.astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on ast.astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on ast.astvefcode = vefcode "
    SQLQuery = SQLQuery & " INNER JOIN vpf_Vehicle_Options on vefcode = vpfvefkcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names on AfrMissedMnfCode = mnfcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on ast.astcpfcode = cpfcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN ast ast_miss4mg on ast.astlkastcode = ast_miss4mg.astcode "

    If optUseDates(1).Value = True Then             'use air dates vs fed dates
        SQLQuery = SQLQuery & "WHERE (ast.astAirDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND ast.astAirDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    Else
        SQLQuery = SQLQuery & "WHERE (ast.astFeedDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND ast.astFeedDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    End If
    
    SQLQuery = SQLQuery & " and (ast.astAirTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "' AND ast.astAirTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "')"
    'SQLQuery = SQLQuery & slDiscrepOnly
    SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
        

    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    lgCPCount = lgCPCount
    If igRptIndex = PLEDGEVSAIR_RPT Then
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfPldgAir.rpt", "AfPldgAir"
    Else                'fed vs air has status discrepancy option:  if pledge status doesnt coincide with status option its considered a status discrepancy
        If tmStatusOptions.bStatusDiscrep = True Then               '12-11-13
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfFedAirStatusDiscrp.rpt", "AfFedAirStatusDiscrp"
        Else
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfFedAir.rpt", "AfFedAir"
        End If
    End If
    
    'debugging only for timing tests
    'sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
    gUserActivityLog "S", sgReportListName & ": Clear AFR"
    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "PldgAirRpt-cmdReport_Click"
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
    gHandleError "", "frmPldgAirRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmPldgAirRpt
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
    gSelectiveStationsFromImport lbcSelection(1), CkcAll, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Activate()
'    'grdVehAff.Columns(0).Width = grdVehAff.Width
'    imListBoxWidth = lbcVehAff.Width
'    If optSortby(0).Value Then
'        optSortby_Click 0
'    End If
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts frmPldgAirRpt
    'lacStatusDiscrep.FontBold = False
    lacStatusDiscrep.FontSize = 8
    lacStatusDiscrep.FontName = "Arial Narrow"

    lacStatusDesc.FontSize = 8
    lacStatusDesc.FontName = "Arial Narrow"
    lacStatusDesc.Move lbcStatus.Left, lbcStatus.Top + lbcStatus.Height + 240

    If igRptIndex = PLEDGEVSAIR_RPT Then
        frmPldgAirRpt.Caption = "Pledged vs Aired Clearance Report - " & sgClientName
        chkSuppressCounts.Visible = True        '.rpt will hide the subtotal spot count subtotals
        lacStatusDiscrep.Visible = False
    Else
        frmPldgAirRpt.Caption = "Fed vs Aired Clearance Report- " & sgClientName
        chkSuppressCounts.Visible = False
    End If
    
    imListBoxWidth = lbcVehAff.Width
    If optSortby(0).Value Then
        optSortby_Click 0
    End If
    
    gCenterForm frmPldgAirRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim lRg As Long
    Dim lRet As Long
    Dim ilRet As Integer
    Dim ilHideNotCarried As Integer
    Dim blIncludeAllAvailNames As Boolean

    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    igRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    imChkListBoxIgnore = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    End If
    

    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True
    gPopSpotStatusCodes lbcStatus, ilHideNotCarried           'populate list box with hard-coded spot status codes
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried         '3-14-12 populate list box with hard-coded extended spot status codes
    gPopVff
    
    'scan to see if any vef (vpf) are using avail names. Dont show the legend on input screen
    'if there are no vehicles using avail names
    lacStatusDesc.Visible = False
    For iLoop = LBound(tgVpfOptions) To UBound(tgVpfOptions) - 1
        If tgVpfOptions(iLoop).sAvailNameOnWeb = "Y" Then
            lacStatusDesc.Visible = True
            Exit For
        End If
    Next iLoop
    
    '12-10-13 add discrepancy only option for Fed vs Aired report to be able to show astStatus inconsistencies with astpledged statuses
    If igRptIndex = PLEDGEVSAIR_RPT Then     'Pledged vs Air
        optTimeSort(0).Caption = "Pledged Date/Time"
        optTimeSort(1).Top = 600
        chkShowExact.Visible = False
        'chkShowExact.Value = vbUnchecked
        chkShowExact.Value = vbChecked        '4-17-08 force to exclude spots not carried
        chkStatusDiscrep.Value = vbUnchecked
    Else                                 'fed vs air
        optTimeSort(0).Caption = "Fed Date/Time"
        'chkShowExact.Visible = True
        chkStatusDiscrep.Move chkSuppressCounts.Left
        chkStatusDiscrep.Visible = True
        chkStatusDiscrep.Value = vbUnchecked            '12-10-13
    End If
    
    If lacStatusDesc.Visible = False Then
        lbcStatus.Height = 1440
    End If
    blIncludeAllAvailNames = True
    gPopAndSortAvailNames blIncludeAllAvailNames, lbcAvailNames

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmPldgAirRpt = Nothing
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
    
    If optSortby(0).Value = True Or optSortby(2).Value = True Then      '4-13-17 add sort option by Station
        chkListBox.Caption = "All Vehicles"
        chkListBox.Value = 0    'False
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
                If igRptIndex = PLEDGEVSAIR_RPT Then       '5-21-12  set the station selected if used in pledged vs air report
                    iIndex = gBinarySearchVff(tgVehicleInfo(iLoop).iCode)
                    If iIndex > 0 Then
                        If tgVffInfo(iIndex).sPledgeVsAir = "Y" Then
                            lbcVehAff.Selected(lbcVehAff.NewIndex) = True
                        End If
                    End If
                End If
            'End If
        Next iLoop
        
        lbcSelection(0).Visible = False
        CkcAll.Visible = False
        
        'lbcVehAff = VEhicles
        'lbcSelection(1) = stations
        'lbcSelection(0) = unused
        'lbcAvailNames = AVail Names
        lbcVehAff.Width = imListBoxWidth * 2 + 360
        lbcSelection(1).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240, (lbcVehAff.Width / 2) - 240, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + ckcAllStations.Height + 240) - 120
        ckcAllStations.Caption = "All Stations"
        ckcAllStations.Value = vbUnchecked
        lbcSelection(1).Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
                    If tgStationInfo(iLoop).sUsedForPledgeVsAir = "Y" And igRptIndex = PLEDGEVSAIR_RPT Then      '5-21-12  set the station selected if used in pledged vs air report
                        lbcSelection(1).Selected(lbcSelection(1).NewIndex) = True
                    End If

                End If
            End If
        Next iLoop
        lbcSelection(1).Visible = True
        ckcAllStations.Move lbcVehAff.Left, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        'TTP 9943
        cmdStationListFile.Move ckcAllStations.Left + ckcAllStations.Width, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        ckcAllStations.Visible = True
        lacAvailNames.Move lbcSelection(0).Left + imListBoxWidth + 360, ckcAllStations.Top
        lbcAvailNames.Move lbcSelection(0).Left + imListBoxWidth + 360, lbcSelection(1).Top, lbcSelection(1).Width, lbcSelection(1).Height
    Else
        chkListBox.Caption = "All Advertisers"
        chkListBox.Value = 0    '
        lbcVehAff.Clear
        For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                lbcVehAff.AddItem Trim$(tgAdvtInfo(iLoop).sAdvtName)
                lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgAdvtInfo(iLoop).iCode
        Next iLoop
        
        'lbcVehAff = Advertisers
        'lbcSelection(1) = stations
        'lbcSelection(0) = vehicles

        
        'populate stations & vehicles
        lbcVehAff.Width = imListBoxWidth
        lbcSelection(0).Move lbcVehAff.Left, lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240, imListBoxWidth, frcSelection.Height - (lbcVehAff.Top + lbcVehAff.Height + CkcAll.Height + 240) - 120
        CkcAll.Caption = "All Vehicles"
        CkcAll.Value = vbUnchecked
        lbcSelection(0).Clear
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
                lbcSelection(0).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgVehicleInfo(iLoop).iCode
                If igRptIndex = PLEDGEVSAIR_RPT Then       '5-21-12  set the station selected if used in pledged vs air report
                    iIndex = gBinarySearchVff(tgVehicleInfo(iLoop).iCode)
                    If iIndex > 0 Then
                        If tgVffInfo(iIndex).sPledgeVsAir = "Y" Then
                            lbcSelection(0).Selected(lbcSelection(0).NewIndex) = True
                        End If
                    End If
                End If
            'End If
        Next iLoop
        lbcSelection(0).Visible = True
        CkcAll.Move lbcSelection(0).Left, lbcSelection(0).Top - (CkcAll.Height + 120) '1800
        CkcAll.Visible = True
    
        lbcSelection(1).Move lbcAvailNames.Left, lbcSelection(0).Top, imListBoxWidth, frcSelection.Height - lbcSelection(0).Top - 120
        ckcAllStations.Caption = "All Stations"
        ckcAllStations.Value = vbUnchecked
        lbcSelection(1).Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                If tgStationInfo(iLoop).iType = 0 Then
                    lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                    lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
                     If tgStationInfo(iLoop).sUsedForPledgeVsAir = "Y" And igRptIndex = PLEDGEVSAIR_RPT Then       '5-21-12  set the station selected if used in pledged vs air report
                        lbcSelection(1).Selected(lbcSelection(1).NewIndex) = True
                    End If
                End If
            End If
        Next iLoop
        lbcSelection(1).Visible = True
        ckcAllStations.Move lbcSelection(1).Left, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        'TTP 9943
        cmdStationListFile.Move ckcAllStations.Left + ckcAllStations.Width, lbcSelection(1).Top - (ckcAllStations.Height + 120)
        ckcAllStations.Visible = True
        lbcAvailNames.Move lbcSelection(1).Left, lbcVehAff.Top      'move avail names box above the Statio list box
        lacAvailNames.Move lbcSelection(1).Left, chkListBox.Top
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

Private Sub optTimeSort_Click(Index As Integer)
    '4-17-08 regardless of sort, always leave ShowExactStationFeed as checked on and not shown
    '        this will exclude spots not carried
    'If optTimeSort(0).Value = True Then
    '    chkShowExact.Visible = True
    'Else
    '    chkShowExact.Visible = False
    '    chkShowExact.Value = vbUnchecked
    'End If
End Sub

Private Sub optUseDates_Click(Index As Integer)
    If Index = 0 Then
        Label3.Caption = "Fed Dates- Start"
    Else
        Label3.Caption = "Aired Dates- Start"
    End If
End Sub
