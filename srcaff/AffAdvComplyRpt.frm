VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdvComplyRpt 
   Caption         =   "Advertiser Compliance Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7575
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
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
      Height          =   4380
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   6960
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Station"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   41
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Network"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   40
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton optDatesTimes 
         Caption         =   "Pledged"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   39
         Top             =   2100
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optDatesTimes 
         Caption         =   "As Sold"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   38
         Top             =   2100
         Width           =   885
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CheckBox ckcShowDiscrepancyCodes 
         Caption         =   "Show Discrepancy Codes"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3480
         Width           =   2775
      End
      Begin VB.CheckBox ckcNewPage 
         Caption         =   "New Page Each Advertiser"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   2775
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Text            =   "12/8/2022"
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
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Text            =   "12/8/2022"
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
      Begin VB.Frame frcUseDates 
         BorderStyle     =   0  'None
         Caption         =   "Time Sort by"
         Height          =   225
         Left            =   960
         TabIndex        =   32
         Top             =   840
         Width           =   1875
         Begin VB.OptionButton optUseDates 
            Caption         =   "Aired"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.OptionButton optUseDates 
            Caption         =   "Fed "
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   645
         End
      End
      Begin VB.ListBox lbcStatus 
         Height          =   840
         ItemData        =   "AffAdvComplyRpt.frx":0000
         Left            =   120
         List            =   "AffAdvComplyRpt.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   1
         ItemData        =   "AffAdvComplyRpt.frx":0004
         Left            =   5280
         List            =   "AffAdvComplyRpt.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   0
         ItemData        =   "AffAdvComplyRpt.frx":0008
         Left            =   3480
         List            =   "AffAdvComplyRpt.frx":000A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
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
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   3165
         Begin VB.OptionButton optSortby 
            Caption         =   "Adv, Station, Vehicle, Date, Time"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   2835
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Adv, Vehicle, Station, Date, Time"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   2685
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1230
         ItemData        =   "AffAdvComplyRpt.frx":000C
         Left            =   3510
         List            =   "AffAdvComplyRpt.frx":000E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   3225
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3510
         TabIndex        =   22
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6480
         Picture         =   "AffAdvComplyRpt.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Select Stations from File.."
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lacNonComply 
         Caption         =   "Non-compliant Only"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lacDatesTimes 
         Caption         =   "Show Days/Dates/Times"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lacSortBy 
         Caption         =   "Sort by-"
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Use Dates"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"AffAdvComplyRpt.frx":057A
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
         TabIndex        =   26
         Top             =   3840
         Visible         =   0   'False
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   570
         Width           =   345
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Feed Dates- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1335
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
      FormDesignHeight=   6300
      FormDesignWidth =   7575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   31
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   30
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   29
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
Attribute VB_Name = "frmAdvComplyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmAdvComply -  Produce a report by advertiser to show whether
'*  stations are airing their spots in compliance with the demands of
'*  the advertiser, as shown on the ad sales order.
'*
'*  The generation of AST spots are retrieved from gGetAstInfo which
'*  returns all spots, whether aired, missed, not carried. This report
'*  will show all spot status except Not Carried.  The spot status are
'*  defaulted in a list box to include, except for Not Carried. This
'*  list box is hidden to user for selection.
'
'*  Spots are tested against the contract days, dates and times for
'*  compliance (vs the pledged days & times).
'*
'   Create a prepass file in AFR which only has a pointer to the AST file
'*
'*  This report has been copied from module frmPldgAir (Pledges & fed vs Aired rpt
'   2-25-13
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
    Unload frmAdvComplyRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    'Dan 7/20/11 this creates 3 variants and 1 integer
'    Dim i, j, X, Y, iPos As Integer
    Dim i As Integer, j As Integer, X As Integer, Y As Integer, iPos As Integer
    Dim sCode As String
    'Dan 7/20/11 bm and sName not used.  Can't dim this way, first 4 in line become variants
'    Dim bm As Variant
'    Dim sName, sVehicles, sStations, sAdvt, sStatus As String
    Dim sVehicles As String, sStations As String, sAdvt As String, sStatus As String
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
    'Dim slDiscrepOnly As String      '6-26-06 selectivity for discreps only
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
    Dim ilShowExact As Integer
    Dim ilIncludeNotCarried As Integer
    Dim ilIncludeNotReported As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim sUserTimes As String
    Dim blTestAsSold As Boolean
    'Dim blDiscrepOnly As Boolean        'true if discrep only, show only those spots not Compliant:  include the Not Reported
    Dim blNetworkDiscrep As Boolean     'true to filter out network discreps
    Dim blStationDiscrep As Boolean    'true to filter out station discreps
    Dim blInclDiscrepAndNonDiscrep As Boolean   'include both discrep & non-discrep spots
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim slSelectedAVailName As String
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    ReDim ilSelectedVehicles(0 To 0) As Integer
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global gen date & time variables so it doesnt have to be passed
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
    sgCrystlFormula5 = sUserTimes
    
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = Format(slInputStartDate, "m/d/yyyy")
    'default to Monday
    Do While Weekday(sStartDate, vbSunday) <> vbMonday
        sStartDate = DateAdd("d", -1, sStartDate)
    Loop

    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    sVehicles = ""
    sStations = ""
    sAdvt = ""
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
    
    'format the sql query for the selection of spot statuses
    'get the description of spot statuses (included/excluded) to show on report
    gGetSQLStatusForCrystal lbcStatus, sStatus, slSelection, ilIncludeNotCarried, ilIncludeNotReported
    sgCrystlFormula6 = slSelection
       
    If optSortby(0).Value = True Then       'advt, vehicle, station
        sgCrystlFormula1 = "'V'"
        iType = 1
    Else
        sgCrystlFormula1 = "'S'"        'advt,station,vehicle
        iType = 2
    End If
    
  
    gInitStatusSelections tmStatusOptions               '3-14-12 set all options to exclude
    gSetStatusOptions lbcStatus, ilIncludeNotReported, tmStatusOptions
    tmStatusOptions.iInclResolveMissed = True           'show the missed reference if including mg/replacements
    
    blNetworkDiscrep = False
    blStationDiscrep = False
    blInclDiscrepAndNonDiscrep = False
    If ckcDiscrep(0).Value = vbChecked Or ckcDiscrep(1).Value = vbChecked Then              'discreps only
        sgCrystlFormula4 = "'Y'"
        'If discrep only, gather only those whose afrCompliant flag is set
        If ckcDiscrep(0).Value = vbChecked Then         'network discrepancies only
            blNetworkDiscrep = True
        End If
        If ckcDiscrep(1).Value = vbChecked Then
            blStationDiscrep = True
        End If
        'blDiscrepOnly = True
    Else
        'blDiscrepOnly = False
        sgCrystlFormula4 = "'N'"                      'All
        blInclDiscrepAndNonDiscrep = True
    End If
    
    If ckcNewPage.Value = vbChecked Then            'New page each adv
        sgCrystlFormula6 = "'Y'"
    Else
        sgCrystlFormula6 = "'N'"                      'All
    End If

    If ckcShowDiscrepancyCodes.Value = vbChecked Then            'Show error code if application, not used for normal non-compliant spot.  ie.  LST/SDF/CLF read error
        sgCrystlFormula7 = "'Y'"
    Else
        sgCrystlFormula7 = "'N'"                      'All
    End If
    
    
    If (optDatesTimes(0).Value = True) Then         'show As Sold days days & times
        sgCrystlFormula8 = "'N'"
        blTestAsSold = True                          'show line compliancy
    Else
        sgCrystlFormula8 = "'S'"                    'show station non-compliant spots
        blTestAsSold = False                        'show pledged compliancy
    End If
        

    'sgCrystlFormula8 used for the SiteCompliantFlag to pass to crystal
   
    'dFWeek = CDate(sStartDate)
    dFWeek = CDate(slInputStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    'dFWeek = CDate(sEndDate)
    dFWeek = CDate(slInputEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    
     ilAdvtOption = False                            'all advt
     If chkListBox.Value = vbUnchecked Then
        ilAdvtOption = True
     End If
    '9-18-08 Used to have a question to show exact times which excluded Not Carried spots.
    'But we had to remove that and give user option to see those Not Carried spots.
    'In order to include/exclude any of the statuses, it will use test the selection of
    'statuses from the list box
    
    If ilIncludeNotCarried Then       'user selected Not Carried?
        ilShowExact = False            '9-18-08
    Else
        ilShowExact = True
    End If
    
    blUseAirDAte = False                'default to use Feed dates to retrieve spots

    gPopDaypart
   
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
                                
    ilIncludeNonRegionSpots = True          '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
    blFilterAvailNames = False
    blIncludePledgeInfo = True      'need pledged info
    
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
    tlSpotRptOptions.bNetworkDiscrep = blNetworkDiscrep
    tlSpotRptOptions.bStationDiscrep = blStationDiscrep
    tlSpotRptOptions.lContractNumber = 0            '6-4-18 no single contract option in this report
    
    'list box sent after blFilterAvailName is n/a for this report
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterBy, lbcVehAff, tmStatusOptions, blFilterAvailNames, lbcVehAff, blTestAsSold, blDiscrepOnly
    '2-19-14 change call for new ast design
    gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff, blTestAsSold      ', blDiscrepOnly
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcVehAff, lbcVehAff, blTestAsSold      ', blDiscrepOnly     '5-30-18

    
    SQLQuery = "Select afrastCode, afrPledgeDate, afrPledgeStartTime, afrPledgeEndTime, afrPledgeStatus,  "
    '12/11/13: Pledge information obtained from astInfo instead of ast
    SQLQuery = SQLQuery & "astAtfCode,  astAirDate, astAirTime, astStatus, astCPStatus, shttCallLetters, adfName, VefName, vpfAllowSplitCopy "

    
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrastcode = astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on astvefcode = vefcode "
    SQLQuery = SQLQuery & " INNER JOIN vpf_Vehicle_Options on vefcode = vpfvefkcode "

    SQLQuery = SQLQuery & "WHERE (astFeedDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " and (astAirTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "' AND astAirTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "')"
    SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfAdvComply.rpt", "AfAdvComply"
        
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
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "AdvComplyRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
        
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    gUserActivityLog "E", sgReportListName & ": Clear AFR"
    'ilRet = gClearAFR(frmAdvComplyRpt)
 
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AdvComplyRpt-cmdReport_Click"
    Exit Sub
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAdvComplyRpt
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
    End If
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts frmAdvComplyRpt
    lacStatusDesc.Move lbcStatus.Left, lbcStatus.Top + lbcStatus.Height + 60
    lacStatusDesc.FontSize = 8
       
    gCenterForm frmAdvComplyRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim lRg As Long
    Dim lRet As Long
    Dim ilRet As Integer
    Dim ilHideNotCarried As Integer

    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    igRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    imChkListBoxIgnore = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
'    8-4-14 this site has been removed
'    SQLQuery = "SELECT * From Site Where siteCode = 1"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
'        If rst!siteCompliantBy = "A" Then
'            sgCrystlFormula8 = "'A'"
'        End If
'    End If
'
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True
    gPopSpotStatusCodes lbcStatus, ilHideNotCarried           'populate list box with hard-coded spot status codes
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried         '3-14-12 populate list box with hard-coded extended spot status codes
    gPopVff
    
    'scan to see if any vef (vpf) are using avail names. Dont show the legend on input screen
    'if there are no vehicles using avail names
    lacStatusDesc.Visible = False
    ckcShowDiscrepancyCodes.Visible = False     'this is for debugging only, to show any error codes that may occur
'dont show on this report for now
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
    
    Set frmAdvComplyRpt = Nothing
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

Private Sub optDatesTimes_Click(Index As Integer)
   If Index = 0 Then           'show as sold; turn off ability to get noncopliant for station non compliance
        ckcDiscrep(1).Value = vbUnchecked
        ckcDiscrep(1).Enabled = False
        ckcDiscrep(0).Enabled = True
    Else
        ckcDiscrep(0).Value = vbUnchecked       'show pledged dates & times
        ckcDiscrep(0).Enabled = False
        ckcDiscrep(1).Enabled = True
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
    

    chkListBox.Caption = "All Advertisers"
    chkListBox.Value = vbUnchecked    '
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
    
    lbcSelection(1).Move lbcVehAff.Left + lbcVehAff.Width - lbcSelection(0).Width, lbcSelection(0).Top, lbcSelection(0).Width, frcSelection.Height - lbcSelection(0).Top - 120
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
    ckcAllStations.Visible = True
    cmdStationListFile.Move lbcSelection(1).Left + ckcAllStations.Width, lbcSelection(1).Top - (ckcAllStations.Height + 120)
  
    Screen.MousePointer = vbDefault
End Sub

