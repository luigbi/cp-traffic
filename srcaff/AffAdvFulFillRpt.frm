VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdvFulFillRpt 
   Caption         =   "Advertiser Fulfillment Report"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9810
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
      Height          =   5460
      Left            =   240
      TabIndex        =   36
      Top             =   1680
      Width           =   9240
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8640
         Picture         =   "AffAdvFulFillRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Select Stations from File.."
         Top             =   240
         Width           =   360
      End
      Begin VB.CheckBox ckcAllContracts 
         Caption         =   "All Contracts"
         Height          =   255
         Left            =   7080
         TabIndex        =   26
         Top             =   2730
         Width           =   1215
      End
      Begin VB.ListBox lbcContracts 
         Height          =   2010
         ItemData        =   "AffAdvFulFillRpt.frx":056A
         Left            =   7080
         List            =   "AffAdvFulFillRpt.frx":056C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   3000
         Width           =   2000
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Text            =   "6/13/24"
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
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Network"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   915
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Station"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Text            =   "6/13/24"
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
      Begin VB.ListBox lbcAvailNames 
         Height          =   1425
         ItemData        =   "AffAdvFulFillRpt.frx":056E
         Left            =   4920
         List            =   "AffAdvFulFillRpt.frx":0570
         TabIndex        =   48
         Top             =   3360
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ComboBox cbcAVailNames 
         Height          =   315
         ItemData        =   "AffAdvFulFillRpt.frx":0572
         Left            =   2400
         List            =   "AffAdvFulFillRpt.frx":0574
         TabIndex        =   9
         Top             =   1515
         Width           =   1020
      End
      Begin VB.CheckBox ckcMarkRegional 
         Caption         =   "Highlight Regional Copy"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   1800
         Width           =   1920
      End
      Begin VB.CheckBox ckcExclSpotsLackReg 
         Caption         =   "Exclude Spots Lacking Regional Copy"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2970
      End
      Begin VB.CheckBox ckcExclTotals 
         Caption         =   "Exclude Station Spot Counts"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   2100
      End
      Begin VB.ComboBox cbcExtra 
         Height          =   315
         ItemData        =   "AffAdvFulFillRpt.frx":0576
         Left            =   600
         List            =   "AffAdvFulFillRpt.frx":0578
         TabIndex        =   8
         Top             =   1515
         Width           =   1020
      End
      Begin VB.ComboBox cbcSort 
         Height          =   315
         ItemData        =   "AffAdvFulFillRpt.frx":057A
         Left            =   1080
         List            =   "AffAdvFulFillRpt.frx":057C
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1125
         Width           =   2340
      End
      Begin VB.CheckBox ckcAllCategories 
         Caption         =   "All Categories"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   2730
         Width           =   1380
      End
      Begin VB.ListBox lbcSelection 
         Height          =   2010
         Index           =   2
         ItemData        =   "AffAdvFulFillRpt.frx":057E
         Left            =   2760
         List            =   "AffAdvFulFillRpt.frx":0580
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   3000
         Width           =   2000
      End
      Begin VB.Frame frcUseDates 
         BorderStyle     =   0  'None
         Caption         =   "Time Sort by"
         Height          =   225
         Left            =   1380
         TabIndex        =   42
         Top             =   120
         Width           =   1740
         Begin VB.OptionButton optUseDates 
            Caption         =   "Aired"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   2
            Top             =   15
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton optUseDates 
            Caption         =   "Fed "
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   1
            Top             =   15
            Width           =   705
         End
      End
      Begin VB.CheckBox ckcShowStatus 
         Caption         =   "Show Status Codes"
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   3300
         Width           =   1785
      End
      Begin VB.ListBox lbcStatus 
         Height          =   1230
         ItemData        =   "AffAdvFulFillRpt.frx":0582
         Left            =   120
         List            =   "AffAdvFulFillRpt.frx":0584
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   3600
         Width           =   2550
      End
      Begin VB.ListBox lbcSelection 
         Height          =   2010
         Index           =   1
         ItemData        =   "AffAdvFulFillRpt.frx":0586
         Left            =   7080
         List            =   "AffAdvFulFillRpt.frx":0588
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   2000
      End
      Begin VB.ListBox lbcSelection 
         Height          =   2010
         Index           =   0
         ItemData        =   "AffAdvFulFillRpt.frx":058A
         Left            =   4920
         List            =   "AffAdvFulFillRpt.frx":058C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   2000
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   7080
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CkcAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtEndTime 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "12M"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtStartTime 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "12M"
         Top             =   780
         Width           =   975
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   2010
         ItemData        =   "AffAdvFulFillRpt.frx":058E
         Left            =   4920
         List            =   "AffAdvFulFillRpt.frx":0590
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   3000
         Width           =   2000
      End
      Begin VB.CheckBox ckcAllAdvt 
         Caption         =   "All Advertisers"
         Height          =   255
         Left            =   4920
         TabIndex        =   24
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label lacDiscrep 
         Caption         =   "Non-Compliant Only"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2535
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Avails"
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label lacStatus 
         Caption         =   "Spot Status Codes"
         Height          =   225
         Left            =   120
         TabIndex        =   46
         Top             =   3075
         Width           =   1455
      End
      Begin VB.Label lacExtraInfo 
         Caption         =   "Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1545
         Width           =   555
      End
      Begin VB.Label lacSort 
         Caption         =   "Sort Adv by"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Use Dates/Times"
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"AffAdvFulFillRpt.frx":0592
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   4800
         Width           =   2475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   810
         Width           =   465
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Dates- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   450
         Width           =   525
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
      FormDesignHeight=   7200
      FormDesignWidth =   9810
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   30
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   29
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   28
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
         TabIndex        =   34
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   35
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Top             =   810
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   32
         Top             =   525
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmAdvFulFillRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmAdvFulFillRpt (Advertiser Fulfillment)- Create a report of spots (
'   from AST) that will be sorted by advertiser
'   (with many different subsorts: vehicle, station, market,
'   zip code, format, etc)
'*  Spots are generated from AST (whether reported or not, by option).
'   Selectivity always includes all vehicles, stations, and advertisers (altho
'   it should primarily be requested for an advertiser.
'   Create a prepass file in AFR which only has a pointer to the AST file
'   All spots for the vehicle are created and the filtering of spots is
'   processed thru the sql call to Crystal.
'
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imckcAllAdvtIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private imckcAllCategoriesIgnore As Integer
Private imckcAllContractsIgnore As Integer
Private hmAst As Integer
Private smExtraField As String
Private bmFirstTime As Boolean      '2-20-20
Private imPrevSortBy As Integer
Private tmStatusOptions As STATUSOPTIONS
Private Sub mGetAllContracts()
    'ByVal smFWKDate As String, ByVal smLWkDate As String
    
    Dim ilLoop As Integer
    Dim imAdfCode As Integer
    
    If Trim(CalOnAirDate.Text) = "" Or Trim(CalOffAirDate.Text) = "" Then Exit Sub
    
    lbcContracts.Clear
    ckcAllContracts.Value = 0
    
    For ilLoop = 0 To lbcVehAff.ListCount - 1
        If lbcVehAff.Selected(ilLoop) Then
            imAdfCode = lbcVehAff.ItemData(ilLoop)
            
            'eliminated join with SDF, AdfCode is available in CHF  Date: 8/2/2019  FYM
            'TTP 10911 - Advertiser Fulfillment report: contract selectivity list is not including manually scheduled contracts
            'SQLQuery = "SELECT DISTINCT chfCntrNo from chf_contract_header WHERE chfschstatus = 'F' and chfdelete = 'N' "
            SQLQuery = "SELECT DISTINCT chfCntrNo from chf_contract_header WHERE chfschstatus in ('F','M') and chfdelete = 'N' "
            SQLQuery = SQLQuery + " AND ((chfenddate >= '" & Format$(CalOnAirDate.Text, sgSQLDateForm) & "' AND chfstartdate <= '" & Format$(CalOffAirDate.Text, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " AND chfAdfCode = " & imAdfCode & ")"
            SQLQuery = SQLQuery + " ORDER BY chfCntrNo"
            
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                ckcAllContracts.Visible = True
                lbcContracts.Visible = True
                ckcAllAdvt.Value = vbUnchecked
            End If
            While Not rst.EOF
                lbcContracts.AddItem rst!chfCntrNo  ', " & rst(1).Value & ""
                rst.MoveNext
            Wend
        End If
    Next ilLoop

End Sub

'Private igRptIndex As Integer


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
Private Sub cbcExtra_Change()
Dim ilIndex As Integer
    ilIndex = cbcExtra.ListIndex
    smExtraField = cbcExtra.List(ilIndex)
End Sub
Private Sub cbcExtra_Click()
    cbcExtra_Change
End Sub

Private Sub cbcSort_Click()
Dim ilSetIndex As Integer
Dim ilLoop As Integer
Dim llLoop As Long      'was using ilLoop (integer)
Dim ilZero As Integer
Dim slStr As String
    ilSetIndex = cbcSort.ListIndex
    lbcSelection(2).Clear       'clear the categories list box
    ilZero = False          'for rank, only show number 0 once
    lbcSelection(2).Visible = True
    ckcAllCategories.Visible = True
    ckcAllCategories.Value = vbUnchecked

    If ilSetIndex = SORT_ISCI_STN Or ilSetIndex = SORT_ISCI_VEH Then
        If imPrevSortBy <> ilSetIndex Then               'if its isci sort and previous was isci sort, may not have to load box
            'get the isci that match any selected advertisers
            mGetAllISCI
        End If
    Else
        ckcAllContracts.Caption = "All Contracts"
        If imPrevSortBy <> ilSetIndex Then
            mGetAllContracts
        End If
    End If
    
    Select Case ilSetIndex
'        Case 0:         'DMA Market
        Case SORT_DMA_NAME:         'DMA Market
            ckcAllCategories.Caption = "All DMA Markets"
            For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgMarketInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMarketInfo(llLoop).lCode
            Next llLoop
'        Case 1:         'DMA Rank
        Case SORT_DMA_RANK:         'DMA Rank
            ckcAllCategories.Caption = "All DMA Mkt Ranks"
            For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                If tgMarketInfo(llLoop).iRank = 0 Then         'only show number 0 once
                     If Not ilZero Then
                         slStr = Trim$(Str(tgMarketInfo(llLoop).iRank))
                         Do While Len(slStr) < 5
                            slStr = "0" & slStr
                         Loop
                         lbcSelection(2).AddItem slStr
                         lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMarketInfo(llLoop).lCode
                         ilZero = True
                     End If
                Else
                    slStr = Trim$(Str(tgMarketInfo(llLoop).iRank))
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    lbcSelection(2).AddItem slStr
                    lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMarketInfo(llLoop).lCode
                End If
            Next llLoop
'        Case 2:         'Format
        Case SORT_FORMAT:         'Format
            ckcAllCategories.Caption = "All Formats"
           For llLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgFormatInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgFormatInfo(llLoop).lCode
            Next llLoop
'        Case 3:         'MSAA Market
        Case SORT_MSA_NAME:         'MSAA Market
            ckcAllCategories.Caption = "All MSA Markets"
            For llLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgMSAMarketInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMSAMarketInfo(llLoop).lCode
            Next llLoop
'        Case 4:         'MSA Rank
        Case SORT_MSA_RANK:         'MSA Rank
            ckcAllCategories.Caption = "All MSA Mkt Ranks"
            For llLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
                If tgMSAMarketInfo(llLoop).iRank = 0 Then         'only show number 0 once
                     If Not ilZero Then
                         slStr = Trim$(Str(tgMSAMarketInfo(llLoop).iRank))
                         Do While Len(slStr) < 5
                            slStr = "0" & slStr
                         Loop
                         lbcSelection(2).AddItem slStr
                         lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMSAMarketInfo(llLoop).lCode
                         ilZero = True
                     End If
                Else
                    slStr = Trim$(Str(tgMSAMarketInfo(llLoop).iRank))
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    lbcSelection(2).AddItem slStr
                    lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMSAMarketInfo(llLoop).lCode
                End If

            Next llLoop
'        Case 5:         'State
        Case SORT_STATE:         'State
            ckcAllCategories.Caption = "All States"
            For ilLoop = 0 To UBound(tgStateInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgStateInfo(ilLoop).sPostalName) & " (" & Trim$(tgStateInfo(ilLoop).sName) & ")"
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgStateInfo(ilLoop).iCode
            Next ilLoop
'        Case 6:         'stations, list box already populated
        Case SORT_STN:         'stations, list box already populated
            lbcSelection(2).Visible = False
            ckcAllCategories.Visible = False
            ckcAllCategories.Value = vbChecked
'        Case 7:     'time zone
        Case SORT_TZ:     'time zone
            ckcAllCategories.Caption = "All Time Zones"
            For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgTimeZoneInfo(ilLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgTimeZoneInfo(ilLoop).iCode
            Next ilLoop
'        Case 8, 9:    'vehicles/dma name, vehicles/dma rank, already shown. disable the categories listbox
        Case SORT_VEH_DMA, SORT_VEH_RANK:    'vehicles/dma name, vehicles/dma rank, already shown. disable the categories listbox
            lbcSelection(2).Visible = False
            ckcAllCategories.Visible = False
            ckcAllCategories.Value = vbChecked
        Case SORT_ISCI_STN, SORT_ISCI_VEH:     'vehicles/dma name, vehicles/dma rank, already shown. disable the categories listbox
            lbcSelection(2).Visible = False
            ckcAllCategories.Visible = False
            ckcAllCategories.Value = vbChecked
            ckcAllContracts.Caption = "All ISCI"
    End Select
    imPrevSortBy = ilSetIndex
End Sub

Private Sub ckcAllAdvt_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imckcAllAdvtIgnore Then
        Exit Sub
    End If
    If ckcAllAdvt.Value = vbChecked Then
        iValue = True
        'Date: 8/6/2019 hide ALL Contracts selection FYM
        ckcAllContracts.Visible = False
        lbcContracts.Clear
        lbcContracts.Visible = False
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imckcAllAdvtIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        ckcAllAdvt.Value = vbChecked
        imckcAllAdvtIgnore = False
    End If
    
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub





Private Sub ckcAllCategories_Click()
Dim iValue As Integer
Dim i As Integer
Dim lErr As Long
Dim lRet As Long
Dim lRg As Long


    If imckcAllCategoriesIgnore Then
        Exit Sub
    End If
    If ckcAllCategories.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSelection(2).ListCount > 0 Then
        imckcAllCategoriesIgnore = True
        lRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(2).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllCategoriesIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllContracts_Click()
Dim iValue As Integer
Dim lErr As Long
Dim lRg As Long
Dim lRet As Long
    
    If imckcAllContractsIgnore Then
        Exit Sub
    End If
    If ckcAllContracts.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcContracts.ListCount > 0 Then
        imckcAllContractsIgnore = True
        lRg = CLng(lbcContracts.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcContracts.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllContractsIgnore = False
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

Private Sub ckcAllVehicles_Click()
Dim iValue As Integer
Dim lErr As Long
Dim lRg As Long
Dim lRet As Long
    If imckcAllIgnore Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = vbChecked Then
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

Private Sub ckcDiscrep_Click(Index As Integer)
   If ckcDiscrep(0).Value = vbChecked Or ckcDiscrep(1).Value = vbChecked Then        'discrepancy option
        lbcStatus.Enabled = False
    Else
        lbcStatus.Enabled = True
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmAdvFulFillRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    Dim ilTemp As Integer
    Dim sCode As String         'contract #
    Dim sName, sVehicles, sStations, sAdvt, sStatus As String
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
    Dim ilSelected As Integer
    Dim ilNotSelected As Integer
    Dim slStatusSelected As String
    Dim slStatusNotSelected As String
    Dim slSelection As String
    'Dim sGenDate As String
    'Dim sGenTime As String
    Dim slInputStartDate As String
    Dim slInputEndDate As String
    Dim ilAdvt As Integer               'set to -1 to retrieve all advt
    Dim ilShowExact As Integer
    Dim ilIncludeNotCarried As Integer
    Dim ilIncludeNotReported As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim sUserTimes As String                'user times selected (text for crystal report heading)
    Dim ilSortBy As Integer                 'sort index selected
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    Dim blNetworkDiscrep As Boolean
    Dim blStationDiscrep As Boolean
    ReDim ilSelectedVehicles(0 To 0) As Integer       '5-30-18
    
    Dim iLoop As Integer            'loop counter for contracts Date:8/2/2019   FYM
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)

    
    cbcExtra_Change
    ilSortBy = cbcSort.ListIndex            'user selected sort option
    
    'adjust the #s sent to crystal so it doesnt all have to be changed for the different versions and new numbering since adding 2 new sort options by isci
    ilTemp = ilSortBy
    If ilTemp = 3 Then
        ilTemp = 10
    ElseIf ilTemp = 4 Then
        ilTemp = 11
    ElseIf ilTemp >= 5 Then
        ilTemp = ilTemp - 2
    End If
'    sgCrystlFormula1 = Str$(ilSortBy)       'send to crystal
    sgCrystlFormula1 = Str$(ilTemp)       'send to crystal
    
    
    sgCrystlFormula8 = Trim$(smExtraField)  'extra field to show requested
    ilIncludeNotCarried = False                 'never show spots not carried
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
    
'    sCode = Trim$(edcContract.Text)
'    If Not IsNumeric(sCode) And (sCode <> "") Then
'        gMsgBox "Enter valid contract number", vbOKOnly
'        edcContract.SetFocus
'        Exit Sub
'    End If
    
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
    sgCrystlFormula5 = sUserTimes
    
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmAdvFulFillRpt, "PledgeAired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    blUseAirDAte = True                             'assume to use air date for spot filter
    
    sStartDate = Format(slInputStartDate, "m/d/yyyy")
    'default to Monday
    Do While Weekday(sStartDate, vbSunday) <> vbMonday
        sStartDate = DateAdd("d", -1, sStartDate)
    Loop

    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    
    If optUseDates(1).Value = True Then     'use air dates vs fed dates, need to backup the week and process extra week
                                            'spot may air outside the week
         'sStartDate = DateAdd("d", -7, sStartDate)     '2-10-14 no need to backup the aired date with new ast design
         sgCrystlFormula4 = "'A'"
         blUseAirDAte = True                             'assume to use air date for spot filter
   Else
        sgCrystlFormula4 = "'F'"
        blUseAirDAte = False
    End If
    
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
    slStatusSelected = ""
    slStatusNotSelected = ""
    
    ilAdvtOption = False            'no checking of advt
    If ckcAllAdvt.Value = vbUnchecked Then      '1 or more advt selected
        ilAdvtOption = True
    End If
          
    'determine option to include non-reported stations
    If Not ckcInclNotRecd.Value = vbChecked Then    'exclude non-reported (or cp not received) stations
        ilIncludeNotReported = False
        sCPStatus = " and (astCPStatus = 1 Or astCPStatus = 2)"
    Else
        ilIncludeNotReported = True
    End If
      
    'format the sql query for the selection of spot statuses
    'get the description of spot statuses (included/excluded) to show on report
    gGetSQLStatusForCrystal lbcStatus, sStatus, slSelection, ilIncludeNotCarried, ilIncludeNotReported
    sgCrystlFormula6 = slSelection
     
  
    gInitStatusSelections tmStatusOptions               '3-14-12 set all options to exclude
    gSetStatusOptions lbcStatus, ilIncludeNotReported, tmStatusOptions
    tmStatusOptions.iInclResolveMissed = False              'dont show the missed part of a mg or replacement spot since only status codes are shown
  
    If ckcShowStatus.Value = vbChecked Then            'show the status codes on the report
        sgCrystlFormula7 = "'Y'"
    Else
        sgCrystlFormula7 = "'N'"
    End If
    
    If ckcExclTotals.Value = vbChecked Then            'exclude station spot counts
        sgCrystlFormula9 = "'E'"
    Else
        sgCrystlFormula9 = "'I'"
    End If
    
    If ckcExclSpotsLackReg.Value = vbChecked Then   'exclude spots lacking regional copy?
        ilIncludeNonRegionSpots = False     'exclude spots without regional copy
    Else
        ilIncludeNonRegionSpots = True      'incl spots with/without regional copy
    End If
    
    '8-19-10 highlight regional spots
    If ckcMarkRegional.Value = vbChecked Then   'highlight regional spots with color?
        sgCrystlFormula10 = "'Y'       "
    Else
        sgCrystlFormula10 = "'N'      "
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
    '4-17-08 Pledge vs Air and Fed vs Air will always exclude spots not carried.
    '        Question on screen  Show Exact Station Feed has been hidden and defaulted to Yes
    
      '9-18-08 ckcShowExact value has been defaulted unchecked (to allow for the list of statuses to
    'be used instead for selection of Not Carried )
    If Not ilIncludeNotCarried Then
       ilShowExact = True
    Else
        ilShowExact = False
    End If
    
    'determine which category to filter if applicable
    ilFilterCatBy = -1           'nothing to test to filter
    If Not ckcAllCategories.Value = vbChecked Then      'some selective categories picked
    '2-20-20 2 new sort options by ISCI added
'       If ilSortBy = 0 Then            'metro name
        If ilSortBy = SORT_DMA_NAME Then            'metro name
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 1 Then        'metro rank
        ElseIf ilSortBy = SORT_DMA_RANK Then        'metro rank
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 2 Then        'format
        ElseIf ilSortBy = SORT_FORMAT Then        'format
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 3 Then        'msa name
        ElseIf ilSortBy = SORT_MSA_NAME Then        'msa name
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 4 Then        'msa rank
        ElseIf ilSortBy = SORT_MSA_RANK Then        'msa rank
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 5 Then        'state
        ElseIf ilSortBy = SORT_STATE Then        'state
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 6 Then        'station, do nothing (it has a different list box selection)
        ElseIf ilSortBy = SORT_STN Then        'station, do nothing (it has a different list box selection)
'        ElseIf ilSortBy = 7 Then        'time zone
        ElseIf ilSortBy = SORT_TZ Then        'time zone
            ilFilterCatBy = ilSortBy
'        ElseIf ilSortBy = 8 Or ilSortBy = 9 Then     'vehicle, do nothing (it has a different list box selection)
        ElseIf ilSortBy = SORT_VEH_DMA Or ilSortBy = SORT_VEH_RANK Then     'vehicle, do nothing (it has a different list box selection)
        ElseIf ilSortBy = SORT_ISCI_STN Or ilSortBy = SORT_ISCI_VEH Then        'isci, station,vehicle or isci, vehicle, station
        End If
    End If
    
    blNetworkDiscrep = False
    blStationDiscrep = False
    sgCrystlFormula12 = "' '"                  'no compliance discrepancies selected
    If ckcDiscrep(0).Value = vbChecked Or ckcDiscrep(1).Value = vbChecked Then             'discreps only
        '8-4-14 check only flags in common routine to see if non-compliant
        If ckcDiscrep(0).Value = vbChecked Then
            blNetworkDiscrep = True
            sgCrystlFormula12 = "'N'"
        End If
        If ckcDiscrep(1).Value = vbChecked Then
            blStationDiscrep = True
            If ckcDiscrep(0).Value = vbChecked Then     'both types of non-compliance checked
                sgCrystlFormula12 = "'B'"
            Else
                sgCrystlFormula12 = "'S'"
            End If
        End If
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    lbcAvailNames.ListIndex = cbcAVailNames.ListIndex
    
    blFilterAvailNames = False
    sgCrystlFormula11 = "'All Avail Names'"
    If lbcAvailNames.ListIndex > 0 Then           '1st element is All Avails, if not selected then need to test for the matching avail name in spot
        blFilterAvailNames = True
        For ilTemp = 0 To lbcAvailNames.ListCount - 1 Step 1
            If lbcAvailNames.Selected(ilTemp) Then
                sgCrystlFormula11 = "'" & Trim$(lbcAvailNames.List(ilTemp)) & " Avail Name'"
                Exit For
            End If
        Next ilTemp
    End If

    If tmStatusOptions.iNotReported = True Then        '3-26-15 if including not reported, need to see pledge data for those agreements not posted.  The pledge status is tested.
        blIncludePledgeInfo = True
    End If

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
    tlSpotRptOptions.lContractNumber = 0        'Val(sCode)       2-20-20 single selection replaced by multi-selection
    If ilSortBy = SORT_ISCI_STN Or ilSortBy = SORT_ISCI_VEH Then        'isci, station,vehicle or isci, vehicle, station
        tlSpotRptOptions.lContractNumber = -1                       'indicate isci filtering and testing
    End If
    'single contract selection from text box no longer applicable since multi-contract selection has been implemented
    'use this code by flagging as -1 to indicate selective ISCI codes rather than contract codes for filtering and testing
    
    
    'gBuildAstStnClr hmAst, sStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2)
    '3-26-12 new subrtn to select on all statuses, get mg/replacement references, and only create afr records that are to be reported
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2), tmStatusOptions
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2), tmStatusOptions, blFilterAvailNames, lbcAvailNames
    '2-10-14 use a structure to pass many rpt option variables
    gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcVehAff, lbcSelection(2), lbcAvailNames
    'Date:8/3/2019 added new optional parameter (lbcContracts)  FYM
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcVehAff, lbcSelection(2), lbcAvailNames, , lbcContracts       '5-30-18
    
    SQLQuery = "Select afrastCode, afrISCI,afrCreative, "
    SQLQuery = SQLQuery & "astAtfCode,  astAirDate, astAirTime, astStatus, astFeedDate, astFeedTime, astCntrno, astLen, "
    SQLQuery = SQLQuery & "shttCallLetters, shttState, shttStationID, adfName, VefName, "
    SQLQuery = SQLQuery & "mktName, mktRank, metName, metRank, "
    SQLQuery = SQLQuery & "fmtName, tztName, cpfCreative,"
    SQLQuery = SQLQuery & "attXDReceiverID "

    
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrastcode = astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on astvefcode = vefcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN att on astatfcode = attcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mkt on shttMktCode = mktCode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN met on shttMetCode = metCode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN fmt_Station_Format on shttFmtCode = FmtCode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN tzt on shttTztCode = tztCode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on astcpfcode = cpfCode "
    
    If optUseDates(1).Value = True Then             'use air dates vs fed dates
        SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    Else
        SQLQuery = SQLQuery & "WHERE (astFeedDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    End If
    
    SQLQuery = SQLQuery & " and (astAirTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "' AND astAirTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "')"
    SQLQuery = SQLQuery & " and (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfAdvFulFill.rpt", "AfAdvFulFill"
    
'    cmdReport.Enabled = True            'give user back control to gen, done buttons
'    cmdDone.Enabled = True
'    cmdReturn.Enabled = True
     
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
        gHandleError "AffErrorLog.txt", "AdvFulFillRpt-cmdReport_Click"
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
    gHandleError "AffErrorLog.txt", "AdvFulFillRpt-cmdReport_Click"
     Exit Sub
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAdvFulFillRpt
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
End Sub
Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts frmAdvFulFillRpt
    lacStatusDesc.Move lbcStatus.Left, lbcStatus.Top + lbcStatus.Height + 60
    lacStatusDesc.FontSize = 8
    frmAdvFulFillRpt.Caption = "Advertiser Fulfillment Report - " & sgClientName

    gCenterForm frmAdvFulFillRpt
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

    imckcAllAdvtIgnore = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    End If
    

    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried           'populate list box with hard-coded spot status codes

    'scan to see if any vef (vpf) are using avail names. Dont show the legend on input screen
    'if there are no vehicles using avail names
    lacStatusDesc.Visible = False
    For iLoop = LBound(tgVpfOptions) To UBound(tgVpfOptions) - 1
        If tgVpfOptions(iLoop).sAvailNameOnWeb = "Y" Then
            lacStatusDesc.Visible = True
            Exit For
        End If
    Next iLoop
        
    lbcSelection(0).Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcSelection(0).AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    
    lbcSelection(1).Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
        
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgAdvtInfo(iLoop).sAdvtName)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgAdvtInfo(iLoop).iCode
    Next iLoop
    
    blIncludeAllAvailNames = True
    gPopAndSortAvailNames blIncludeAllAvailNames, lbcAvailNames     'load avail names into hidden list box for common routines to process selected ones
    'transfer avail names into combo box, no space on screen
    For iLoop = 0 To lbcAvailNames.ListCount - 1 Step 1
        cbcAVailNames.AddItem lbcAvailNames.List(iLoop)
    Next iLoop
    cbcAVailNames.ListIndex = 0

    cbcSort.AddItem "DMA Market Name,Station,Vehicle,Date,Time"
    cbcSort.AddItem "DMA Market Rank/Mkt,Station,Vehicle,Date,Time"
    cbcSort.AddItem "Format,Station,Vehicle,Date,Time"
    cbcSort.AddItem "ISCI,Station,Vehicle,Date,Time"                'added 2-20-20
    cbcSort.AddItem "ISCI,Vehicle,Station,Date,Time"                'added 2-20-20
    cbcSort.AddItem "MSA Market Name,Station,Vehicle,Date,Time"
    cbcSort.AddItem "MSA Market Rank/Mkt,Station,Vehicle,Date,Time"
    cbcSort.AddItem "State,Station,Vehicle,Date,Time"
    cbcSort.AddItem "Station,Vehicle,Date,Time"
    cbcSort.AddItem "Time Zone,Station,Vehicle,Date,Time"
    'cbcSort.AddItem "Vehicle,Station,Date,Time"
    '8-11-10, replace VEhicle, station, Date, Time with 2 other vehicle sorts
    cbcSort.AddItem "Vehicle,DMA Name,Station,Date,Time"
    cbcSort.AddItem "Vehicle,DMA Rank,Station,Date,Time"
    cbcSort.ListIndex = 0
    bmFirstTime = True                                              '2-20-20
    imPrevSortBy = 0
    cbcExtra.AddItem "None"
    cbcExtra.AddItem "Creative Title"
    'cbcExtra.AddItem "DMA Market Name"
    'cbcExtra.AddItem "DMA Market Rank"
    cbcExtra.AddItem "Format"
    cbcExtra.AddItem "MSA Mkt Rank/Name"
    'cbcExtra.AddItem "MSA Market Rank"
    cbcExtra.AddItem "State"
    cbcExtra.AddItem "Time Zone"
    cbcExtra.AddItem "XDS Station ID"
    cbcExtra.ListIndex = 0

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmAdvFulFillRpt = Nothing
End Sub


Private Sub lbcContracts_Click()
       If imckcAllContractsIgnore Then
           Exit Sub
       End If
       If ckcAllContracts.Value = vbChecked Then
           imckcAllContractsIgnore = True
           ckcAllContracts.Value = vbUnchecked
           imckcAllContractsIgnore = False
       End If
End Sub

Private Sub lbcSelection_Click(Index As Integer)
 
    If Index = 0 Then                          'more vehicle or station selection
       If imckcAllIgnore Then
           Exit Sub
       End If
       If ckcAllVehicles.Value = vbChecked Then
           imckcAllIgnore = True
           ckcAllVehicles.Value = vbUnchecked
           imckcAllIgnore = False
       End If
    ElseIf Index = 1 Then                                    'station selection
       If imckcAllStationsIgnore Then
           Exit Sub
       End If
       If ckcAllStations.Value = vbChecked Then
           imckcAllStationsIgnore = True
           ckcAllStations.Value = vbUnchecked
           imckcAllStationsIgnore = False
       End If
    Else                                        'categories
        If imckcAllCategoriesIgnore Then
           Exit Sub
       End If
       If ckcAllCategories.Value = vbChecked Then
           imckcAllCategoriesIgnore = True
           ckcAllCategories.Value = vbUnchecked
           imckcAllCategoriesIgnore = False
       End If
    End If
 
End Sub
Private Sub lbcVehAff_Click()
    Dim sCode As String
    Dim ilLoop As Integer
    Dim slInputStartDate As String
    Dim slInputEndDate  As String
    Dim sStartTime As String
    Dim sEndTime As String
    Dim smFWkDate As String
    Dim smLWkDate As String
    Dim imAdfCode As Integer
    
    On Error GoTo ErrHand
    
    If imckcAllAdvtIgnore Then
        Exit Sub
    End If
    If ckcAllAdvt.Value = vbChecked Then
        imckcAllAdvtIgnore = True
        'ckcAllAdvt.Value = False
        ckcAllAdvt.Value = vbUnchecked    'chged from false to 0 10-22-99
        imckcAllAdvtIgnore = False
    End If
    
    lbcContracts.Clear
    ckcAllContracts.Value = 0        'chged from False to 0 10-22-99

    If lbcVehAff.ListIndex < 0 Then
        Exit Sub
    End If
    
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
    
'    If CalWeek.Text = "" Then
'        gMsgBox "Date must be specified.", vbOKOnly
'        CalWeek.SetFocus
'        Exit Sub
'    End If
    
'    If gIsDate(CalWeek.Text) = False Then
'        Beep
'        gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
'        CalWeek.SetFocus
'    Else
        smFWkDate = Format(slInputStartDate, sgShowDateForm)
''    End If
    
'    If CalWeek.Text = "" Then
'        gMsgBox "Date must be specified.", vbOKOnly
'        CalEndDate.SetFocus
'        Exit Sub
'    End If
    
'''    If gIsDate(CalEndDate.Text) = False Then
'''        Beep
'''        gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
'''        CalEndDate.SetFocus
'''    Else
        smLWkDate = Format(slInputEndDate, sgShowDateForm)
''    End If
    
    Screen.MousePointer = vbHourglass
    'smFWkDate & smLWkDAte = earliest/latest requested dates
    
    '2-20-20
    If cbcSort.ListIndex = SORT_ISCI_STN Or cbcSort.ListIndex = SORT_ISCI_VEH Then
        mGetAllISCI         '2-20-20
    Else
        mGetAllContracts 'Date: 8/6/2018 get all contracts based on selected advertisers FYM
    End If
    
''    For ilLoop = 0 To lbcVehAff.ListCount - 1
''        If lbcVehAff.Selected(ilLoop) Then
''            imAdfCode = lbcVehAff.ItemData(ilLoop)
''
''            '7-24-19 change access of unique contract #s from lst to sdf for speedup
'''            SQLQuery = "SELECT DISTINCT chfCntrNo from sdf_spot_detail inner join chf_contract_header on sdfchfcode = chfcode "
'''            SQLQuery = SQLQuery + " WHERE ((sdfDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND sdfDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
'''            SQLQuery = SQLQuery + " AND sdfAdfCode = " & imAdfCode & ")"
'''            SQLQuery = SQLQuery + " ORDER BY chfCntrNo"
''
''            'eliminated join with SDF, AdfCode is available in CHF  Date: 8/2/2019  FYM
''            SQLQuery = "SELECT DISTINCT chfCntrNo from chf_contract_header WHERE chfschstatus = 'F' and chfdelete = 'N' "
''            SQLQuery = SQLQuery + " AND ((chfenddate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND chfstartdate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
''            SQLQuery = SQLQuery + " AND chfAdfCode = " & imAdfCode & ")"
''            SQLQuery = SQLQuery + " ORDER BY chfCntrNo"
''
''            Set rst = gSQLSelectCall(SQLQuery)
''            If Not rst.EOF Then
''                ckcAllContracts.Visible = True
''                lbcContracts.Visible = True
''                ckcAllAdvt.Value = vbUnchecked
''            End If
''            While Not rst.EOF
''                lbcContracts.AddItem rst!chfCntrNo  ', " & rst(1).Value & ""
''                rst.MoveNext
''            Wend
''        End If
''    Next ilLoop
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmAdvFulfillRpt-lbcVehAff"
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

Private Sub optUseDates_Click(Index As Integer)
    If Index = 0 Then
        Label3.Caption = "Fed Dates- Start"
    Else
        Label3.Caption = "Aired Dates- Start"
    End If
End Sub
'
'       mGetAllISCI - Build up the ISCI with matching ADV and user entered dates with rotation active dates
'
Public Sub mGetAllISCI()
    
    Dim ilLoop As Integer
    Dim imAdfCode As Integer
    
    If Trim(CalOnAirDate.Text) = "" Or Trim(CalOffAirDate.Text) = "" Then Exit Sub
    
    lbcContracts.Clear
    ckcAllContracts.Value = 0
    
    For ilLoop = 0 To lbcVehAff.ListCount - 1
        If lbcVehAff.Selected(ilLoop) Then
            imAdfCode = lbcVehAff.ItemData(ilLoop)
            
            'bring in all isci matching the advertiser selected for the period and rotation matching date
            SQLQuery = "SELECT DISTINCT cpfISCI, cpfCode from cpf_Copy_Prodct_ISCI inner join Cif_Copy_Inventory on cifcpfcode = cpfcode where "
            SQLQuery = SQLQuery + "  ((cifrotenddate >= '" & Format$(CalOnAirDate.Text, sgSQLDateForm) & "' AND cifrotstartdate <= '" & Format$(CalOffAirDate.Text, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " AND cifAdfCode = " & imAdfCode & ")"
            SQLQuery = SQLQuery + " ORDER BY cpfISCI"

            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                ckcAllContracts.Visible = True
                lbcContracts.Visible = True
                ckcAllAdvt.Value = vbUnchecked
            End If
            While Not rst.EOF
                lbcContracts.AddItem rst!cpfISCI
                lbcContracts.ItemData(lbcContracts.NewIndex) = rst!cpfCode
                rst.MoveNext
            Wend
        End If
    Next ilLoop

End Sub
