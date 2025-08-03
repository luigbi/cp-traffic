VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdvPlaceRpt 
   Caption         =   "Advertiser Placement Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
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
         Left            =   6720
         Picture         =   "AffAdvPlaceRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Select Stations from File.."
         Top             =   240
         Width           =   360
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Station"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox ckcDiscrep 
         Caption         =   "Network"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   45
         Top             =   2520
         Width           =   915
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
         Height          =   1230
         ItemData        =   "AffAdvPlaceRpt.frx":056A
         Left            =   3840
         List            =   "AffAdvPlaceRpt.frx":056C
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cbcAVailNames 
         Height          =   315
         ItemData        =   "AffAdvPlaceRpt.frx":056E
         Left            =   2400
         List            =   "AffAdvPlaceRpt.frx":0570
         TabIndex        =   42
         Top             =   1515
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.CheckBox ckcMarkRegional 
         Caption         =   "Highlight Regional Copy"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1800
         Width           =   1560
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
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
      Begin VB.CheckBox ckcExclSpotsLackReg 
         Caption         =   "Exclude Spots Lacking Regional Copy"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   2130
      End
      Begin VB.CheckBox ckcExclTotals 
         Caption         =   "Exclude Station Spot Counts"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1620
      End
      Begin VB.ComboBox cbcExtra 
         Height          =   315
         ItemData        =   "AffAdvPlaceRpt.frx":0572
         Left            =   600
         List            =   "AffAdvPlaceRpt.frx":0574
         TabIndex        =   16
         Top             =   1515
         Width           =   1020
      End
      Begin VB.ComboBox cbcSort 
         Height          =   315
         ItemData        =   "AffAdvPlaceRpt.frx":0576
         Left            =   1080
         List            =   "AffAdvPlaceRpt.frx":0578
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   1125
         Width           =   2340
      End
      Begin VB.CheckBox ckcAllCategories 
         Caption         =   "All Categories"
         Height          =   255
         Left            =   5520
         TabIndex        =   29
         Top             =   2280
         Width           =   1620
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1815
         Index           =   2
         ItemData        =   "AffAdvPlaceRpt.frx":057A
         Left            =   5520
         List            =   "AffAdvPlaceRpt.frx":057C
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Frame frcUseDates 
         BorderStyle     =   0  'None
         Caption         =   "Time Sort by"
         Height          =   225
         Left            =   1380
         TabIndex        =   37
         Top             =   840
         Width           =   1740
         Begin VB.OptionButton optUseDates 
            Caption         =   "Aired"
            Height          =   255
            Index           =   1
            Left            =   750
            TabIndex        =   14
            Top             =   15
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.OptionButton optUseDates 
            Caption         =   "Fed "
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   15
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin VB.CheckBox ckcShowStatus 
         Caption         =   "Show Status Codes"
         Height          =   255
         Left            =   1575
         TabIndex        =   21
         Top             =   2775
         Width           =   1785
      End
      Begin VB.ListBox lbcStatus 
         Height          =   1035
         ItemData        =   "AffAdvPlaceRpt.frx":057E
         Left            =   105
         List            =   "AffAdvPlaceRpt.frx":0580
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         Top             =   3030
         Width           =   3195
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1620
         Index           =   1
         ItemData        =   "AffAdvPlaceRpt.frx":0582
         Left            =   5520
         List            =   "AffAdvPlaceRpt.frx":0584
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   510
         Width           =   1515
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1620
         Index           =   0
         ItemData        =   "AffAdvPlaceRpt.frx":0586
         Left            =   3840
         List            =   "AffAdvPlaceRpt.frx":0588
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   510
         Width           =   1515
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CkcAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "12M"
         Top             =   540
         Width           =   855
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1815
         ItemData        =   "AffAdvPlaceRpt.frx":058A
         Left            =   3840
         List            =   "AffAdvPlaceRpt.frx":058C
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   2565
         Width           =   1515
      End
      Begin VB.CheckBox ckcAllAdvt 
         Caption         =   "All Advertisers"
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lacDiscrep 
         Caption         =   "Non-Compliant Only"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2535
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Avails"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lacStatus 
         Caption         =   "Spot Status Codes"
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   2805
         Width           =   1455
      End
      Begin VB.Label lacExtraInfo 
         Caption         =   "Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1545
         Width           =   555
      End
      Begin VB.Label lacSort 
         Caption         =   "Sort Adv by"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Use Dates/Times"
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   870
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"AffAdvPlaceRpt.frx":058E
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4020
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   570
         Width           =   465
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   34
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
         Left            =   2040
         TabIndex        =   10
         Top             =   210
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
      FormDesignHeight=   6300
      FormDesignWidth =   7575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   33
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   32
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   31
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
Attribute VB_Name = "frmAdvPlaceRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmAdvPlaceRpt (Advertiser Placement)- Create a report of spots (
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

Private Advrst As ADODB.Recordset

Private imckcAllAdvtIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private imckcAllCategoriesIgnore As Integer
Private hmAst As Integer
Private smExtraField As String
Private imExtraIndex As Integer
Dim hmAfr As Integer
Dim tmAfr As AFR
Dim imAfrRecLen As Integer
Dim tmAfrKey As AFRKEY0
Private tmStatusOptions As STATUSOPTIONS

Private Sub cbcExtra_Change()
Dim ilIndex As Integer
    imExtraIndex = cbcExtra.ListIndex
    smExtraField = cbcExtra.List(imExtraIndex)
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

    Select Case ilSetIndex
        Case 0:         'DMA Market
            ckcAllCategories.Caption = "All DMA Markets"
            For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgMarketInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMarketInfo(llLoop).lCode
            Next llLoop
        Case 1:         'DMA Rank
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
        Case 2:         'Format
            ckcAllCategories.Caption = "All Formats"
           For llLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgFormatInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgFormatInfo(llLoop).lCode
            Next llLoop
        Case 3:         'MSAA Market
            ckcAllCategories.Caption = "All MSA Markets"
            For llLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgMSAMarketInfo(llLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgMSAMarketInfo(llLoop).lCode
            Next llLoop
        Case 4:         'MSA Rank
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
        Case 5:         'State
            ckcAllCategories.Caption = "All States"
            For ilLoop = 0 To UBound(tgStateInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgStateInfo(ilLoop).sPostalName) & " (" & Trim$(tgStateInfo(ilLoop).sName) & ")"
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgStateInfo(ilLoop).iCode
            Next ilLoop
        Case 6:         'stations, list box already populated
            lbcSelection(2).Visible = False
            ckcAllCategories.Visible = False
            ckcAllCategories.Value = vbChecked
        Case 7:     'time zone
            ckcAllCategories.Caption = "All Time Zones"
            For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
                lbcSelection(2).AddItem Trim$(tgTimeZoneInfo(ilLoop).sName)
                lbcSelection(2).ItemData(lbcSelection(2).NewIndex) = tgTimeZoneInfo(ilLoop).iCode
            Next ilLoop
        Case 8, 9:    'vehicles/dma name, vehicles/dma rank, already shown. disable the categories listbox
            lbcSelection(2).Visible = False
            ckcAllCategories.Visible = False
            ckcAllCategories.Value = vbChecked
    End Select
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
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imckcAllAdvtIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
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
    If CkcAllVehicles.Value = vbChecked Then
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
    Unload frmAdvPlaceRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    Dim ilTemp As Integer
    Dim llTemp As Long
    Dim sCode As String
    Dim sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim slDateSelection As String
    Dim slRegionOnlySelected As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim llFieldCode As Long
    'Dim NewForm As New frmViewReport
    Dim sStartTime As String
    Dim sEndTime As String
    Dim sCPStatus As String         '12-24-03 option to include non-reported stations
    Dim slNow As String
    Dim ilSelected As Integer
    Dim ilNotSelected As Integer
    Dim slSelection As String
    Dim slAdvtSelected As String
    Dim slTimesSelected As String
    Dim ilShttInx As Integer
    Dim slFormatName As String
    Dim slCreativeName As String
    Dim slISCI As String
    Dim slMSAName As String
    Dim ilMSARank As Integer
    Dim slDMAName As String
    Dim ilDMARank As Integer
    Dim slState As String
    Dim slTimeZone As String
    
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
    Dim iRet As Integer
    Dim ilStatus As Integer
    Dim blStatusOK As Boolean
    'ReDim ilUseShttCodes(1 To 1) As Integer     'shtt codes to include or exclude
    ReDim ilUseShttCodes(0 To 0) As Integer     'shtt codes to include or exclude
    Dim ilIncludeShttCodes As Integer           'true to include codes, false to exclude
    'ReDim ilUseVefCodes(1 To 1) As Integer     'vehicle codes to include or exclude
    ReDim ilUseVefCodes(0 To 0) As Integer     'vehicle codes to include or exclude
    Dim ilIncludeVefCodes As Integer           'true to include codes, false to exclude
    Dim ilFoundVehicle As Integer
    Dim ilFoundStation As Integer
    Dim ilFoundCat As Integer
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    Dim tlStatusOptions As STATUSOPTIONS
    Dim blNetworkDiscrep As Boolean
    Dim blStationDiscrep As Boolean
    Dim slNonCompliantTest As String

    
    On Error GoTo ErrHand
    
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)

    cbcExtra_Change
    ilSortBy = cbcSort.ListIndex            'user selected sort option
    sgCrystlFormula1 = Str$(ilSortBy)       'send to crystal
    
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
    
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    hmAfr = CBtrvTable(ONEHANDLE)
    iRet = btrOpen(hmAfr, "", sgDBPath & "AFR.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If iRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on AFR.mkd"
        iRet = btrClose(hmAfr)
        btrDestroy hmAfr
        Exit Sub
    End If
    imAfrRecLen = Len(tmAfr)
 
    
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
        'gOutputMethod frmAdvPlaceRpt, "PledgeAired.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sStartDate = Format(slInputStartDate, "m/d/yyyy")
    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    
    If optUseDates(1).Value = True Then     'use air dates vs fed dates, need to backup the week and process extra week
                                            'spot may air outside the week
         'sStartDate = DateAdd("d", -7, sStartDate)    '2-10-14 no need to backup with new ast design
         sgCrystlFormula4 = "'A'"
         blUseAirDAte = True                             'always use air date (vs feed date) for spot filter
    Else            'feed date has been disabled
        sgCrystlFormula4 = "'F'"
        blUseAirDAte = False
    End If
    
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Or sEndTime = "12m" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    slTimesSelected = ""
    If (sStartTime <> "12:00AM") Or (sEndTime <> "11:59:59PM") Then    'not 12m-12m
        slTimesSelected = " and (astAirTime >= '" & Format$(sStartTime, sgSQLTimeForm) & "' AND astAirTime <= '" & Format$(sEndTime, sgSQLTimeForm) & "')"
    End If
    
    sStatus = ""
    sCPStatus = ""
           
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
  
    gInitStatusSelections tlStatusOptions               '3-14-12 set all options to exclude

    gSetStatusOptions lbcStatus, ilIncludeNotReported, tlStatusOptions
    tlStatusOptions.iInclResolveMissed = False              'dont show the missed part of a mg or replacement spot since only status codes are shown
  
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
    
    slRegionOnlySelected = ""
    If ckcExclSpotsLackReg.Value = vbChecked Then   'exclude spots lacking regional copy?
        ilIncludeNonRegionSpots = False     'exclude spots without regional copy
        slRegionOnlySelected = " and (astRsfCode > 0) "
    Else
        ilIncludeNonRegionSpots = True      'incl spots with/without regional copy
    End If
    
    '8-19-10 highlight regional spots
    If ckcMarkRegional.Value = vbChecked Then   'highlight regional spots with color?
        sgCrystlFormula10 = "'Y'"
    Else
        sgCrystlFormula10 = "'N'"
    End If
    
    dFWeek = CDate(slInputStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    dFWeek = CDate(slInputEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    If Not ilIncludeNotCarried Then
       ilShowExact = True
    Else
        ilShowExact = False
    End If
    
    'determine which category to filter if applicable
    ilFilterCatBy = -1           'nothing to test to filter
    If Not ckcAllCategories.Value = vbChecked Then      'some selective categories picked
        If ilSortBy = 0 Then            'metro name
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 1 Then        'metro rank
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 2 Then        'format
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 3 Then        'msa name
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 4 Then        'msa rank
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 5 Then        'state
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 6 Then        'station, do nothing (it has a different list box selection)
        ElseIf ilSortBy = 7 Then        'time zone
            ilFilterCatBy = ilSortBy
        ElseIf ilSortBy = 8 Or ilSortBy = 9 Then     'vehicle, do nothing (it has a different list box selection)
        End If
    End If
    
    blNetworkDiscrep = False
    blStationDiscrep = False
    sgCrystlFormula12 = "' '"                  'no compliance discrepancies selected
    slNonCompliantTest = ""
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
    If Trim(sgCrystlFormula12) = "'S'" Then
        slNonCompliantTest = " and (astStationCompliant = 'N')"
    ElseIf Trim(sgCrystlFormula12) = "'N'" Then
        slNonCompliantTest = " and (astAgencyCompliant = 'N')"
    ElseIf Trim(sgCrystlFormula12) = "'B'" Then
        slNonCompliantTest = " and (astStationCompliant = 'N' Or astAgencyCompliant = 'N')"
    End If
    tlSpotRptOptions.bNetworkDiscrep = blNetworkDiscrep
    tlSpotRptOptions.bStationDiscrep = blStationDiscrep


    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    'avail name selectivity hidden for now
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

'    'send the options sent into structure for general rtn
'    tlSpotRptOptions.sStartDate = sStartDate
'    tlSpotRptOptions.sEndDate = sEndDate
'    tlSpotRptOptions.bUseAirDAte = blUseAirDAte
'    tlSpotRptOptions.iAdvtOption = True
'    tlSpotRptOptions.iCreateAstInfo = True
'    tlSpotRptOptions.iShowExact = ilShowExact
'    tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
'    tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
'    tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
'    tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
'
    'gBuildAstStnClr hmAst, sStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2)
    '3-26-12 new subrtn to select on all statuses, get mg/replacement references, and only create afr records that are to be reported
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, iType, lbcSelection(0), lbcSelection(1), ilAdvt, True, lbcVehAff, sGenDate, sGenTime, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2), tmStatusOptions
    'gBuildAstSpotsByStatus hmAst, sStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), ilAdvtOption, lbcVehAff, True, ilShowExact, ilIncludeNonRegionSpots, ilFilterCatBy, lbcSelection(2), tmStatusOptions, blFilterAvailNames, lbcAvailNames
    '2-10-14 use a structure to pass many rpt option variables
    'gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcVehAff, lbcSelection(2), lbcAvailNames
    
    
    slAdvtSelected = ""
    If ckcAllAdvt.Value = vbUnchecked Then
        For ilTemp = 0 To lbcVehAff.ListCount - 1 Step 1
            If lbcVehAff.Selected(ilTemp) Then
                If slAdvtSelected = "" Then
                    slAdvtSelected = " IN (" & lbcVehAff.ItemData(ilTemp)
                Else
                    'has at least one entry
                    slAdvtSelected = slAdvtSelected & "," & lbcVehAff.ItemData(ilTemp)
                End If
            End If
        Next ilTemp
    End If
    
    gObtainCodes lbcSelection(1), ilIncludeShttCodes, ilUseShttCodes()        'build array of which codes to incl/excl
    gObtainCodes lbcSelection(0), ilIncludeVefCodes, ilUseVefCodes()        'build array of which codes to incl/excl
    
    'form up sql query for date selection, always exclude not carried
    slDateSelection = "Select * from ast inner join att on astatfcode = attcode where  attserviceagreement <> 'Y' and aststatus <> 8 and astAirDate >= '" & Format$(sStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEndDate, sgSQLDateForm) & "'" & " "
    If Trim$(slAdvtSelected) <> "" Then
        SQLQuery = slDateSelection & " and astAdfCode " & slAdvtSelected & ")"
    Else
        SQLQuery = slDateSelection
    End If
    
    SQLQuery = SQLQuery & sCPStatus & slTimesSelected & slRegionOnlySelected & slNonCompliantTest  'concatenate the dates, advertisers , status and times selected
    
    gPackDate sgGenDate, tmAfr.iGenDate(0), tmAfr.iGenDate(1)       'generation dsate and time for crystal filter
    tmAfr.lGenTime = gTimeToLong(sgGenTime, False)
    lgSpotCount = 0
    lgSpotCount2 = 0
    Set Advrst = gSQLSelectCall(SQLQuery)
    While Not Advrst.EOF
        lgSpotCount = lgSpotCount + 1       'debugging count
    
        'filter out selectivity
        ilFoundStation = gTestIncludeExclude(Advrst!astShfCode, ilIncludeShttCodes, ilUseShttCodes())       'station selectivity
        ilFoundVehicle = gTestIncludeExclude(Advrst!astVefCode, ilIncludeVefCodes, ilUseVefCodes())       'vehicle selectivity
        ilFoundCat = True           'default to include if filter not test
        
        slFormatName = ""
        slCreativeName = ""
        slMSAName = ""
        ilMSARank = 0
        slDMAName = ""
        ilDMARank = 0
        slState = ""
        slTimeZone = ""
        slISCI = ""
        ilShttInx = gBinarySearchStationInfoByCode(Advrst!astShfCode)
        If ilShttInx < 0 Then
            ilFoundStation = False
        Else
            llTemp = gBinarySearchFmt(CLng(tgStationInfoByCode(ilShttInx).iFormatCode))
            If llTemp >= 0 Then
                slFormatName = Trim$(tgFormatInfo(llTemp).sName)
            End If
            
            llTemp = gBinarySearchMSAMkt(CLng(tgStationInfoByCode(ilShttInx).iMSAMktCode))
            If llTemp >= 0 Then
                slMSAName = Trim$(tgMSAMarketInfo(llTemp).sName)
                ilMSARank = tgMSAMarketInfo(llTemp).iRank
            End If
            
            llTemp = gBinarySearchMkt(CLng(tgStationInfoByCode(ilShttInx).iMktCode))
            If llTemp >= 0 Then
                slDMAName = Trim$(tgMarketInfo(llTemp).sName)
                ilDMARank = tgMarketInfo(llTemp).iRank
            End If
            
            '12/28/15
            'slState = Trim$(tgStationInfoByCode(ilShttInx).sMailState)
            If sgSplitState = "L" Then
                slState = tgStationInfoByCode(ilShttInx).sStateLic
            ElseIf sgSplitState = "P" Then
                slState = tgStationInfoByCode(ilShttInx).sPhyState
            Else
                slState = tgStationInfoByCode(ilShttInx).sMailState
            End If
            slTimeZone = Trim$(tgStationInfoByCode(ilShttInx).sZone)
        End If
        
        If ilFilterCatBy >= 0 Then        'something to filter (-1 indicates no filter testing)
            ilFoundCat = mTestCategory(ilShttInx, ilFilterCatBy, lbcSelection(2))
        End If
        
        ilStatus = gGetAirStatus(Advrst!astStatus)
        blStatusOK = False
        If Advrst!astCPStatus = 0 Then      'not reported?
            If tlStatusOptions.iNotReported Then        'include not reported
                blStatusOK = True
            End If
        ElseIf ilStatus = 0 Then                        'live
            If tlStatusOptions.iInclLive0 Then
                blStatusOK = True
            End If
        ElseIf ilStatus = 1 Then                        'delayed aired
            If tlStatusOptions.iInclDelay1 Then
                blStatusOK = True
            End If
        ElseIf ilStatus = 6 Then
            If tlStatusOptions.iInclAirOutPledge6 Then      'aired outside pledge
                blStatusOK = True
            End If
        ElseIf ilStatus = 7 Then
            If tlStatusOptions.iInclAiredNotPledge7 Then      'aired, not pledge
                blStatusOK = True
            End If
        ElseIf ilStatus = 8 Then
            If tlStatusOptions.iInclNotCarry8 Then      'not carried
                blStatusOK = True
            End If
        ElseIf ilStatus = 9 Then
            If tlStatusOptions.iInclDelayCmmlOnly9 Then      'delay, air comml only
                blStatusOK = True
            End If
        ElseIf ilStatus = 10 Then
            If tlStatusOptions.iInclAirCmmlOnly10 Then      'live, air comml only
                blStatusOK = True
            End If
        ElseIf ((ilStatus = ASTEXTENDED_MG) And (tlStatusOptions.iInclMG11)) Or ((ilStatus = ASTEXTENDED_REPLACEMENT) And (tlStatusOptions.iInclRepl13)) Then
            blStatusOK = True
        ElseIf ilStatus = ASTEXTENDED_BONUS Then
            blStatusOK = True
        ElseIf (ilStatus >= 2 And ilStatus <= 5) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS) Then           '4-12-17 option to ignore missed mg bypass
            If (ilStatus = 2 And tlStatusOptions.iInclMissed2 = True) Or (ilStatus = 3 And tlStatusOptions.iInclMissed3 = True) Or (ilStatus = 4 And tlStatusOptions.iInclMissed4 = True) Or (ilStatus = 5 And tlStatusOptions.iInclMissed5 = True) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS And tlStatusOptions.iInclMissedMGBypass14 = True) Then
                blStatusOK = True
            Else
                blStatusOK = False
            End If
        Else
            blStatusOK = False
        End If
              
        If (ilFoundStation = True) And (ilFoundVehicle = True) And (ilFoundCat = True) And (blStatusOK) Then
            'get the ISCI, all copy has been populated in memory
            llTemp = gBinarySearchCpf(Advrst!astCpfCode)
            If llTemp <> -1 Then
                slISCI = Trim$(tgCpfInfo(llTemp).sISCI)
                slCreativeName = Trim$(tgCpfInfo(llTemp).sCreative)
            End If

            'determine if extra field to be shown
            tmAfr.sCreative = ""             'field to carry the extra field description
            If imExtraIndex > 0 Then
                If imExtraIndex = 1 Then            'creative title
                    tmAfr.sCreative = slCreativeName
                ElseIf imExtraIndex = 2 Then        'format
                    tmAfr.sCreative = slFormatName
                ElseIf imExtraIndex = 3 Then        'msa
                    tmAfr.sCreative = slMSAName
                    tmAfr.iSeqNo = ilMSARank
                ElseIf imExtraIndex = 4 Then        'state
                    tmAfr.sCreative = slState
                ElseIf imExtraIndex = 5 Then        'Time zone
                     tmAfr.sCreative = slTimeZone
                End If
            End If
            
            tmAfr.sISCI = slISCI
            tmAfr.iRegionCopyExists = ilDMARank
            tmAfr.sProduct = slDMAName
            tmAfr.lAstCode = Advrst!astCode         'ast code to process in crystal
            iRet = btrInsert(hmAfr, tmAfr, imAfrRecLen, INDEXKEY0)
            lgSpotCount2 = lgSpotCount2 + 1         'total spots to print
        End If
        Advrst.MoveNext
    Wend

    
    SQLQuery = "Select afrastCode,afrCreative, afrISCI, afrProduct, afrRegionCopyExists, afrSeqno, "
    SQLQuery = SQLQuery & "astAirDate, astAirTime, astStatus, astFeedDate, astFeedTime, "
    SQLQuery = SQLQuery & "shttCallLetters, shttState, adfName, VefName "
    'SQLQuery = SQLQuery & "mktName, mktRank "
    
    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrastcode = astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on astvefcode = vefcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on astshfcode = shttcode "
    'SQLQuery = SQLQuery & " LEFT OUTER JOIN mkt on shttMktCode = mktCode "
   ' SQLQuery = SQLQuery & " LEFT OUTER JOIN cpf_Copy_Prodct_ISCI on astcpfcode = cpfCode "
    SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " and afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "'"

    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfAdvPlace.rpt", "AfAdvPlace"
    
'    cmdReport.Enabled = True            'give user back control to gen, done buttons
'    cmdDone.Enabled = True
'    cmdReturn.Enabled = True
     
    'debugging only for timing tests
    'sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
    
    'remove all the records just printed
    Screen.MousePointer = vbHourglass
    
'    SQLQuery = "DELETE FROM afr "
'    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
'    'cnn.BeginTrans
'    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'        GoSub ErrHand:
'    End If
'    'cnn.CommitTrans
'
    gUserActivityLog "S", sgReportListName & ": Clear AFR"
    iRet = gClearAFR(frmAdvPlaceRpt)
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    gUserActivityLog "E", sgReportListName & ": Clear AFR"

    
'    lgRptSTime1 = timeGetTime
'
'    SQLQuery = "Select Count(afrastcode) FROM afr "
'    SQLQuery = SQLQuery & "WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
'    Set rst = gSQLSelectCall(SQLQuery)
'
'    If Not rst.EOF Then
'        If rst(0).Value > 0 Then
'            llTemp = rst(0).Value
'        Else
'            llTemp = 0
'        End If
'    Else
'        llTemp = 0
'    End If
'
'    Do While llTemp > 0
'        SQLQuery = "delete from afr where datepart(dayofyear, afrGenDate)+afrGenTime+afrAstCode In (Select Top 10000 datepart(dayofyear,afrGenDate)+afrGenTime+afrAstCode From afr "
'        SQLQuery = SQLQuery & "Where afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'        llTemp = llTemp - 10000
'    Loop
'
'    lgRptETime1 = timeGetTime
'    lgRptTtlTime1 = (lgRptETime1 - lgRptSTime1)
    
    'rst.Close
    Advrst.Close
    iRet = btrClose(hmAfr)
    btrDestroy hmAfr

    Screen.MousePointer = vbDefault
        
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Advertiser Placement-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAdvPlaceRpt
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
    gSetFonts frmAdvPlaceRpt
    lacStatusDesc.Move lbcStatus.Left, lbcStatus.Top + lbcStatus.Height + 60
    lacStatusDesc.FontSize = 8
    frmAdvPlaceRpt.Caption = "Advertiser Placement Report - " & sgClientName

    gCenterForm frmAdvPlaceRpt
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
    
    'ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    End If
    rst.Close

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
    
    cbcExtra.AddItem "None"
    cbcExtra.AddItem "Creative Title"
    'cbcExtra.AddItem "DMA Market Name"
    'cbcExtra.AddItem "DMA Market Rank"
    cbcExtra.AddItem "Format"
    cbcExtra.AddItem "MSA Market Rank/Name"
    'cbcExtra.AddItem "MSA Market Rank"
    cbcExtra.AddItem "State"
    cbcExtra.AddItem "Time Zone"
    cbcExtra.ListIndex = 0

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    'ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmAdvPlaceRpt = Nothing
End Sub


Private Sub lbcSelection_Click(Index As Integer)
 
    If Index = 0 Then                          'more vehicle or station selection
       If imckcAllIgnore Then
           Exit Sub
       End If
       If CkcAllVehicles.Value = vbChecked Then
           imckcAllIgnore = True
           CkcAllVehicles.Value = vbUnchecked
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
    If imckcAllAdvtIgnore Then
        Exit Sub
    End If
    If ckcAllAdvt.Value = 1 Then
        imckcAllAdvtIgnore = True
        'ckcAllAdvt.Value = False
        ckcAllAdvt.Value = 0    'chged from false to 0 10-22-99
        imckcAllAdvtIgnore = False
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
'           mTestCategory - this tests category selections from Advertiser Fulfillment report.
'           Categories include State, Format, MSA Market/Rank, DMA Market/Rank and Time Zones
'           <input>  ilShttInx -Station code index
'                     ilFilterCatBy: -1 no testing, process spot, 0 & 1 = market code & rank,
'                     2 = format, 3 & 4 = msa mkt & rank, 5 = state, 7 - time zone
'                     Any other code processes the spot
'                     tlListBox - list of selected categories
'           return - true to process, false to ignore
Private Function mTestCategory(ilShttInx As Integer, ilFilterCatBy As Integer, tlListBox As control) As Integer
Dim ilFoundCat As Integer
Dim ilLoopCat As Integer
Dim llLoopCat As Long   'was using ilLoopCat (integer)
Dim ilFieldCode As Integer
Dim llFieldCode As Long 'was using ilFieldCode (integer)
Dim slState As String * 2
Dim slTemp As String * 2


    On Error GoTo ErrHand:
        ilFoundCat = False
        'ilShttInx = gBinarySearchStationInfoByCode(ilShttCode)
        If ilShttInx <> -1 Then
            If ilFilterCatBy = 0 Or ilFilterCatBy = 1 Then         'market name/rank
                llFieldCode = tgStationInfoByCode(ilShttInx).iMktCode
            ElseIf ilFilterCatBy = 2 Then     'format
                llFieldCode = tgStationInfoByCode(ilShttInx).iFormatCode
            ElseIf ilFilterCatBy = 3 Or ilFilterCatBy = 4 Then     'msa market name/rank
                llFieldCode = tgStationInfoByCode(ilShttInx).iMSAMktCode
            ElseIf ilFilterCatBy = 5 Then     'state
                '12/28/15
                'slState = tgStationInfoByCode(ilShttInx).sMailState        'get only 2 char (postalname)
                If sgSplitState = "L" Then
                    slState = tgStationInfoByCode(ilShttInx).sStateLic
                ElseIf sgSplitState = "P" Then
                    slState = tgStationInfoByCode(ilShttInx).sPhyState
                Else
                    slState = tgStationInfoByCode(ilShttInx).sMailState
                End If
            ElseIf ilFilterCatBy = 7 Then     'time zone
                llFieldCode = tgStationInfoByCode(ilShttInx).iTztCode
            End If
            
            If ilFilterCatBy = 5 Then                       'State
                For ilLoopCat = 0 To tlListBox.ListCount - 1
                    If tlListBox.Selected(ilLoopCat) Then
                        slTemp = Mid(tlListBox.List(llLoopCat), 1, 2)
                        If slTemp = slState Then
                            ilFoundCat = True
                            Exit For
                        End If
                    End If
                Next ilLoopCat
                
            ElseIf (ilFilterCatBy >= 0 And ilFilterCatBy <= 4) Or ilFilterCatBy = 7 Then
                For llLoopCat = 0 To tlListBox.ListCount - 1
                    If tlListBox.Selected(llLoopCat) Then
                        If tlListBox.ItemData(llLoopCat) = llFieldCode Then
                            ilFoundCat = True
                            Exit For
                        End If
                    End If
                Next llLoopCat
                ilFoundCat = ilFoundCat
            Else
                ilFoundCat = True
            End If
        End If
        mTestCategory = ilFoundCat
        Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmAdvPlaceRpt-mTestCategory"
    Exit Function
End Function
