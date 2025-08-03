VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPgmClrRpt 
   Caption         =   "Program Clearance Report"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
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
      Width           =   6960
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6360
         Picture         =   "AffPgmClrRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Select Stations from File.."
         Top             =   1920
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   855
         _extentx        =   1508
         _extenty        =   529
         text            =   "12/12/2022"
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_caldateformat=   1
         font            =   "AffPgmClrRpt.frx":056A
         csi_daynamefont =   "AffPgmClrRpt.frx":0596
         csi_monthnamefont=   "AffPgmClrRpt.frx":05C4
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   -1  'True
         csi_defaultdatetype=   1
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         text            =   "12/12/2022"
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_caldateformat=   1
         font            =   "AffPgmClrRpt.frx":05F2
         csi_daynamefont =   "AffPgmClrRpt.frx":061E
         csi_monthnamefont=   "AffPgmClrRpt.frx":064C
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   -1  'True
         csi_defaultdatetype=   1
      End
      Begin VB.Frame frcFeedTimes 
         Caption         =   "Feed Times"
         Height          =   570
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2220
         Begin VB.OptionButton optFeedTime 
            Caption         =   "Network"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optFeedTimes 
            Caption         =   "Station"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.ListBox lbcStatus 
         Height          =   1620
         ItemData        =   "AffPgmClrRpt.frx":067A
         Left            =   120
         List            =   "AffPgmClrRpt.frx":067C
         MultiSelect     =   2  'Extended
         TabIndex        =   20
         Top             =   2040
         Width           =   3195
      End
      Begin VB.Frame frcMinUnits 
         Caption         =   "Show"
         Height          =   570
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2220
         Begin VB.OptionButton optMinUnits 
            Caption         =   "Units"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optMinUnits 
            Caption         =   "Minutes"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1425
         Index           =   1
         ItemData        =   "AffPgmClrRpt.frx":067E
         Left            =   5280
         List            =   "AffPgmClrRpt.frx":0680
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1425
         Index           =   0
         ItemData        =   "AffPgmClrRpt.frx":0682
         Left            =   3600
         List            =   "AffPgmClrRpt.frx":0684
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   5280
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
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
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtStartTime 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "12M"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   1545
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   2835
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Vehicle, Station, Date, Time"
            Height          =   375
            Index           =   3
            Left            =   90
            TabIndex        =   24
            Top             =   1065
            Width           =   2565
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Advt, Station, Date, Time"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   23
            Top             =   810
            Width           =   2280
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Vehicle, Station, Date, Time"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   22
            Top             =   555
            Width           =   2565
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Station, Date, Time"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   300
            Value           =   -1  'True
            Width           =   2280
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1230
         ItemData        =   "AffPgmClrRpt.frx":0686
         Left            =   3600
         List            =   "AffPgmClrRpt.frx":068D
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   420
         Width           =   3225
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Named Avails"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   165
         Width           =   1935
      End
      Begin VB.Label lacStatusDesc 
         Caption         =   $"AffPgmClrRpt.frx":0694
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
         TabIndex        =   36
         Top             =   3360
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End"
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   630
         Width           =   465
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Times-Start"
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Week of "
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   945
      End
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
Attribute VB_Name = "frmPgmClrRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cprst As ADODB.Recordset
Private missed_rst As ADODB.Recordset
'Private rst_dat As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private Type STATIONPGMCLR
    lFeedTime As Long
    iAnfCode As Integer
    sAnfName As String * 20
    iAired As Integer
    iNotAired As Integer
    iNotRept As Integer
End Type

Dim tmStationPgmClr() As STATIONPGMCLR      'array of times and stats by station & vehicle, ready for output to prepass
Private Const MAXCOL_BYUNIT = 24               'max columns across one page for units option
                                            'set values for indices more than MAXCOL_BYUNIT to 0 (mInsertIntoPCR) if MAXCOL_BYUNIT is changed
Private Const MAXCOL_BYMIN = 20              'max columns across one page for minutes option

Option Explicit

Private imChkListBoxIgnore As Integer
Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private smGenDate As String
Private smGenTime As String
Private smMinUnits As String * 1             '11-29-05 M =  min, U= units
Private imMaxColumns As Integer             '11-29-05
Private hmAst As Integer

Private imIncludeCodes As Integer
Private imUseCodes() As Integer

'
'       determine what spot statuses the user has selected.
'       Send the string to show on the report on which inclusions/exclusions
'       were requested
'
'       <input> lbcStatus - list box containing all the status codes for inclusion/exclusion
'       <output> sStatus - SQL call for the statuses selected
'                slSelection - selection string for crystal
'Public Function mGetSQLStatus(lbcStatus As control, sStatusQuery As String, slSelection As String) As Integer
Public Function mGetSQLStatus(lbcStatus As control, ilStatusSelection() As Integer, slSelection As String) As Integer
Dim i As Integer
Dim ilSelected As Integer
Dim ilNotSelected As Integer
Dim slStatusSelected As String
Dim slStatusNotSelected As String
Dim ilIncludeNotCarried As Integer
Dim slExcludeNR As String
Dim sStatus As String
    sStatus = ""
    slStatusSelected = ""
    slStatusNotSelected = ""
    ilSelected = 0
    ilNotSelected = 0
    ilIncludeNotCarried = True
    slExcludeNR = ""
    For i = 0 To lbcStatus.ListCount - 1 Step 1
        If lbcStatus.Selected(i) Then
            ilSelected = ilSelected + 1
            If Len(slStatusSelected) = 0 Then
                'sStatus = "and ((astStatus = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = "Included:" & lbcStatus.List(i)
            Else
                'sStatus = sStatus & " OR (astStatus = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = slStatusSelected & ", " & lbcStatus.List(i)
            End If
        Else
            ilNotSelected = ilNotSelected + 1
            If lbcStatus.List(i) = "9-Not Carried" Then  'if Not carried not selected, set flag to exclude
                ilIncludeNotCarried = False
                '12-24-13 astPledgeStatus no longer in AST
                'slExcludeNR = "and (astPledgeStatus <> 4 and astPledgeStatus <> 8 ) "
            End If
                
            If Len(slStatusNotSelected) = 0 Then
                slStatusNotSelected = "Excluded:" & lbcStatus.List(i)
            Else
                slStatusNotSelected = slStatusNotSelected & "," & lbcStatus.List(i)
            End If
            ilStatusSelection(UBound(ilStatusSelection)) = lbcStatus.ItemData(i)
            ReDim Preserve ilStatusSelection(0 To UBound(ilStatusSelection) + 1) As Integer
        End If
    Next i
    'sStatus = sStatus & ")"
    'slExcludeNR = ""
    'If Not ilIncludeNotCarried Then         'exclude Not carried  (not fed via pledge)
        'slExcludeNR = "and ((astPledgeStatus <> 4 and astPledgeStatus <> 8 and Mod(astStatus, 100) <> 8)  "

        If lbcStatus.SelCount <= ((lbcStatus.ListCount) / 2) Then
            For i = 0 To lbcStatus.ListCount - 1 Step 1
                If lbcStatus.Selected(i) Then
                    ilSelected = ilSelected + 1
                    If Len(sStatus) = 0 Then
                        sStatus = "and ( (Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                    Else
                        sStatus = sStatus & " OR (Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                    End If
                End If
            Next i
            If Trim$(sStatus) <> "" Then
            sStatus = sStatus & ")"
            End If
        Else
            For i = 0 To lbcStatus.ListCount - 1 Step 1
                If Not lbcStatus.Selected(i) Then
                    ilSelected = ilSelected + 1
                    If Len(sStatus) = 0 Then
                        sStatus = "and ( (Mod(astStatus, 100) <>" & lbcStatus.ItemData(i) & ")"
                    Else
                        sStatus = sStatus & " and (Mod(astStatus, 100) <> " & lbcStatus.ItemData(i) & ")"
                    End If
                End If
            Next i
            If Trim$(sStatus) <> "" Then
            sStatus = sStatus & ")"
        End If
        End If
    'End If
    'sStatusQuery = Trim$(slExcludeNR) & Trim$(sStatus)
    
    'If lbcStatus.ListCount <> ilSelected Then    'not all statuses selected
    If lbcStatus.ListCount <> lbcStatus.SelCount Then   'not all statuses selected
        'determine if inclusion or exclusion
        If ilSelected >= ilNotSelected Then
            slSelection = Trim$(slStatusNotSelected)       'less exclusions, show them
        Else
            slSelection = Trim(slStatusSelected)           'less inclusions, show them
        End If
    Else        'everything included
        slSelection = "Included:  All Statuses"
    End If
    mGetSQLStatus = ilIncludeNotCarried         'special case
End Function
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
'        For i = 0 To lbcListBox.ListCount - 1 Step 1
'            If lbcListBox.Selected(i) Then
'                If Len(slStr) = 0 Then
'                    slStr = "(shttCode = " & lbcListBox.ItemData(i) & ")"
'                Else
'                    slStr = slStr & " OR (shttCode = " & lbcListBox.ItemData(i) & ")"
'                End If
'            End If
'        Next i

        For i = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
            If Len(slStr) = 0 Then
                'slStr = "(astshfCode = " & imUseCodes(i) & ")"
                slStr = "astshfcode in (" & imUseCodes(i)
            Else
                'slStr = slStr & " OR (astshfCode = " & imUseCodes(i) & ")"
                slStr = slStr & "," & imUseCodes(i)
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
    Unload frmPgmClrRpt
End Sub
'******************************************************************
'*  ProgramClearanceRpt - List of spots aired for vehicles and/or stations
'*                if the spot doesnt exist in AST, do not go out to
'*                retrieve it from the LST.  Also, include only those
'*                spots that have been imported or posted
'*
'*  Created 7/30/03 D Hosaka
'
'   12-13-06 Gather data for spots with valid agreement code only; workaround
'           to avoid getting too many spots due to unreferenced astatfcodes
'*
'*  Copyright Counterpoint Software, Inc.
'****************************************************************************
Private Sub cmdReport_Click()
    Dim i As Integer
    Dim sAvailName As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilRet As Integer
    Dim ilErr As Integer
    Dim dFWeek As Date
    Dim slStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'Dim NewForm As New frmViewReport
    Dim sStartTime As String
    Dim sEndTime As String
    Dim slNow As String
    Dim slSelection0 As String       '3-6-05 added selectivity for station and/or vehicle
    Dim slSelection1 As String      '3-6-05 added selectivity for station and/or vehicle
    Dim tlStatsByStation() As STATSBYSTATION
    Dim tlAvailsByVehicle() As AVAILSBYVEHICLE
    Dim ilLoopStats As Integer
    Dim ilFound As Integer
    Dim slTime As String
    Dim llFeedTime As Long
    Dim llFeedDate As Long  '6-12-07
    Dim llAirTime As Long
    Dim llAirDate As Long
    Dim ilAired As Integer
    Dim ilNotAired As Integer
    Dim ilNotReported As Integer
    Dim ilUpperStat As Integer
    Dim ilStatus As Integer
    Dim ilCPStatus As Integer
   
    Dim ilInclude As Integer
    Dim ilVefInx As Integer
    Dim ilVef As Integer
    Dim ilHowManyAvails As Integer
    Dim ilLoInx As Integer
    Dim ilHiInx As Integer
    
    Dim ilPrevShttCode As Integer
    Dim ilWhichColumn As Integer
    Dim ilUpperAvail As Integer
    Dim ilLoopAvails As Integer
    Dim ilHowManySets As Integer
    Dim ilSetCount As Integer
    Dim ilStation As Integer
    Dim ilUpperStation As Integer
    Dim ilTimeInx As Integer
    Dim ilStatInx As Integer
    Dim ilTemp As Integer
    Dim ilHowManyDefined As Integer
    Dim ilHowManySelected As Integer
    Dim slAvailsSelected As String
    Dim ilWhichSet As Integer
    Dim llEndDate As Long
    Dim llStartDate
    Dim slUserTimes As String
    Dim ilMaxColumns As Integer
    Dim slSQLForMinUnits As String
    Dim ilLen As Integer                'spot length if minutes option, otherwise 1
    Dim llTotalAired As Long            'total minutes or units aired for a station/vehicle
    Dim llTotalNotRept As Long          'total minutes or units not reported for a station/vehicle
    Dim llTotalNotAired As Long         'total minutes or units not aired for station/vehicle
    Dim sStatus As String
    Dim slSelection As String
    
    Dim slZone As String
    Dim ilLocalAdj As Integer
    Dim ilZoneFound As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilZone As Integer
    Dim ilStaCode As Long
    Dim ilVefArrayInx As Integer
    Dim llDate As Long
    Dim llTime As Long
    Dim llSpotTime As Long
    Dim llPrevAttCode As Long
    'Dim ilDACode As Integer
    Dim slDACode As String
    Dim ilIncludeNotCarried As Integer
    Dim ilPledgedStatus As Integer
    Dim ilIncludeANF As Integer
    Dim ilShowOther As Integer
    Dim ilAnfcode As Integer
    Dim iAdfCode As Integer
    Dim sSDate As String
    Dim llLoopOnAST As Long
    Dim ilFoundStation As Integer
    Dim iVef As Integer
    Dim iRet As Integer
    ReDim ilStatusSelection(0 To 0) As Integer
    'ReDim imUseCodes(1 To 1) As Integer
    ReDim imUseCodes(0 To 0) As Integer

    
    On Error GoTo ErrHand
    sStartDate = Trim$(CalOnAirDate.Text)
    
    'sEndDate = Trim$(CalOffAirDate.Text)
    'test for valid date
    If gIsDate(sStartDate) = False Or (Len(Trim$(sStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    'date must be a monday
    If Weekday(sStartDate, vbSunday) <> vbMonday Then
        gMsgBox "Date Must be a Monday", vbOKOnly
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    
    
    llEndDate = (DateValue(gAdjYear(CalOnAirDate.Text))) + 6
    sEndDate = Format$(llEndDate, "m/d/yy")
    
    sStartDate = Format(sStartDate, "m/d/yy")
    llStartDate = DateValue(sStartDate)
    sEndDate = Format$(llEndDate, "m/d/yy")

    sStartTime = txtStartTime.Text
    If (gIsTime(sStartTime) = False) Or (Len(Trim$(sStartTime)) = 0) Then   'Time not valid.
        Beep
        gMsgBox "Please enter a valid start time (h:mm:ssA/P)", vbCritical
        txtStartTime.SetFocus
        Exit Sub
    End If
    
    slStr = gConvertTime(sStartTime)
    If Second(slStr) = 0 Then
        slStr = Format$(slStr, sgShowTimeWOSecForm)
    Else
        slStr = Format$(slStr, sgShowTimeWSecForm)
    End If
    llStartTime = gTimeToLong(slStr, False)
    slUserTimes = Trim$(slStr) & "-"
    
    sEndTime = txtEndTime.Text
    If (gIsTime(sEndTime) = False) Or (Len(Trim$(sEndTime)) = 0) Then   'Time not valid.
        Beep
        gMsgBox "Please enter a valid end time (h:mm:ssA/P)", vbCritical
        txtEndTime.SetFocus
        Exit Sub
    End If
     slStr = gConvertTime(sEndTime)
    If Second(slStr) = 0 Then
        slStr = Format$(slStr, sgShowTimeWOSecForm)
    Else
        slStr = Format$(slStr, sgShowTimeWSecForm)
    End If
    llEndTime = gTimeToLong(slStr, True)
    slUserTimes = slUserTimes & Trim$(slStr)
    
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmPgmClrRpt, "PgmClr.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    ilRet = gPopShttInfo
    If Not ilRet Then
        gMsgBox "gPopShttInfo failed, call Counterpoint"
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'format the sql query for the selection of spot statuses
    'get the description of spot statuses (included/excluded) to show on report
    ilIncludeNotCarried = mGetSQLStatus(lbcStatus, ilStatusSelection(), slSelection)

    sgCrystlFormula5 = slSelection
    
    '11-28-05 Determine if showing clearnace by minutes or units
     If optMinUnits(1).Value = True Then        'test if units requested
        smMinUnits = "U"                      'units option
        ilMaxColumns = MAXCOL_BYUNIT
    Else
        smMinUnits = "M"
        ilMaxColumns = MAXCOL_BYMIN
    End If
    
    slSQLForMinUnits = "SELECT astStatus, astCPStatus,astFeedTime,astShfCode,astcode, astFeedDate, astvefCode, astAtfCode,   lstAnfCode ,lstlen, attCode, attPledgeType from ast "
    slSQLForMinUnits = slSQLForMinUnits & " left outer join lst on astlsfcode = lstcode inner join att on astatfcode = attcode "
        
    sStartTime = gConvertTime(sStartTime)
    If sEndTime = "12M" Then
        sEndTime = "11:59:59PM"
    End If
    sEndTime = gConvertTime(sEndTime)
    
    smGenDate = Format$(gNow(), "m/d/yyyy")          'generation date & time for prepass filter
    smGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
    slAvailsSelected = ""
    ilHowManyDefined = lbcVehAff.ListCount
    ilHowManySelected = lbcVehAff.SelCount
    If chkListBox.Value = vbChecked Then
        slAvailsSelected = "Include: All Named Avails"
    Else
        If ilHowManySelected > ilHowManyDefined / 2 Then    'more than half selected
            For ilTemp = 0 To lbcVehAff.ListCount - 1       'show exclusions
                If Not lbcVehAff.Selected(ilTemp) Then
                    If slAvailsSelected = "" Then
                        slAvailsSelected = "Exclude: " & lbcVehAff.List(ilTemp)
                    Else
                        slAvailsSelected = slAvailsSelected & ", " & lbcVehAff.List(ilTemp)
                    End If
                End If
            Next ilTemp
        Else                                                'show inclusions
            For ilTemp = 0 To lbcVehAff.ListCount - 1       'show exclusions
                If lbcVehAff.Selected(ilTemp) Then
                    If slAvailsSelected = "" Then
                        slAvailsSelected = "Include: " & lbcVehAff.List(ilTemp)
                    Else
                        slAvailsSelected = slAvailsSelected & ", " & lbcVehAff.List(ilTemp)
                    End If
                End If
            Next ilTemp
        End If
    End If
    sgCrystlFormula4 = "'" & Trim$(slAvailsSelected) & "'"         'prepare for formula to pass to crystal
    
    sAvailName = ""
    slSelection0 = ""       '3-6-05 additonal vehicle selection depending on sort option
    slSelection1 = ""       '3-6-05 more station selectivity depending on sort option
     
    'determine named avails, vehicle and station selectivity
    If chkListBox.Value = 0 Then    '= 0 Then                        'User did NOT select all stations
        For i = 0 To lbcVehAff.ListCount - 1 Step 1
            If lbcVehAff.Selected(i) Then
                If Len(sAvailName) = 0 Then
                    sAvailName = "(anfCode = " & lbcVehAff.ItemData(i) & ")"
                Else
                    sAvailName = sAvailName & " OR (anfCode = " & lbcVehAff.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    slSelection0 = mGetVehicleSelection(CkcAll.Value, lbcSelection(0))     'format sql call of selected vehicles
    gObtainCodes lbcSelection(1), imIncludeCodes, imUseCodes()        'build array of which station codes to incl/excl
    slSelection1 = mGetStationSelection(ckcAllStations.Value, lbcSelection(1))  'selected stations

    
    dFWeek = CDate(sStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dFWeek = CDate(sEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    bgTaskBlocked = False
    sgTaskBlockedName = sgReportListName
    
    'loop and process all stations for one vehicle at a time based on the cptts
    For ilVefInx = 0 To lbcSelection(0).ListCount - 1
        If lbcSelection(0).Selected(ilVefInx) Then
            llPrevAttCode = -1
            ilVef = lbcSelection(0).ItemData(ilVefInx)
            
            ReDim tlStatsByStation(0 To 0) As STATSBYSTATION
            ReDim tlAvailsByVehicle(0 To 0) As AVAILSBYVEHICLE
            ilUpperStat = 0
            ilUpperAvail = 0
            ilUpperStation = 0
            
            iAdfCode = -1           'assume to retrieve all advertiser for vehicle or station option
            sSDate = gObtainPrevMonday(sStartDate)      'cp returned status is based on weeks
             
            'Get CPTT so that Stations requiring CP can be obtained
            SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
            SQLQuery = SQLQuery & " from cptt inner join shtt on cpttshfcode = shttcode inner join att on cpttatfcode = attcode "
            SQLQuery = SQLQuery & " where ( cpttVefCode = " & lbcSelection(0).ItemData(ilVefInx)
            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
            
            Set cprst = gSQLSelectCall(SQLQuery)
            While Not cprst.EOF
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cprst!cpttCode
                tgCPPosting(0).iStatus = cprst!cpttStatus
                tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cprst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cprst!attTimeType
                tgCPPosting(0).iVefCode = cprst!cpttvefcode
                tgCPPosting(0).iShttCode = cprst!shttCode
                tgCPPosting(0).sZone = cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                
                ilFoundStation = False
                If imIncludeCodes Then
                    'For ilTemp = 1 To UBound(imUseCodes) - 1 Step 1
                    For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
                        If imUseCodes(ilTemp) = cprst!shttCode Then
                            ilFoundStation = True
                            Exit For
                        End If
                    Next ilTemp
                Else
                    ilFoundStation = True        '8/23/99 when more than half selected, selection fixed
                    'For ilTemp = 1 To UBound(imUseCodes) - 1 Step 1
                    For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
                        If imUseCodes(ilTemp) = cprst!shttCode Then
                            ilFoundStation = False
                            Exit For
                        End If
                    Next ilTemp
                End If
                'Create AST records
                If ilFoundStation Then
                    ReDim tmAstInfo(0 To 0) As ASTINFO
                    igTimes = 1 'By Week
                    iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, True, False, , , True) 'False)  9-22-09 always create the ast records

                    'Cycle thru all the ast records by vehicle and station.  Process one vehicle and station at a time, and build array of
                    'of feed times and their stats (aired, not aired, not reported).  When theres a change in station/vehicle, write
                    'out that stations/vehicle data to prepass temporary file
            
                For llLoopOnAST = LBound(tmAstInfo) To UBound(tmAstInfo) - 1
    
                    slTime = tmAstInfo(llLoopOnAST).sFeedTime
                    llTime = gTimeToLong(slTime, False)
                    llDate = DateValue(tmAstInfo(llLoopOnAST).sFeedDate)
                    If optFeedTime(0).Value = True Then             'use network feed time vs station feed time
                        ilStaCode = gBinarySearchShtt(tmAstInfo(llLoopOnAST).iShttCode)
                        slZone = UCase$(Trim$(tgShttInfo1(ilStaCode).shttTimeZone))
                        ilVefArrayInx = gBinarySearchVef(CLng(tmAstInfo(llLoopOnAST).iVefCode))
                        ilLocalAdj = 0
                        ilZoneFound = False
                        ilNumberAsterisk = 0
                        ' Adjust time zone properly.
                        If Len(slZone) <> 0 Then
                            'Get zone
                            For ilZone = LBound(tgVehicleInfo(ilVefArrayInx).sZone) To UBound(tgVehicleInfo(ilVefArrayInx).sZone) Step 1
                                If Trim$(tgVehicleInfo(ilVefArrayInx).sZone(ilZone)) = slZone Then
                                    If tgVehicleInfo(ilVefArrayInx).sFed(ilZone) <> "*" Then
                                        slZone = tgVehicleInfo(ilVefArrayInx).sZone(tgVehicleInfo(ilVefArrayInx).iBaseZone(ilZone))
                                        ilLocalAdj = tgVehicleInfo(ilVefArrayInx).iLocalAdj(ilZone)
                                        ilZoneFound = True
                                    End If
                                    Exit For
                                End If
                            Next ilZone
                            For ilZone = LBound(tgVehicleInfo(ilVefArrayInx).sZone) To UBound(tgVehicleInfo(ilVefArrayInx).sZone) Step 1
                                If tgVehicleInfo(ilVefArrayInx).sFed(ilZone) = "*" Then
                                    ilNumberAsterisk = ilNumberAsterisk + 1
                                End If
                            Next ilZone
                        End If
                        If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
                            slZone = ""
                        End If
                        ilLocalAdj = -1 * ilLocalAdj
                        llSpotTime = llTime + 3600 * ilLocalAdj
                        If llSpotTime < 0 Then
                            llSpotTime = llSpotTime + 86400
                            llDate = llDate - 1
                        ElseIf llSpotTime > 86400 Then
                            llSpotTime = llSpotTime - 86400
                            llDate = llDate + 1
                        End If
                        
                         If llPrevAttCode <> tmAstInfo(llLoopOnAST).lAttCode Then
                            llPrevAttCode = tmAstInfo(llLoopOnAST).lAttCode
    
                        End If
                        
                        If slDACode = "C" Then          'CD (C) no longer applicable, these are for old records
                            llSpotTime = gTimeToLong(tmAstInfo(llLoopOnAST).sAirTime, False)
                            llDate = DateValue(tmAstInfo(llLoopOnAST).sFeedDate)
                        End If
                        
                        llFeedTime = llSpotTime         'new adjust time based on conversion
                    Else                'use station feed time
                        llFeedTime = llTime
                    End If
                                        
                    ilStatus = gGetAirStatus(tmAstInfo(llLoopOnAST).iStatus)
                    ilCPStatus = tmAstInfo(llLoopOnAST).iCPStatus
                    If smMinUnits = "U" Then         'SDF wasnt retrieved if by Units
                        ilLen = 1
                    Else
                        ilLen = tmAstInfo(llLoopOnAST).iLen
                    End If
                    
                    '2-6-13 put anfcode in temporary field, cannot use the result set if null
                    'If IsNull(rst!lstAnfCode) Then
                    '    ilAnfcode = 0
                    'Else
                        'ilAnfcode = rst!lstAnfCode
                        ilAnfcode = tmAstInfo(llLoopOnAST).iAnfCode
                    'End If
                        
                    ilInclude = mTestFilterTimeandStation(tmAstInfo(llLoopOnAST).iShttCode, llFeedTime, llStartTime, llEndTime)
                    ilIncludeANF = mTestFilterANFCode(ilAnfcode, ilShowOther)
    
                    ilPledgedStatus = True
    
                    For ilTemp = LBound(ilStatusSelection) To UBound(ilStatusSelection) - 1
                        If ilStatus = ilStatusSelection(ilTemp) Then
                            ilPledgedStatus = False
                            Exit For
                        End If
                    Next ilTemp
                    'ilInclude = 0: valid station
                    'ilincludeANF >= 0:  valid named avail, ilInclude ANF -1:  no avail found selected
                    'ilShowOther - no avail found, show in Other column (invalid named avail)
                    'If ilInclude >= 0 And (tgStatusTypes(gGetAirStatus(rst!astPledgeStatus)).iPledged <> 2) Then
                    If (ilInclude = 0) And (ilIncludeANF >= 0 Or ilShowOther = True) And ilPledgedStatus = True Then
                        '4-17-08 use astPledgeStatus as index into statustype array; if 2 its not carried so ignore spot
                        '-1 indicates no named avail match or no station match
                        'other ilinclude is the named avail index selection to the list box
                        'save this value to put into table to retrieve the description for crystal report header
                        
                        'create the number of unique avail times for vehicle report header.  Some stations may not
                        'carry all programs (avails), but all avails need to be shown across the page
                        'For example:
                        '     6:01A      6:15A       7:01A        7:15A        8:01A         8:15A
                        'KABC                           X            X           X             X
                        'KLOS   X           X           X            X
                        'KNX                            X            X
                        
                        ilFound = False
                        If (ilIncludeANF >= 0) Then         'valid named avail to include
                            For ilLoopAvails = 0 To UBound(tlAvailsByVehicle) - 1
                                If tlAvailsByVehicle(ilLoopAvails).lFeedTime = llFeedTime And tlAvailsByVehicle(ilLoopAvails).iAnfCode = ilAnfcode Then         '2-6-13
                                    ilFound = True
                                    'save index to entry (illoopavails)
                                    Exit For
                                End If
                            Next ilLoopAvails
                            
                            If Not ilFound Then
                                slStr = Trim$(Str$(llFeedTime))
                                Do While Len(slStr) < 5
                                    slStr = "0" & slStr
                                Loop
                                tlAvailsByVehicle(ilUpperAvail).sKey = slStr
                                tlAvailsByVehicle(ilUpperAvail).lFeedTime = llFeedTime     'avail time
                                tlAvailsByVehicle(ilUpperAvail).iAnfCode = ilAnfcode    '2-6-13 rst!lstAnfCode     'named avail
                                tlAvailsByVehicle(ilUpperAvail).iAnfInx = ilInclude     'named avail index into list box for retrieval of named avail description
                                tlAvailsByVehicle(ilUpperAvail).iAnfInx = ilIncludeANF     'named avail index into list box for retrieval of named avail description
                                ilLoopAvails = ilUpperAvail
                                ilUpperAvail = ilUpperAvail + 1
                                ReDim Preserve tlAvailsByVehicle(0 To ilUpperAvail) As AVAILSBYVEHICLE
                            End If
                        Else
                            If ilShowOther Then     'show the data in Other column if its an invalid avil code
                            'force in Other column as 11:59:59PM and 32000
                                llFeedTime = gTimeToLong("11:59:59PM", True)
                                For ilLoopAvails = 0 To UBound(tlAvailsByVehicle) - 1
                                    If tlAvailsByVehicle(ilLoopAvails).lFeedTime = llFeedTime And tlAvailsByVehicle(ilLoopAvails).iAnfCode = 32000 Then
                                        ilFound = True
                                        'save index to entry (illoopavails)
                                        Exit For
                                    End If
                                Next ilLoopAvails
                                
                                If Not ilFound Then
                                    slStr = Trim$(Str$(llFeedTime))
                                    Do While Len(slStr) < 5
                                        slStr = "0" & slStr
                                    Loop
                                    tlAvailsByVehicle(ilUpperAvail).sKey = slStr
                                    tlAvailsByVehicle(ilUpperAvail).lFeedTime = llFeedTime     'avail time
                                    tlAvailsByVehicle(ilUpperAvail).iAnfCode = 32000     'named avail
                                    tlAvailsByVehicle(ilUpperAvail).iAnfInx = 32000     'named avail index , code for "Other"
                                    ilLoopAvails = ilUpperAvail
                                    ilUpperAvail = ilUpperAvail + 1
                                    ReDim Preserve tlAvailsByVehicle(0 To ilUpperAvail) As AVAILSBYVEHICLE
                                End If
                            End If
                        End If
                        
                        'create the statistics for the station (aired, not aired, not reported) by time
                        ilFound = False
                        'the spot was necessary to build the feed times based on the feed date, but spot has to be aired within the user feed dates.  If not, ignore the spot since not aired in the week requeseted
                        slTime = tmAstInfo(llLoopOnAST).sAirTime
                        llAirTime = gTimeToLong(slTime, False)
                        llAirDate = DateValue(tmAstInfo(llLoopOnAST).sAirDate)
                        If (llAirDate >= llStartDate And llAirDate <= llEndDate) And (llAirTime >= llStartTime And llAirTime <= llEndTime) Then         'air date and times within the requested dates
    
                            For ilLoopStats = 0 To UBound(tlStatsByStation) - 1
                                If tlStatsByStation(ilLoopStats).iShfCode = tmAstInfo(llLoopOnAST).iShttCode And tlStatsByStation(ilLoopStats).lFeedTime = llFeedTime And (tlStatsByStation(ilLoopStats).iAnfCode = tmAstInfo(llLoopOnAST).iAnfCode Or tlStatsByStation(ilLoopStats).iAnfCode = 32000) Then
                                    'determine status of spot
                                    mGetStatus ilStatus, ilCPStatus, ilAired, ilNotAired, ilNotReported, ilLen, tmAstInfo(llLoopOnAST).lCode
        
                                    'accum stats
                                    tlStatsByStation(ilLoopStats).iAired = tlStatsByStation(ilLoopStats).iAired + ilAired
                                    tlStatsByStation(ilLoopStats).iNotAired = tlStatsByStation(ilLoopStats).iNotAired + ilNotAired
                                    tlStatsByStation(ilLoopStats).iNotReported = tlStatsByStation(ilLoopStats).iNotReported + ilNotReported
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoopStats
                                
                            If Not ilFound Then         'if entry not found, create new one for this feed time
                                mGetStatus ilStatus, ilCPStatus, ilAired, ilNotAired, ilNotReported, ilLen, tmAstInfo(llLoopOnAST).lCode
                                tlStatsByStation(ilUpperStat).iShfCode = tmAstInfo(llLoopOnAST).iShttCode
                                slStr = Trim$(Str$(tmAstInfo(llLoopOnAST).iShttCode))
                                Do While Len(slStr) < 5
                                    slStr = "0" & slStr
                                Loop
                                tlStatsByStation(ilUpperStat).sKey = slStr
                                tlStatsByStation(ilUpperStat).lFeedTime = llFeedTime
                                If ilIncludeANF >= 0 Then
                                    tlStatsByStation(ilUpperStat).iAnfCode = ilAnfcode      '2-6-13 rst!lstAnfCode
                                Else
                                    tlStatsByStation(ilUpperStat).iAnfCode = 32000
                                End If
                                tlStatsByStation(ilUpperStat).iAired = tlStatsByStation(ilUpperStat).iAired + ilAired
                                tlStatsByStation(ilUpperStat).iNotAired = tlStatsByStation(ilUpperStat).iNotAired + ilNotAired
                                tlStatsByStation(ilUpperStat).iNotReported = tlStatsByStation(ilUpperStat).iNotReported + ilNotReported
                                ilUpperStat = ilUpperStat + 1
                                ReDim Preserve tlStatsByStation(0 To ilUpperStat) As STATSBYSTATION
                            End If
                        End If
                    End If
                Next llLoopOnAST
             End If
               cprst.MoveNext
            Wend
            
            ilHowManySets = UBound(tlAvailsByVehicle) / ilMaxColumns   'determine  how many sets of avails across 1 page, 28 max across on a page for this vehicle
            If ilHowManySets * ilMaxColumns < UBound(tlAvailsByVehicle) Then
                ilHowManySets = ilHowManySets + 1               'adjust for the remainder of avails that dont fit across the page
            End If
 
            'sort the avail feed times so that they will come out in correct time order on the report headers
            If UBound(tlAvailsByVehicle) - 1 > 0 Then
                ArraySortTyp fnAV(tlAvailsByVehicle(), 0), UBound(tlAvailsByVehicle), 0, LenB(tlAvailsByVehicle(0)), 0, LenB(tlAvailsByVehicle(0).sKey), 0
            End If
            
            'sort the stats by station so all records for one stations are processed together to prevent multiple passes by station
            If UBound(tlStatsByStation) - 1 > 0 Then
                ArraySortTyp fnAV(tlStatsByStation(), 0), UBound(tlStatsByStation), 0, LenB(tlStatsByStation(0)), 0, LenB(tlStatsByStation(0).sKey), 0
            End If

            'cycled thru all spots for the vehicle,
            'dump all stations for the current vehicle to temporary prepass for crystal reports
            ilLoInx = 0
            ilHiInx = UBound(tlStatsByStation) - 1
            'find out how many avails there are for this station
            ilPrevShttCode = tlStatsByStation(0).iShfCode
            Do While ilHiInx < UBound(tlStatsByStation)
                'determine how many spots there are for this station
                'set its starting and ending points in the array
                For ilHowManyAvails = ilLoInx To UBound(tlStatsByStation) - 1
                    If ilPrevShttCode = tlStatsByStation(ilHowManyAvails).iShfCode Then
                       ilHiInx = ilHowManyAvails
                    Else
                        Exit For
                    End If
                Next ilHowManyAvails
                
                'create an array for as many complete sets needs
                ReDim tmStationPgmClr(0 To ilHowManySets * ilMaxColumns) As STATIONPGMCLR
                'initialize the string fields for columns not used by this station
                llTotalAired = 0
                llTotalNotRept = 0
                llTotalNotAired = 0
                For ilTemp = 0 To UBound(tmStationPgmClr)
                    tmStationPgmClr(ilTemp).sAnfName = ""
                Next ilTemp

                'each entry (in tlStatsbyStation) represents an avail(feed time) for the station,
                'with its stats for aired, not aired, or not reported.
                'build the array of stats where the avail should show on the page; but build
                'in one long array; then write out each set of 28 avails in one record.
                'For example, if station carries every program for the vehicle, it will print 28
                'avails across the page, and skip to a new page to print the next 28 (or whatever # of
                'avails are remaining)
                'Not all stations carry allprograms, therefore some columns may be skipped without any data.
                
                '9-1-06 change outer loop to loop by station avails within vehicle avails
                For ilTemp = LBound(tlAvailsByVehicle) To UBound(tlAvailsByVehicle) - 1
                    For ilWhichColumn = ilLoInx To ilHiInx      'loop thru the station stats array (lo & hi inx represent span of 1 stations info)
                    'put the stations stats info into proper column based on the array availsbyVehicle, which is
                    'in the sorted order to to be shown on the report.
                        If tlAvailsByVehicle(ilTemp).lFeedTime = tlStatsByStation(ilWhichColumn).lFeedTime And tlAvailsByVehicle(ilTemp).iAnfCode = tlStatsByStation(ilWhichColumn).iAnfCode Then
                            tmStationPgmClr(ilTemp).lFeedTime = tlAvailsByVehicle(ilTemp).lFeedTime
                            If tlAvailsByVehicle(ilTemp).iAnfInx = 32000 Then
                                tmStationPgmClr(ilTemp).sAnfName = "Other"
                            Else
                                tmStationPgmClr(ilTemp).sAnfName = gFileNameFilter(Trim$(lbcVehAff.List(tlAvailsByVehicle(ilTemp).iAnfInx)))    'remove illegal characters
                            End If
                            'tmStationPgmClr(ilTemp).sAnfName = Trim$(lbcVehAff.List(tlAvailsByVehicle(ilTemp).iAnfInx))
                            tmStationPgmClr(ilTemp).iAired = tlStatsByStation(ilWhichColumn).iAired
                            tmStationPgmClr(ilTemp).iNotAired = tlStatsByStation(ilWhichColumn).iNotAired
                            tmStationPgmClr(ilTemp).iNotRept = tlStatsByStation(ilWhichColumn).iNotReported
                            'accumulate the station totals for minutes or units aired, not reported or not aired
                            llTotalAired = llTotalAired + tmStationPgmClr(ilTemp).iAired
                            llTotalNotRept = llTotalNotRept + tmStationPgmClr(ilTemp).iNotRept
                            llTotalNotAired = llTotalNotAired + tmStationPgmClr(ilTemp).iNotAired

                            Exit For
                        Else
                            tmStationPgmClr(ilTemp).lFeedTime = tlAvailsByVehicle(ilTemp).lFeedTime
                            If tlAvailsByVehicle(ilTemp).iAnfInx = 32000 Then   'either a 0 or invalid named avail internal code
                                tmStationPgmClr(ilTemp).sAnfName = "Other"
                            Else
                                tmStationPgmClr(ilTemp).sAnfName = gFileNameFilter(Trim$(lbcVehAff.List(tlAvailsByVehicle(ilTemp).iAnfInx)))    'remove illegal characters
                            End If
                        End If
                    Next ilWhichColumn

                Next ilTemp
            
                'write the record(s) to disk that crystal will report from
                ilErr = mInsertIntoPCR(ilVef, ilPrevShttCode, ilHowManySets, ilMaxColumns, llTotalAired, llTotalNotRept, llTotalNotAired)
                If ilErr Then           'if error,  message should have been reported in mInsertintoPCR
                    Exit Sub
                End If

                ilPrevShttCode = tlStatsByStation(ilHiInx + 1).iShfCode   'next station  code to process
                ilLoInx = ilHiInx + 1
                ilHiInx = ilLoInx
            Loop
        End If                          'selected
    Next ilVefInx                     'loop on vehicles selection box
    
    gCloseRegionSQLRst
    
    If bgTaskBlocked And igReportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Report generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If

    bgTaskBlocked = False
    sgTaskBlockedName = ""
    
    'send formula to crystal for heading selectivtiy
    If sEndTime = "11:59:59PM" Then
        slStr = sStartTime & "-12M"
    Else
        slStr = sStartTime & "-" & sEndTime
    End If
    sgCrystlFormula1 = "'" & "For week of " & sStartDate & ", " & slUserTimes
    If optFeedTime(0).Value = True Then         'use network times
        sgCrystlFormula1 = sgCrystlFormula1 & " using Network Feed Times" & "'"
    Else
        sgCrystlFormula1 = sgCrystlFormula1 & " using Station Feed Times" & "'"
    End If
        
     'Prepare records to pass to Crystal
        SQLQuery = "SELECT *"
        SQLQuery = SQLQuery & " FROM pcr, shtt, VEF_Vehicles "
        SQLQuery = SQLQuery + " WHERE (vefCode = pcrvefCode"
        SQLQuery = SQLQuery + " AND shttCode = pcrshfcode "
        SQLQuery = SQLQuery + " AND pcrGenDate = '" & Format$(smGenDate, sgSQLDateForm) & "' AND pcrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(smGenTime, False))))) & "')"
        
    gUserActivityLog "E", sgReportListName & ": Prepass"
    If smMinUnits = "U" Then            'by units
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfPgmClr.rpt", "AfPgmClr"
    Else                                'by minutes
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfPgmClrMin.rpt", "AfPgmClrMin"
    End If
    
    
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime

    gUserActivityLog "S", sgReportListName & ": Clear PCR"
   
    ' Delete the info we stored in the PCR prepass table
    SQLQuery = "DELETE FROM PCR"
    SQLQuery = SQLQuery & " WHERE (pcrGenDate = '" & Format$(smGenDate, sgSQLDateForm) & "' " & "and pcrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(smGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "PgmClrRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans

    cmdReport.Enabled = True
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    gUserActivityLog "E", sgReportListName & ": Clear PCR"
    
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmPgmClrRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmPgmClrRpt
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

Private Sub Form_Initialize()
Dim ilRet As Integer
Dim ilHalf As Integer
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts frmPgmClrRpt
    lacStatusDesc.Move lbcStatus.Left, lbcStatus.Top + lbcStatus.Height + 60
    lacStatusDesc.FontSize = 8
    ilHalf = (frcSelection.Height - chkListBox.Height - CkcAll.Height - 240) / 2
    lbcVehAff.Move chkListBox.Left, chkListBox.Top + chkListBox.Height + 30
    lbcVehAff.Height = ilHalf - 120
    CkcAll.Top = lbcVehAff.Top + lbcVehAff.Height + 30
    ckcAllStations.Top = CkcAll.Top
    lbcSelection(0).Move CkcAll.Left, CkcAll.Top + CkcAll.Height + 30
    lbcSelection(0).Height = ilHalf
    lbcSelection(1).Top = lbcSelection(0).Top
    lbcSelection(1).Height = ilHalf
    

    gCenterForm frmPgmClrRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim ilRet As Integer
    Dim ilHideNotCarried As Integer
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    
    imChkListBoxIgnore = False
    frmPgmClrRpt.Caption = "Program Clearance Report - " & sgClientName
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    ilRet = gPopAvailNames
    'SQLQuery = "SELECT * From Site Where siteCode = 1"
    'Set rst = gSQLSelectCall(SQLQuery)
    'If Not rst.EOF Then
    '    sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    'End If
    
    
    'frcSortBy.Height = 975
    'optSortby(1).Top = 300               'vehicle, station, date, time
    'optSortby(3).Top = 555               'advt, vehicle, station, date time
    'optSortby(0).Visible = False
    'optSortby(2).Visible = False
    'optSortby(1).Value = True
    'optSortby(1).Visible = True          'vehicle, station, date, time
    'optSortby(3).Visible = True           'advt, vehicle, station, date time
    'optSortby(2).Top = 300
    'optSortby(3).Top = 555
    
    'determine height of main (top) list box
    lbcVehAff.Height = (frcSelection.Height - chkListBox.Height - CkcAll.Height - 480) / 2

   
    chkListBox.Caption = "All Named Avails"
    chkListBox.Value = 0    '
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgAvailNamesInfo) - 1 Step 1
            lbcVehAff.AddItem Trim$(tgAvailNamesInfo(iLoop).sName)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgAvailNamesInfo(iLoop).iCode
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
    Screen.MousePointer = vbDefault
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True             '9-18-08 do not show Not Carried spots, default deselected
    'gPopSpotStatusCodes lbcStatus, ilHideNotCarried          'populate list box with hard-coded spot status codes
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried          '3-27-12 populate list box with new status codes (mg/bonus/replacements)

    'scan to see if any vef (vpf) are using avail names. Dont show the legend on input screen
    'if there are no vehicles using avail names
    lacStatusDesc.Visible = False
    For iLoop = LBound(tgVpfOptions) To UBound(tgVpfOptions) - 1
        If tgVpfOptions(iLoop).sAvailNameOnWeb = "Y" Then
            lacStatusDesc.Visible = True
            Exit For
        End If
    Next iLoop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmStationPgmClr
    Erase imUseCodes
    cprst.Close
    missed_rst.Close
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmPgmClrRpt = Nothing
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


'
'
'           Create AST for the Station Clearance report
'           Determine if running report by Station or Vehicle
'           <input> sStartDate - user entered start date
'           1-25-04
'
Private Sub mBuildAstPgmClr(sStartDate As String, iVef As Integer)
    Dim sSDate As String
    Dim iNoWeeks As Integer
    Dim sRptOption As String
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim iAdfCode As Integer
    Dim ilTemp As Integer
    Dim ilFoundStation As Integer
    'Dim ilIncludeCodes As Integer
    ''ReDim ilUseCodes(1 To 1) As Integer
    'ReDim ilUseCodes(0 To 0) As Integer
    
      'gObtainCodes lbcSelection(1), ilIncludeCodes, ilUseCodes()        'build array of which station codes to incl/excl
    
      iAdfCode = -1           'assume to retrieve all advertiser for vehicle or station option
      sSDate = gObtainPrevMonday(sStartDate)      'cp returned status is based on weeks
     
         'Get CPTT so that Stations requiring CP can be obtained
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
'        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
'        SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
'        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
'        SQLQuery = SQLQuery & " AND cpttVefCode = " & lbcSelection(0).ItemData(iVef)

        'SQLQuery = SQLQuery + " FROM shtt inner join cptt on shttcode = cpttshfcode inner join att on cpttatfcode = attcode "
        SQLQuery = SQLQuery & " from cptt inner join shtt on cpttshfcode = shttcode inner join att on cpttatfcode = attcode "
        SQLQuery = SQLQuery & " where ( cpttVefCode = " & lbcSelection(0).ItemData(iVef)
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
        Set cprst = gSQLSelectCall(SQLQuery)
        'D.S. 11/21/05
'        iRet = gGetMaxAstCode()
'        If Not iRet Then
'            Exit Sub
'        End If7
        
        While Not cprst.EOF
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = cprst!cpttCode
            tgCPPosting(0).iStatus = cprst!cpttStatus
            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
            tgCPPosting(0).lAttCode = cprst!cpttatfCode
            tgCPPosting(0).iAttTimeType = cprst!attTimeType
            tgCPPosting(0).iVefCode = cprst!cpttvefcode
            tgCPPosting(0).iShttCode = cprst!shttCode
            tgCPPosting(0).sZone = cprst!shttTimeZone
            tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
            tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
            
            ilFoundStation = False
            If imIncludeCodes Then
                'For ilTemp = 1 To UBound(imUseCodes) - 1 Step 1
                For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
                    If imUseCodes(ilTemp) = cprst!shttCode Then
                        ilFoundStation = True
                        Exit For
                    End If
                Next ilTemp
            Else
                ilFoundStation = True        '8/23/99 when more than half selected, selection fixed
                'For ilTemp = 1 To UBound(imUseCodes) - 1 Step 1
                For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
                    If imUseCodes(ilTemp) = cprst!shttCode Then
                        ilFoundStation = False
                        Exit For
                    End If
                Next ilTemp
            End If
            'Create AST records
            If ilFoundStation Then
                igTimes = 1 'By Week
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, True)   'False)  9-22-09 always create the ast records
            End If
            cprst.MoveNext
        Wend
        If (lbcSelection(1).ListCount = 0) Or (ckcAllStations.Value = vbChecked) Or (lbcSelection(1).ListCount = lbcSelection(1).SelCount) Then
            gClearASTInfo True
        Else
            gClearASTInfo False
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
'           mGetStatus - determine the status of a station spot (AST)
'           <output> ilStatus = status of spot (aired, not aired)
'                    ilCPStatus - Not reported or not aired
'                    ilAired - AST is aired spot
'                    ilNotAired - AST is not aired spot
'                    ilNotReported - AST is not reported spot
'                    ilLen = spot length (from  SDF)only if Minutes option; otherwise 1
Private Sub mGetStatus(ilStatus As Integer, ilCPStatus As Integer, ilAired As Integer, ilNotAired As Integer, ilNotReported As Integer, ilLen As Integer, llAstCode As Long)
Dim ilStatusOK As Integer
Dim SQLQuery As String

    ilAired = 0
    ilNotAired = 0
    ilNotReported = 0
    
    If ilCPStatus = 0 Then     'not reproted yet, if station is partially posted, it is still considered N/R
        ilNotReported = ilLen
    ElseIf ilCPStatus = 2 Then     'None aired (all spots run as pledged)
        ilNotAired = ilLen
    ElseIf ilStatus = 2 Or ilStatus = 3 Or ilStatus = 4 Or ilStatus = 5 Or ilStatus = 8 Or ilStatus = 14 Then       'not aired; 14 = missed-mg bypass
        'if resolved missed, needs to be ignored
        SQLQuery = "SELECT altLinkToAstCode, altastCode,  astcode, astStatus  From alt left outer JOIN ast ON altastcode = astcode "
        SQLQuery = SQLQuery & "Where altlinktoastcode = " & llAstCode & " or altastcode = " & llAstCode
        
        Set missed_rst = gSQLSelectCall(SQLQuery)         'read the associated ALT (associations) for the spot
        ilStatusOK = True
        While Not missed_rst.EOF
            'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
            If (missed_rst!altLinkToAstCode > 0) Then
                ilStatusOK = False
            Else            'reference should exist
                If missed_rst!altastcode <> llAstCode Then         'missed as a mg or replacement reference
                    ilStatusOK = False
                End If
            End If
            missed_rst.MoveNext
        Wend
        If ilStatusOK Then
            ilNotAired = ilLen
        End If
    Else
        ilAired = ilLen
    End If
End Sub
'
'               mTestFilterTimeAndStation - determine if the station has been selected
'               <input> list box of station
'               return - index to selected list box item, else -1
'               5-29-12 make anf filter a separate function; show invalid anf codes in a separate "Other" column
Private Function mTestFilterTimeandStation(ilShfCode As Integer, llFeedTime As Long, llStartTime As Long, llEndTime As Long)
Dim ilLoop As Integer
Dim ilInclude  As Integer

    ilInclude = -1
                          
    If llFeedTime >= llStartTime And llFeedTime <= llEndTime Then   'feed time must be within the requested times
        If ckcAllStations.Value = vbUnchecked Then              'if not all checked, see if selected

            For ilLoop = LBound(imUseCodes) To UBound(imUseCodes) - 1
                'If lbcSelection(1).Selected(ilLoop) Then
                If imUseCodes(ilLoop) = ilShfCode Then
                    ilInclude = 0
                    Exit For
                End If
                'End If
            Next ilLoop

        Else
            ilInclude = 0
        End If
    End If
'    If ilInclude = 0 Then
'        ilInclude = -1
'            For ilLoop = 0 To lbcVehAff.ListCount - 1
'                If lbcVehAff.Selected(ilLoop) Then
'                    If lbcVehAff.ItemData(ilLoop) = ilAnfcode Then
'                        ilInclude = ilLoop
'                        Exit For
'                    End If
'                End If
'            Next ilLoop
'
'    End If
    
    mTestFilterTimeandStation = ilInclude
End Function
'
'               Create SQL call to create the temporaray prepass record for Crystal to report from
'           <input> ilVef - vehicle code
'                   ilShttCode - station code
'                   ilHowManySets - # of pages required to print to complete set of avails for a station
'                   ilMaxColumns - max columns based on minutes or units option
'                   llTotalAired - total minutes (or  units) aired for station/vehicle
'                   llTotalNotRept - total minutes (or units) not reported for station/vehicle
'                   llTotalNotAired - total minutes (or  units) not reported for station/vehicle
Private Function mInsertIntoPCR(ilVef As Integer, ilShttCode As Integer, ilHowManySets As Integer, ilMaxColumns As Integer, llTotalAired As Long, llTotalNotRept As Long, llTotalNotAired As Long) As Integer
Dim ilLoopStats As Integer
Dim ilStatInx As Integer

    On Error GoTo ErrHand

    For ilLoopStats = 1 To ilHowManySets
            
           ilStatInx = (ilLoopStats - 1) * ilMaxColumns   'determine column 1 to 28 to place the stats in record for placement on page
   
           SQLQuery = "INSERT INTO " & "pcr "
           SQLQuery = SQLQuery & " (pcrvefCode,pcrshfcode, pcrSetNumber, pcrTotalAired, pcrTotalNotRept, pcrTotalNotAired, "
           
           '24 Named avail codes for columm headers (by units)
           '20 Named avail codes for column headers (by minute)
           SQLQuery = SQLQuery & "pcrAnfName1,pcrAnfName2,pcrAnfName3,pcrAnfName4,pcrAnfName5,pcrAnfName6,pcrAnfName7,pcrAnfName8,pcrAnfName9,pcrAnfName10,"
           SQLQuery = SQLQuery & "pcrAnfName11,pcrAnfName12,pcrAnfName13,pcrAnfName14,pcrAnfName15,pcrAnfName16,pcrAnfName17,pcrAnfName18,pcrAnfName19,pcrAnfName20,"
           
           If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "pcrAnfName21,pcrAnfName22,pcrAnfName23,pcrAnfName24,"
           End If
           
           '24 times for columm headers (by units)
           '20 times for column headers (by minute)
           SQLQuery = SQLQuery & "pcrFeedTime1,pcrFeedTime2,pcrFeedTime3,pcrFeedTime4,pcrFeedTime5,pcrFeedTime6,pcrFeedTime7,pcrFeedTime8,pcrFeedTime9,pcrFeedTime10,"
           SQLQuery = SQLQuery & "pcrFeedTime11,pcrFeedTime12,pcrFeedTime13,pcrFeedTime14,pcrFeedTime15,pcrFeedTime16,pcrFeedTime17,pcrFeedTime18,pcrFeedTime19,pcrFeedTime20,"
            
           If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "pcrFeedTime21,pcrFeedTime22,pcrFeedTime23,pcrFeedTime24,"
           End If
           
           '24 aired station stats for columm headers (by units)
           '20 aired station stats for column headers (by minute)F
           SQLQuery = SQLQuery & "pcrAired1,pcrAired2,pcrAired3,pcrAired4,pcrAired5,pcrAired6,pcrAired7,pcrAired8,pcrAired9,pcrAired10,"
           SQLQuery = SQLQuery & "pcrAired11,pcrAired12,pcrAired13,pcrAired14,pcrAired15,pcrAired16,pcrAired17,pcrAired18,pcrAired19,pcrAired20,"
           
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "pcrAired21,pcrAired22,pcrAired23,pcrAired24,"
            End If
            
           '24 not aired stats for columm headers (by units)
           '20 not aired stats for column headers (by minute)
           SQLQuery = SQLQuery & "pcrNotAired1,pcrNotAired2,pcrNotAired3,pcrNotAired4,pcrNotAired5,pcrNotAired6,pcrNotAired7,pcrNotAired8,pcrNotAired9,pcrNotAired10,"
           SQLQuery = SQLQuery & "pcrNotAired11,pcrNotAired12,pcrNotAired13,pcrNotAired14,pcrNotAired15,pcrNotAired16,pcrNotAired17,pcrNotAired18,pcrNotAired19,pcrNotAired20,"
            
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "pcrNotAired21,pcrNotAired22,pcrNotAired23,pcrNotAired24,"
            End If
            
           '24 not reported stats for columm headers (by units)
           '20 not reported stats for column headers (by minute)
           SQLQuery = SQLQuery & "pcrNotRept1,pcrNotRept2,pcrNotRept3,pcrNotRept4,pcrNotRept5,pcrNotRept6,pcrNotRept7,pcrNotRept8,pcrNotRept9,pcrNotRept10,"
           SQLQuery = SQLQuery & "pcrNotRept11,pcrNotRept12,pcrNotRept13,pcrNotRept14,pcrNotRept15,pcrNotRept16,pcrNotRept17,pcrNotRept18,pcrNotRept19,pcrNotRept20,"
           
           If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "pcrNotRept21,pcrNotRept22,pcrNotRept23,pcrNotRept24,"
           End If
           
           SQLQuery = SQLQuery & " pcrgendate, pcrGenTime) "
           
           'set the values of each field
            'set values for indices more than ilMaxColumns to 0 if ilMaxColumns is changed
           SQLQuery = SQLQuery & "Values ( "
           SQLQuery = SQLQuery & ilVef & ", "
           SQLQuery = SQLQuery & ilShttCode & ", "
           SQLQuery = SQLQuery & ilLoopStats & ", "
           SQLQuery = SQLQuery & llTotalAired & "," & llTotalNotRept & "," & llTotalNotAired & ","
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 1).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 2).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 3).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 4).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 5).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 6).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 7).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 8).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 9).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 10).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 11).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 12).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 13).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 14).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 15).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 16).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 17).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 18).sAnfName) & "', "
           SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 19).sAnfName) & "', "
           
           If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 20).sAnfName) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 21).sAnfName) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 22).sAnfName) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(tmStationPgmClr(ilStatInx + 23).sAnfName) & "', "
           End If
           
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 1).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 2).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 3).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 4).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 5).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 6).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 7).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 8).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 9).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 10).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 11).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 12).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 13).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 14).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 15).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 16).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 17).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 18).lFeedTime), sgShowTimeWOSecForm) & "', "
           SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 19).lFeedTime), sgShowTimeWOSecForm) & "', "
           
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 20).lFeedTime), sgShowTimeWOSecForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 21).lFeedTime), sgShowTimeWOSecForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 22).lFeedTime), sgShowTimeWOSecForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(gLongToTime(tmStationPgmClr(ilStatInx + 23).lFeedTime), sgShowTimeWOSecForm) & "', "
           End If
           
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 1).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 2).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 3).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 4).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 5).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 6).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 7).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 8).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 9).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 10).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 11).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 12).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 13).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 14).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 15).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 16).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 17).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 18).iAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 19).iAired & ", "
           
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 20).iAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 21).iAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 22).iAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 23).iAired & ", "
            End If
            
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 1).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 2).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 3).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 4).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 5).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 6).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 7).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 8).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 9).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 10).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 11).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 12).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 13).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 14).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 15).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 16).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 17).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 18).iNotAired & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 19).iNotAired & ", "
           
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 20).iNotAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 21).iNotAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 22).iNotAired & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 23).iNotAired & ", "
           End If
           
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 1).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 2).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 3).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 4).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 5).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 6).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 7).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 8).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 9).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 10).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 11).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 12).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 13).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 14).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 15).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 16).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 17).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 18).iNotRept & ", "
           SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 19).iNotRept & ", "
           
            If smMinUnits = "U" Then
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 20).iNotRept & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 21).iNotRept & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 22).iNotRept & ", "
                SQLQuery = SQLQuery & tmStationPgmClr(ilStatInx + 23).iNotRept & ", "
           End If
           
          
           SQLQuery = SQLQuery & " '" & Format$(smGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(smGenTime, False))))) & "')"
     
           cnn.BeginTrans
           'cnn.Execute SQLQuery, rdExecDirect
           If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "PgmClrRpt-mInsertIntoPCR"
                cnn.RollbackTrans
                mInsertIntoPCR = False
                Exit Function
           End If
           cnn.CommitTrans

       Next ilLoopStats
       mInsertIntoPCR = False
       Exit Function
       
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmPgmClrRpt" & "-mInsertIntoPCR"
    mInsertIntoPCR = True
    Exit Function
End Function
'
'               mTestFilterANFCode - determine if the Avail name has been selected
'               <input> list of anf codes
'               return - index to selected list box item, else -1
Private Function mTestFilterANFCode(ilAnfcode As Integer, ilShowOther As Integer)
Dim ilLoop As Integer
Dim ilInclude  As Integer
Dim ilValidAnf As Integer
Dim ilFound As Integer
                          
'    If llFeedTime >= llStartTime And llFeedTime <= llEndTime Then   'feed time must be within the requested times
'        If ckcAllStations.Value = vbUnchecked Then              'if not all checked, see if selected
'            For ilLoop = LBound(imUseCodes) To UBound(imUseCodes) - 1
'                'If lbcSelection(1).Selected(ilLoop) Then
'                If imUseCodes(ilLoop) = ilShfCode Then
'                    ilInclude = 0
'                    Exit For
'                End If
'                'End If
'            Next ilLoop
'
'        Else
'            ilInclude = 0
'        End If
'    End If
'    If ilInclude = 0 Then
        ilInclude = -1
        ilShowOther = False
        If ilAnfcode = 0 Then         'no named avail defined, put in Other colunm
            ilShowOther = True
            mTestFilterANFCode = ilInclude
            Exit Function
        End If
        ilFound = False
        'If chkListBox.Value = vbUnchecked Then              'if not all checked, see if selected
            For ilLoop = 0 To lbcVehAff.ListCount - 1
                If lbcVehAff.Selected(ilLoop) Then
                    If lbcVehAff.ItemData(ilLoop) = ilAnfcode Then
                        ilInclude = ilLoop
                        ilFound = True
                        Exit For
                    End If
                Else
                    'not selected, but see if its a valid one, but excluded
                    If lbcVehAff.ItemData(ilLoop) = ilAnfcode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilLoop
            If ilFound = False Then      'didnt find a valid anf code that was either selectedor not selected, so it needs to be shown in Other column
                ilShowOther = True
            End If
    'End If
    
    mTestFilterANFCode = ilInclude
    Exit Function
End Function
