VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogActivityRpt 
   Caption         =   "Web Log Activity Report"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "AffLogActivityRpt.frx":0000
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
      Height          =   3180
      Left            =   240
      TabIndex        =   6
      Top             =   1845
      Width           =   6960
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6240
         Picture         =   "AffLogActivityRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Select Stations from File.."
         Top             =   285
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1710
         TabIndex        =   8
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
         font            =   "AffLogActivityRpt.frx":0E34
         csi_daynamefont =   "AffLogActivityRpt.frx":0E60
         csi_monthnamefont=   "AffLogActivityRpt.frx":0E8E
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   1710
         TabIndex        =   10
         Top             =   600
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
         font            =   "AffLogActivityRpt.frx":0EBC
         csi_daynamefont =   "AffLogActivityRpt.frx":0EE8
         csi_monthnamefont=   "AffLogActivityRpt.frx":0F16
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sort by"
         Height          =   900
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Width           =   2835
         Begin VB.OptionButton optSortby 
            Caption         =   "Vehicle, Station, Print Date"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   13
            Top             =   555
            Width           =   2565
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Station, Print Date"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   300
            Value           =   -1  'True
            Width           =   2280
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   2400
         ItemData        =   "AffLogActivityRpt.frx":0F44
         Left            =   3075
         List            =   "AffLogActivityRpt.frx":0F46
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   3555
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3060
         TabIndex        =   14
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Earliest Log Date:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Latest Log Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   615
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   17
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   16
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
         ItemData        =   "AffLogActivityRpt.frx":0F48
         Left            =   1050
         List            =   "AffLogActivityRpt.frx":0F4A
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   135
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
         Left            =   135
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
Attribute VB_Name = "frmLogActivityRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'*  frmLogActivityRpt -
'*
'*  Created 10/7/03 D Smith
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
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

Private Sub cmdDone_Click()
    Unload frmLogActivityRpt
End Sub
Private Sub cmdReport_Click()
    Dim i As Integer
    Dim sVehicles As String
    Dim sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim dFWeek As Date
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    
    On Error GoTo ErrHand
    
    sStartDate = Trim$(CalOnAirDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If
    sEndDate = Trim$(CalOffAirDate.Text)
    If sEndDate = "" Then
        sEndDate = "12/31/2069"
    End If
    If gIsDate(sStartDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex    '3-15-04 get user export type selected
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sVehicles = ""
    sStations = ""
     
    ' Detrmine what to sort by
    If optSortby(1).Value = True Then     'vehicles
        If chkListBox.Value = 0 Then      'User did NOT select all vehicles
            For i = 0 To lbcVehAff.ListCount - 1 Step 1
                If lbcVehAff.Selected(i) Then
                    If Len(sVehicles) = 0 Then
                        sVehicles = "(vefCode = " & lbcVehAff.ItemData(i) & ")"
                    Else
                        sVehicles = sVehicles & " OR (vefCode = " & lbcVehAff.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
    Else                                          'station
        If chkListBox.Value = 0 Then              'User did NOT select all stations
            For i = 0 To lbcVehAff.ListCount - 1 Step 1
                If lbcVehAff.Selected(i) Then
                    If Len(sStations) = 0 Then
                        sStations = "(shttCode = " & lbcVehAff.ItemData(i) & ")"
                    Else
                        sStations = sStations & " OR (shttCode = " & lbcVehAff.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
    End If
    
    If optSortby(0).Value = True Then          'station
        sgCrystlFormula1 = "'S'"
    Else                                'vehicle
        sgCrystlFormula1 = "'V'"
    End If
    
    dFWeek = CDate(sStartDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dFWeek = CDate(sEndDate)
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
  
    SQLQuery = "SELECT weblCallLetters, weblVehicleName, weblUserName, WeblPostDay, weblDate "
    'SQLQuery = SQLQuery & "FROM WebL, Shtt, Att, VEF_Vehicles "
    '1-16-13 use the current call letters / vehicle name vs what is in webl
    SQLQuery = SQLQuery & "FROM WebL inner join att on weblattcode = attcode inner join shtt on attshfcode = shttcode inner join vef_vehicles on attvefcode = vefcode "
    SQLQuery = SQLQuery & "WHERE (WebLPostDay >= '" & Format$(sStartDate, sgSQLDateForm) & "' AND WebLPostDay <= '" & Format$(sEndDate, sgSQLDateForm) & "') "
    SQLQuery = SQLQuery & "And (WeblType = 2) "
    'SQLQuery = SQLQuery & "And WeblAttCode = AttCode "
    'SQLQuery = SQLQuery & "And WeblCallLetters = ShttCallLetters "
    'SQLQuery = SQLQuery & "and WeblVehicleName = VefName "
    
    If optSortby(0).Value = True Then         'sort by  station
        'Prepare records to pass to Crystal
        If sStations <> "" Then
            SQLQuery = SQLQuery + " AND (" & sStations & ")"
        End If
    Else                                       'sort by  vehicle
        If sVehicles <> "" Then
            SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
        End If
    End If
      
    gUserActivityLog "E", sgReportListName & ": Prepass"
    If optSortby(0).Value = True Then          'station
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfLogStaAct.rpt", "AfLogStaAct"
    Else                                'vehicle
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfLogVefAct.rpt", "AfLogVefAct"
    End If
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmLogActivityRpt" & "-cmdReport_Click"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmLogActivityRpt
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
    gSelectiveStationsFromImport lbcVehAff, chkListBox, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmLogActivityRpt
    gCenterForm frmLogActivityRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    
    imChkListBoxIgnore = False
    frmLogActivityRpt.Caption = "Web Log Activity Report - " & sgClientName
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    'default to show the stations (vs vehicles)
    chkListBox.Caption = "All Stations"
    chkListBox.Value = 0
    lbcVehAff.Clear
    
    Call LoadStationNames
    
    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = False
End Sub

Private Sub LoadStationNames()
    Dim StationRST As ADODB.Recordset
    
    ' JD 04/27/05 - Load only station names that are web enabled.
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode FROM shtt, att "
    SQLQuery = SQLQuery + "Where shtt.shttCode = att.attShfCode"
    SQLQuery = SQLQuery + " And attExportType = 1"  ' Only Web Enabled Stations
    Set StationRST = gSQLSelectCall(SQLQuery)
    lbcVehAff.Clear
    While Not StationRST.EOF
        lbcVehAff.AddItem (StationRST!shttCallLetters)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = StationRST!shttCode
        StationRST.MoveNext
    Wend
End Sub

Private Sub LoadVehicleNames()
    Dim VehicleRST As ADODB.Recordset
    
    ' JD 04/27/05 - Load only vehicle names that are web enabled.
    SQLQuery = "SELECT DISTINCT vefName, vefCode FROM shtt, att, VEF_Vehicles"
    SQLQuery = SQLQuery + " Where shttCode = attShfCode"
    SQLQuery = SQLQuery + " And vefCode = attvefCode"
    SQLQuery = SQLQuery + " And attExportType = 1"
    Set VehicleRST = gSQLSelectCall(SQLQuery)
    lbcVehAff.Clear
    While Not VehicleRST.EOF
        lbcVehAff.AddItem (VehicleRST!vefName)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = VehicleRST!vefCode
        VehicleRST.MoveNext
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmLogActivityRpt = Nothing
End Sub

'Private Sub grdVehAff_Click()
'    If chkListBox.Value = 1 Then
'        imChkListBoxIgnore = True
'        'chkListBox.Value = False
'        chkListBox.Value = 0    'chged from false to 0 10-22-99
'        imChkListBoxIgnore = False
'    End If
'End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        chkListBox.Value = 0
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 dfault to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub optSortby_Click(Index As Integer)
Dim iLoop As Integer
Dim iIndex As Integer
    
    Screen.MousePointer = vbHourglass
    If optSortby(1).Value = True Then
        cmdStationListFile.Visible = False
        chkListBox.Caption = "All Vehicles"
        chkListBox.Value = 0
        lbcVehAff.Clear
        Call LoadVehicleNames
    Else
        cmdStationListFile.Visible = True
        chkListBox.Caption = "All Stations"
        chkListBox.Value = 0
        lbcVehAff.Clear
        Call LoadStationNames
     End If
    Screen.MousePointer = vbDefault
End Sub
