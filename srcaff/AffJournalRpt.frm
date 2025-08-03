VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmJournalRpt 
   Caption         =   "Export Journal Report"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   375
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6480
      FormDesignWidth =   9360
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
      Height          =   4600
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8280
         Picture         =   "AffJournalRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Select Stations from File.."
         Top             =   285
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalLogFrom 
         Height          =   285
         Left            =   1305
         TabIndex        =   13
         Top             =   600
         Width           =   915
         _extentx        =   1984
         _extenty        =   582
         borderstyle     =   1
         csi_showdropdownonfocus=   -1
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffJournalRpt.frx":056A
         csi_daynamefont =   "AffJournalRpt.frx":0596
         csi_monthnamefont=   "AffJournalRpt.frx":05C4
      End
      Begin VB.CheckBox ckcDiscrepOnly 
         Caption         =   "Discrepancies Only"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox ckcStationInfo 
         Caption         =   "Include Stations"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.ListBox lbcStations 
         Height          =   3570
         ItemData        =   "AffJournalRpt.frx":05F2
         Left            =   6735
         List            =   "AffJournalRpt.frx":05F9
         MultiSelect     =   2  'Extended
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   660
         Width           =   1980
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   6735
         TabIndex        =   24
         Top             =   285
         Width           =   1455
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3570
         ItemData        =   "AffJournalRpt.frx":0600
         Left            =   4485
         List            =   "AffJournalRpt.frx":0602
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   660
         Width           =   1995
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4485
         TabIndex        =   22
         Top             =   285
         Width           =   1935
      End
      Begin V81Affiliate.CSI_Calendar CalActivityStart 
         Height          =   285
         Left            =   1305
         TabIndex        =   16
         Top             =   960
         Width           =   915
         _extentx        =   1614
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffJournalRpt.frx":0604
         csi_daynamefont =   "AffJournalRpt.frx":0630
         csi_monthnamefont=   "AffJournalRpt.frx":065E
      End
      Begin V81Affiliate.CSI_Calendar CalActivityEnd 
         Height          =   285
         Left            =   3045
         TabIndex        =   17
         Top             =   960
         Width           =   915
         _extentx        =   1614
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   0
         csi_defaultdatetype=   0
         csi_caldateformat=   1
         font            =   "AffJournalRpt.frx":068C
         csi_daynamefont =   "AffJournalRpt.frx":06B8
         csi_monthnamefont=   "AffJournalRpt.frx":06E6
      End
      Begin VB.Frame frcDates 
         Caption         =   "Dates"
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4230
         Begin VB.TextBox txtLogTo 
            Height          =   285
            Left            =   3225
            TabIndex        =   14
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Activity- From"
            Height          =   225
            Left            =   120
            TabIndex        =   18
            Top             =   615
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "To"
            Height          =   225
            Left            =   2520
            TabIndex        =   19
            Top             =   615
            Width           =   435
         End
         Begin VB.Label lacFrom 
            Caption         =   "Log Date"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   255
            Width           =   840
         End
         Begin VB.Label lacTo 
            Caption         =   "# Days"
            Height          =   225
            Left            =   2520
            TabIndex        =   15
            Top             =   255
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   5115
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4920
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
      Left            =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   4
         Left            =   2130
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffJournalRpt.frx":0714
         Left            =   1335
         List            =   "AffJournalRpt.frx":0716
         TabIndex        =   4
         Top             =   765
         Width           =   2040
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Mail List"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1125
         Width           =   1335
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmJournalRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmJournalRpt - Dump the log that contains the status of the affiliate to
'*                  web export
'*
'*
'*  Copyright Counterpoint Software, Inc.
'

'****************************************************************************
Option Explicit

Private hmMail As Integer
Private smToFile As String
Private imChkStationIgnore As Integer
Private imChkListBoxIgnore As Integer

Private Sub CalActivityEnd_GotFocus()
    gCtrlGotFocus CalActivityEnd
    CalActivityEnd.ZOrder (vbBringToFront)
End Sub

Private Sub CalActivityStart_GotFocus()
    gCtrlGotFocus CalActivityStart
    CalActivityStart.ZOrder (vbBringToFront)
End Sub

Private Sub CalLogFrom_GotFocus()
    gCtrlGotFocus CalLogFrom
    CalLogFrom.ZOrder (vbBringToFront)
End Sub

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    Dim iLoop As Integer
    
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
 
    ckcAllStations.Value = vbUnchecked
    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
            lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
        End If
    Next iLoop
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllStations_Click()
 Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStationIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStationIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmJournalRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim iRet As Integer
    Dim sVehicles As String
    Dim sStations As String
    Dim sStationType As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sMail As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim slActivityBetweenStart As String    'sql query
    Dim slActivityBetweenEnd As String      'sql query
    Dim slActivityBetween
    
    Dim slLogBetweenStart As String    'sql query
    Dim slLogBetweenEnd As String      'sql query
    Dim slLogBetween
    Dim slDays As String

    Dim slDescription As String
    Dim slDescrepts As String
        
        On Error GoTo ErrHand
        
        'Check the activity dates entered
        slActivityBetweenStart = Trim$(CalActivityStart.Text)
        If slActivityBetweenStart = "" Then
            slActivityBetweenStart = "1/1/1970"
        End If
        
        slActivityBetweenEnd = Trim$(CalActivityEnd.Text)
        If slActivityBetweenEnd = "" Then
            slActivityBetweenEnd = "12/31/2069"
        End If
        
        'verify valid of activity start & end dates
        If gIsDate(slActivityBetweenEnd) = False Then
            Beep
            gMsgBox "Please enter a valid activity date (m/d/yy)", vbCritical
            CalActivityStart.SetFocus
            Exit Sub
        End If
        If gIsDate(slActivityBetweenEnd) = False Then
            Beep
            gMsgBox "Please enter a valid activity date (m/d/yy)", vbCritical
            CalActivityEnd.SetFocus
            Exit Sub
        End If
        
        'Check the log dates entered (log start date + # days)
        slLogBetweenStart = Trim$(CalLogFrom.Text)
        
        If gIsDate(slLogBetweenStart) = False And slLogBetweenStart <> "" Then
            Beep
            gMsgBox "Please enter a valid log date (m/d/yy)", vbCritical
            CalLogFrom.SetFocus
            Exit Sub
        End If
        
        If slLogBetweenStart = "" Then
            slLogBetweenStart = "1/1/1970"
        End If
        slDays = Trim$(txtLogTo.Text)
        If Val(slDays) > 365 Then
            gMsgBox "Max 1 year allowed"
            txtLogTo.SetFocus
            Exit Sub
        End If
        slLogBetweenEnd = Format$(DateAdd("d", Val(slDays) - 1, slLogBetweenStart), "mm/dd/yy")
    
        If slDays = "" Then                         'no # days entered, assume entire file
            slLogBetweenEnd = "12/31/2069"
        End If
        
        'at least 1 vehicle and station required
        If lbcVehAff.SelCount = 0 Or lbcStations.SelCount = 0 Then
            gMsgBox "AT least one vehicle and station must be selected"
            Exit Sub
        End If
        
        
        Screen.MousePointer = vbHourglass
      
        If optRptDest(0).Value = True Then
            'CRpt1.Destination = crptToWindow
            ilRptDest = 0
        ElseIf optRptDest(1).Value = True Then
            'CRpt1.Destination = crptToPrinter
            ilRptDest = 1
        ElseIf optRptDest(2).Value = True Then
            'gOutputMethod frmJournalRpt, "ExpJournal.rpt", sOutput
            ilRptDest = 2
            'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
            ilExportType = cboFileType.ListIndex    '3-15-04
        ElseIf optRptDest(3).Value = True Then
            iRet = OpenMsgFile(hmMail, smToFile)
            If iRet = False Then
                Exit Sub
            End If
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False

        gUserActivityLog "S", sgReportListName & ": Prepass"
        
        slActivityBetweenStart = Format(slActivityBetweenStart, "m/d/yyyy")     'insure year appended
        slActivityBetweenEnd = Format(slActivityBetweenEnd, "m/d/yyyy")         'insure year appended
        slLogBetweenStart = Format(slLogBetweenStart, "m/d/yyyy")     'insure year appended
        slLogBetweenEnd = Format(slLogBetweenEnd, "m/d/yyyy")         'insure year appended
       
        'create sql query to get export info starting between activity dates enteredestStartDate <=" & "'" + Format$(slActivityBetweenStart2, sgSQLDateForm) & "'"
        slActivityBetween = " (esfEndDate >=" & "'" & Format$(slActivityBetweenStart, sgSQLDateForm) & "'" & " And esfStartDate <=" & "'" + Format$(slActivityBetweenEnd, sgSQLDateForm) & "')"
        'create sql query to get export info for logs exported  between 2 spans
        slLogBetween = " ((esfExpDate + esfNumDays-1) >=" & "'" & Format$(slLogBetweenStart, sgSQLDateForm) & "'" & " And esfExpDate <=" & "'" + Format$(slLogBetweenEnd, sgSQLDateForm) & "')"
        
        sVehicles = ""
        sStations = ""
           
        If chkListBox.Value = 0 Then    '= 0 Then
            'User did NOT select all vehicles
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
    
        'User did NOT select all stations
        If ckcAllStations.Value = 0 Then    '= 0 Then
            'User did NOT select all stations
            For i = 0 To lbcStations.ListCount - 1 Step 1
                If lbcStations.Selected(i) Then
                    If Len(sStations) = 0 Then
                        sStations = "(shttCode = " & lbcStations.ItemData(i) & ")"
                    Else
                        sStations = sStations & " OR (shttCode = " & lbcStations.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
            
        SQLQuery = "SELECT * from  (((edf_export_detail left Outer Join esf_export_summary on edfesfcode = esfcode)"
        SQLQuery = SQLQuery + " left Outer Join vef_vehicles on edfvefcode = vefcode)"
        SQLQuery = SQLQuery + " left Outer Join shtt on edfshttcode = shttcode)"
        
        SQLQuery = SQLQuery + " where "
        SQLQuery = SQLQuery + slActivityBetween + " and " + slLogBetween
        
        If sVehicles <> "" Then
            SQLQuery = SQLQuery + " and (" & sVehicles & ")"
        End If
        If sStations <> "" Then
            SQLQuery = SQLQuery + " and (" & sStations & ")"
        End If

        
        'Setup description for activity dates entered
        If slActivityBetweenStart = "1/1/1970" And slActivityBetweenEnd = "12/31/2069" Then
            sgCrystlFormula1 = "All Activity Dates"
        ElseIf slActivityBetweenStart = "1/1/1970" Then
            sgCrystlFormula1 = "Activity Dates thru " + slActivityBetweenEnd
        ElseIf slActivityBetweenEnd = "12/31/2069" Then
            sgCrystlFormula1 = "Activity Dates from " + slActivityBetweenStart
        Else
            sgCrystlFormula1 = "Activity Dates " + slActivityBetweenStart + "-" + slActivityBetweenEnd
        End If
        
        'setup description for Log dates
        If slLogBetweenStart = "1/1/1970" And slLogBetweenEnd = "12/31/2069" Then
            sgCrystlFormula2 = "All Log Dates"
        ElseIf slLogBetweenStart = "1/1/1970" Then
            sgCrystlFormula2 = "Log Dates thru " + slLogBetweenEnd
        ElseIf slLogBetweenEnd = "12/31/2069" Then
            sgCrystlFormula2 = "Log Dates from " + slLogBetweenStart
        Else
            sgCrystlFormula2 = "Log Dates " + slLogBetweenStart + "-" + slLogBetweenEnd
        End If

         
        If ckcStationInfo.Value = vbChecked Then        'show station info
            sgCrystlFormula3 = "'Y'"
        Else
            sgCrystlFormula3 = "'N'"
        End If
        
        If ckcDiscrepOnly.Value = vbChecked Then        'Discrepancies only
            sgCrystlFormula4 = "'Y'"
            slDescrepts = "  and (esfErrors <> " & "'" & "'" & " or edfAlert <> " & "'" & "')"
            'slDescrepts = ""        'take everything, filter out in .rpt
        Else
            sgCrystlFormula4 = "'N'"
            slDescrepts = ""
        End If
        
        sgCrystlFormula5 = "'V'"          'force sort by vehicle, currently only 1 sort option
    
        SQLQuery = SQLQuery + slDescrepts
        
        slRptName = "AfJournal.rpt"
        slExportName = "JournalRpt"
        gUserActivityLog "E", sgReportListName & ": Prepass"
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
        SQLQuery = ""
    
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault
    'If optRptDest(2).Value = True Then
    '    gMsgBox "Output Sent To: " & sOutput, vbInformation
    'End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "affJournalRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmJournalRpt
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
    gSelectiveStationsFromImport lbcStations, ckcAllStations, Trim$(CommonDialog1.fileName)
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
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmJournalRpt
    gCenterForm frmJournalRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmJournalRpt.Caption = "Export Journal Report - " & sgClientName
    imChkListBoxIgnore = False
    
    slDate = Format$(gNow(), "m/d/yyyy")
    'Do While Weekday(slDate, vbSunday) <> vbMonday
    '    slDate = DateAdd("d", -1, slDate)
    'Loop
    'CalActivityStart.Text = slDate
    'CalActivityEnd.Text = DateAdd("d", 6, slDate)
    
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    chkListBox.Value = 0    'chged from false to 0 10-22-99
        

    gPopExportTypes cboFileType     '3-15-04
    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmJournalRpt = Nothing
End Sub

Private Sub grdVehAff_Click()
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub lbcStations_Click()
If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        imChkStationIgnore = True
        'chkListBox.Value = False
        ckcAllStations.Value = 0    'chged from false to 0 10-22-99
        imChkStationIgnore = False
    End If
End Sub

Private Sub lbcVehAff_Click()
Dim iLoop As Integer
Dim llStaCode As Long
Dim ilVefCode As Integer
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
    'if # of vehicles selected is more than 1, show all vehicles; otherwise
    'just show the stations that belong to the vehicle
    ckcAllStations.Value = vbUnchecked
    If lbcVehAff.SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    Else            'show only those that belong to the vehicle
        ilVefCode = 0
        For iLoop = 0 To lbcVehAff.ListCount
            If lbcVehAff.Selected(iLoop) Then
                ilVefCode = lbcVehAff.ItemData(iLoop)
                Exit For
            End If
        Next iLoop
        lbcStations.Clear
'        SQLQuery = "SELECT Distinct shttCallLetters, shttCode"
'        SQLQuery = SQLQuery + " FROM shtt , att"
'        SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode"
'        SQLQuery = SQLQuery + " AND attVefCode = " & ilVefCode & ")"
'        SQLQuery = SQLQuery + " ORDER BY shttCallLetters"

        '4-23-18 retrieve the market to show next to station in list box
        SQLQuery = "SELECT Distinct shttCallLetters, shttCode, mktname "
        SQLQuery = SQLQuery + " FROM shtt inner join att on shttcode = attshfcode inner join mkt on shttmktcode = mktcode "
        SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode"
        SQLQuery = SQLQuery + " AND attVefCode = " & ilVefCode & ")"
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters"

        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF

            llStaCode = gBinarySearchStation(rst!shttCallLetters)
            If llStaCode <> -1 Then
                lbcStations.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(rst!mktName)
                lbcStations.ItemData(lbcStations.NewIndex) = rst!shttCode
            End If
            rst.MoveNext
        Wend
    End If

End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub optVehAff_Click(Index As Integer)
    Dim iLoop As Integer
    Dim iIndex As Integer
End Sub
