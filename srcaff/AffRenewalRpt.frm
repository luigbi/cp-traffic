VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRenewalRpt 
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
      Top             =   960
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
      TabIndex        =   9
      Top             =   1680
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8280
         Picture         =   "AffRenewalRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Select Stations from File.."
         Top             =   285
         Width           =   360
      End
      Begin VB.ListBox lbcStations 
         Height          =   3570
         ItemData        =   "AffRenewalRpt.frx":056A
         Left            =   6735
         List            =   "AffRenewalRpt.frx":056C
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   660
         Width           =   1980
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   6735
         TabIndex        =   20
         Top             =   285
         Width           =   1455
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3570
         ItemData        =   "AffRenewalRpt.frx":056E
         Left            =   4485
         List            =   "AffRenewalRpt.frx":0570
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   660
         Width           =   1995
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sort by"
         Height          =   675
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   3105
         Begin VB.OptionButton optVehAff 
            Caption         =   "End Date"
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   1100
         End
         Begin VB.OptionButton optVehAff 
            Caption         =   "Stations"
            Height          =   255
            Index           =   1
            Left            =   1245
            TabIndex        =   16
            Top             =   270
            Width           =   1100
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4485
         TabIndex        =   17
         Top             =   285
         Width           =   1320
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   270
         Left            =   2985
         TabIndex        =   11
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalEnterFrom 
         Height          =   270
         Left            =   1290
         TabIndex        =   12
         Top             =   1860
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   2
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   1515
         TabIndex        =   10
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalEnterTo 
         Height          =   270
         Left            =   2790
         TabIndex        =   13
         Top             =   1830
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
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
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   2
      End
      Begin VB.Frame frcDates 
         Caption         =   "Agreements"
         Height          =   690
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4230
         Begin VB.Label Label4 
            Caption         =   "To"
            Height          =   225
            Left            =   2490
            TabIndex        =   23
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label3 
            Caption         =   "Renewals - From"
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Label lacFrom 
         Caption         =   "Entered- From"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   25
         Top             =   1905
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lacTo 
         Caption         =   "To"
         Height          =   225
         Left            =   2340
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5355
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   5115
      TabIndex        =   7
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1140
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         TabIndex        =   4
         Top             =   765
         Width           =   2040
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
Attribute VB_Name = "frmRenewalRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmRenewalRpt - Report of agreements that are due for renewal based on
'   user entered date spans against the agreement end date
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'

'****************************************************************************
Option Explicit

Private imChkStationIgnore As Integer
Private imChkListBoxIgnore As Integer
Private rst_att As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset
Dim imRptIndex As Integer                   'report option

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
    
    If lbcVehAff.SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
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
    Unload frmRenewalRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim iRet As Integer
    Dim sCode As String
    Dim sName As String
    Dim sVehicles As String
    Dim sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim slEnterFrom As String
    Dim slEnterTo As String
    Dim slEnteredRange As String
    Dim slGenDate As String
    Dim slGenTime As String
        
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
            ilRptDest = 2
            ilExportType = cboFileType.ListIndex    '3-15-04
        End If
            
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False

        gUserActivityLog "S", sgReportListName & ": Prepass"
        
        sStartDate = Format(sStartDate, "m/d/yyyy")
        sEndDate = Format(sEndDate, "m/d/yyyy")
       
        'Test Agreement end date falls between the user entered date spans
        'only use the end date for filtering
        'sDateRange = " (attDropDate >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attDropdate <= " & "'" + Format$(sEndDate, sgSQLDateForm) & "') or "
        'sDateRange = sDateRange & " (attAgreeEnd >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attAgreeEnd <= " & "'" + Format$(sEndDate, sgSQLDateForm) & "') "
        sDateRange = " (attAgreeEnd >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attAgreeEnd <= " & "'" + Format$(sEndDate, sgSQLDateForm) & "') "

        sVehicles = ""
        sStations = ""
        
        If chkListBox.Value = vbUnchecked Then
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
        If ckcAllStations.Value = vbUnchecked Then    '= 0 Then
            'User did NOT select all stations
            For i = 0 To lbcStations.ListCount - 1 Step 1
                If lbcStations.Selected(i) Then
                    If Len(sStations) = 0 Then
                        sStations = "(attShfCode = " & lbcStations.ItemData(i) & ")"
                    Else
                        sStations = sStations & " OR (attShfCode = " & lbcStations.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
        'End If
        
       
'        If optVehAff(0).Value = True Then   'VEHICLE SORT
            
            SQLQuery = "SELECT * From shtt INNER JOIN att att ON shttCode = att.attShfCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
            SQLQuery = SQLQuery + " INNER JOIN   VEF_Vehicles ON attVefCode = vefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ust ON shttMktRepUstCode = ust.ustCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustserv ON shttServRepUstCode = ustserv.ustCode "
    
            SQLQuery = SQLQuery + " LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events ON ust.ustEMailCefCode = CEF_Comments_Events.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustatt ON attMktRepUstCode = ustatt.ustCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN ust ustattserv ON attServRepUstCode = ustattserv.ustCode "
    
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_Att ON ustatt.ustEMailCefCode = CEF_Comments_Events_Att.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_AttServ ON ustattserv.ustEMailCefCode = CEF_Comments_Events_AttServ.cefCode "
            SQLQuery = SQLQuery + "LEFT OUTER JOIN CEF_Comments_Events CEF_Comments_Events_Serv ON ustserv.ustEMailCefCode = CEF_Comments_Events_Serv.cefCode "

            SQLQuery = SQLQuery + " Where (" & sDateRange & ")"
            
            SQLQuery = SQLQuery + " AND ( shttType = 0 ) "
            If sVehicles <> "" Then
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
            If sStations <> "" Then '12-13-00
                SQLQuery = SQLQuery + " AND (" & sStations & ")"
            End If
            
            If optVehAff(0).Value = True Then               'sort by end date
                SQLQuery = SQLQuery + " ORDER BY attAgreeEnd, shttCallLetters, vefName"
                sgCrystlFormula1 = "'D'"
            Else
                SQLQuery = SQLQuery + " ORDER BY shttCallLetters, attAgreeEnd, vefName"
                sgCrystlFormula1 = "'S'"
            End If
    
'        Else            'STATION SORT
'
'            SQLQuery = "SELECT * From shtt INNER JOIN att ON shttCode = attShfCode "
'            SQLQuery = SQLQuery & "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
'            SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles ON attVefCode = vefCode "
'
'            SQLQuery = SQLQuery + " Where (" & sDateRange & ")"
'            SQLQuery = SQLQuery + " AND ( shttType = 0 ) "
'
'            If sStations <> "" Then
'                SQLQuery = SQLQuery + " AND (" & sStations & ")"
'            End If
'            If sVehicles <> "" Then     '12-13-00
'                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
'            End If
'        End If
        slGenDate = Format$(gNow(), "m/d/yyyy")
        slGenTime = Format$(gNow(), sgShowTimeWSecForm)
        gUserActivityLog "E", sgReportListName & ": Prepass"

        Set rst_att = gSQLSelectCall(SQLQuery)
          
        
        dFWeek = CDate(sStartDate)
        sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")" 'StartDate
        dFWeek = CDate(sEndDate)
        sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")" 'EndDate
    
        slRptName = "AfRenewal.rpt"
        slExportName = "Renewal"
        
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
      
        On Error Resume Next
        rst_att.Close
        rst_Shtt.Close
        
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True

        Screen.MousePointer = vbDefault
        Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRenewalRpt-cmdReport"
    Exit Sub
    
ErrHandUpdate:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRenewalRpt-cmdReport"
    Exit Sub
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmRenewalRpt
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
    gSetFonts frmRenewalRpt
    gCenterForm frmRenewalRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    imRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    
    frmRenewalRpt.Caption = "Affiliate Agreement Renewal Status Report - " & sgClientName
    
    imChkListBoxIgnore = False
    
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = slDate
    CalOffAirDate.Text = DateAdd("d", 6, slDate)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    
    CalEnterFrom.ZOrder (0)
    CalOnAirDate.ZOrder (0)
    CalEnterTo.ZOrder (0)
    CalOffAirDate.ZOrder (0)

    lbcStations.Clear
    
    'dont show all stations unless more than 1 vehicle is selected; otherwise, only show those stations that have an agreement with the vehicle
'    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
'            lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
'            lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
'        End If
'    Next iLoop
'
    ilRet = gPopVtf()               'obtain the vehicle text info

    gPopExportTypes cboFileType     '3-15-04
    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_att.Close
    rst_Shtt.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmRenewalRpt = Nothing
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
Dim ilVefCode As Integer
Dim llShfCode As Long
Dim ilRet As Integer

    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
    If lbcVehAff.SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    Else
        'only one, show only the ones with agreements
        lbcStations.Clear
        ilVefCode = lbcVehAff.ItemData(lbcVehAff.ListIndex)
        SQLQuery = "Select distinct attshfcode from att where attvefcode = " & ilVefCode
        Set rst_Shtt = gSQLSelectCall(SQLQuery)
        While Not rst_Shtt.EOF
            llShfCode = gBinarySearchStationInfoByCode(rst_Shtt!attshfcode)
            If llShfCode <> -1 Then
                lbcStations.AddItem Trim$(tgStationInfoByCode(llShfCode).sCallLetters) & ", " & Trim$(tgStationInfoByCode(llShfCode).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfoByCode(llShfCode).iCode
            End If

        rst_Shtt.MoveNext
        Wend
    End If
    ckcAllStations.Value = vbUnchecked
    

End Sub


Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

