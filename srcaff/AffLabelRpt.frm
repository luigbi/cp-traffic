VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmLabelRpt 
   Caption         =   "Mailing Labels"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "AffLabelRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   6975
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3405
      Top             =   1095
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6450
      FormDesignWidth =   6975
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
      Height          =   4545
      Left            =   240
      TabIndex        =   6
      Top             =   1725
      Width           =   6495
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   900
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
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
         Left            =   1080
         TabIndex        =   13
         Top             =   1320
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
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
      Begin VB.ListBox lbcContacts 
         Height          =   1425
         ItemData        =   "AffLabelRpt.frx":08CA
         Left            =   3120
         List            =   "AffLabelRpt.frx":08CC
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Frame Frame5 
         Caption         =   "Number of Columns"
         Height          =   795
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   2355
         Begin VB.OptionButton OptCols 
            Caption         =   "3 (1"" x 2 5/8"")"
            Height          =   210
            Index           =   2
            Left            =   90
            TabIndex        =   28
            Top             =   480
            Width           =   2010
         End
         Begin VB.OptionButton OptCols 
            Caption         =   "2 (2"" x 4"")"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   27
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton OptCols 
            Caption         =   "2 (1"" x 4"")"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   255
            Value           =   -1  'True
            Width           =   1080
         End
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3375
         ItemData        =   "AffLabelRpt.frx":08CE
         Left            =   3120
         List            =   "AffLabelRpt.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   570
         Width           =   3165
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contact"
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2355
         Begin VB.OptionButton optContact 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   17
            Top             =   255
            Value           =   -1  'True
            Width           =   2010
         End
         Begin VB.OptionButton optContact 
            Caption         =   "Affidavit Contact"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   18
            Top             =   540
            Width           =   1935
         End
         Begin VB.OptionButton optContact 
            Caption         =   "Titles"
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   19
            Top             =   870
            Width           =   2010
         End
      End
      Begin VB.TextBox txtOnAirDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   900
         Width           =   945
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   2565
         Begin VB.OptionButton optSP 
            Caption         =   "People"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   9
            Top             =   0
            Width           =   2010
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Stations"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1830
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Both"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   225
         Width           =   1935
      End
      Begin VB.Label lacContact 
         Caption         =   "Contact"
         Height          =   270
         Left            =   3120
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date:"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   930
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4485
      TabIndex        =   24
      Top             =   1185
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4245
      TabIndex        =   23
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   22
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
      Top             =   105
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffLabelRpt.frx":08D2
         Left            =   1065
         List            =   "AffLabelRpt.frx":08D4
         TabIndex        =   4
         Top             =   840
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1185
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   810
         Width           =   765
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmLabelRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmLabelRpt - Mailing Labels
'*
'*  Created July,1998 by Dick LeVine
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
    'Dim NewForm As New frmViewReport
    
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
    'grdVehAff.MoveFirst
    'For i = 0 To grdVehAff.Rows
    '    grdVehAff.SelBookmarks.Add grdVehAff.Bookmark
    '    grdVehAff.MoveNext
    'Next i
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllContacts_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
   
End Sub

Private Sub cmdDone_Click()
    Unload frmLabelRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles, sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sStationType As String
    Dim iNoCDs As Integer
    Dim iLoop As Integer
    Dim iType As Integer
    Dim sContact As String
    Dim sOutput As String
    Dim ilIdx As Integer
    Dim slTest As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    'Dim NewForm As New frmViewReport
                    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbDefault
    ' Validate Input Info
    If lbcVehAff.SelCount <= 0 Then
        gMsgBox "Vehicle must be selected.", vbOKOnly
        Exit Sub
    End If
    
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
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    '1-4-06 contact selection changed to list box selection
    'If optContact(1).Value Then       'Program Director
    '    sContact = "1"
    'ElseIf optContact(2).Value Then   'Music Director
    '    sContact = "2"
    'ElseIf optContact(3).Value Then   'Traffic Director
    '    sContact = "3"
    'ElseIf optContact(4).Value Then   'Affidavit Contact
    '    sContact = "4"
    'Else
    '    sContact = "0"                'None
    'End If

    
   ' CRpt1.Connect = "DSN = " & sgDatabaseName
   ' If optRptDest(0).Value = True Then           'Display
   '     CRpt1.Destination = crptToWindow
   ' ElseIf optRptDest(1).Value = True Then       'Print
   '     CRpt1.Destination = crptToPrinter
   ' ElseIf optRptDest(2).Value = True Then       'File
   '     gOutputMethod frmLabelRpt, "LabelRpt.rpt", sOutput
   ' Else
   '     Screen.MousePointer = vbDefault
   '     Exit Sub
   ' End If
    'Retrieve information from report selection
    If optSP(0).Value Then                          'stations
        sStationType = "shttType = 0"
    ElseIf optSP(1).Value Then                      'people
        sStationType = "shttType = 1"
    Else
        sStationType = ""                           'both
    End If
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sgStdDate = sStartDate
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sDateRange = "(attOffAir >= '" & Format$(sStartDate, sgSQLDateForm) & "') And (attDropDate >= '" & Format$(sStartDate, sgSQLDateForm) & "') And (attOnAir <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    sVehicles = ""
    
    If chkListBox.Value = 0 Then    ' = 0 Then                        'User did NOT select all vehicles
        For i = 0 To lbcVehAff.ListCount - 1 Step 1
            If lbcVehAff.Selected(i) Then
                If Len(sVehicles) = 0 Then
                    sVehicles = "(attVefCode = " & lbcVehAff.ItemData(i) & ")"
                Else
                    sVehicles = sVehicles & " OR (attVefCode = " & lbcVehAff.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    
    If optContact(2).Value Then    'use titles name
        For i = 0 To lbcContacts.ListCount - 1 Step 1
            If lbcContacts.Selected(i) Then
                sContact = Trim$(Str(lbcContacts.ItemData(i)))
                Exit For
            End If
        Next i
    End If
        
    'Generate the report

    'Determine Max number of CD
    If sStationType <> "" Then
        SQLQuery = "Select MAX(attNoCDs) from att, shtt"
        'SQLQuery = "Select MAX(attNoCDs) from att, shtt, vef"
    Else
        SQLQuery = "Select MAX(attNoCDs) from att"
        'SQLQuery = "Select MAX(attNoCDs) from att, vef"
    End If
    SQLQuery = SQLQuery + " WHERE ((" & sDateRange & ")"
    If sStationType <> "" Then
        SQLQuery = SQLQuery + " AND ((attShfCode = shttCode)" & " And " & sStationType & ")"
    End If
    If sVehicles <> "" Then
        SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
    End If
    SQLQuery = SQLQuery + ")"
    Set rst = gSQLSelectCall(SQLQuery)
    'D.S. Avoid invalid use of Null error
    If rst(0).Value > 0 Then
        iNoCDs = rst(0).Value
        For iLoop = 1 To iNoCDs Step 1
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM VEF_Vehicles, shtt, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode"
            SQLQuery = SQLQuery + " AND attshfCode = shttCode  "
            SQLQuery = SQLQuery + " AND attNoCDs >= " + Trim$(Str$(iLoop))
            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
            If sStationType <> "" Then
                SQLQuery = SQLQuery + " AND (" & sStationType & ")"
            End If
            If sVehicles <> "" Then
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
            
            If OptCols(1).Value = True Then         '2" x 4", sort by the new show length field
                SQLQuery = SQLQuery + ")" + "ORDER BY attLabelID, shttCallLetters"
            Else
                SQLQuery = SQLQuery + ")" + "ORDER BY shttCallLetters"
            End If
       
                    
            If OptCols(0).Value Then
                '2 column - manf. EXP # 00517, Avery 05161 & 05261 1" x 4"
                'CRpt1.ReportFileName = sgReportDirectory + "afLabels.rpt"
                slRptName = "aflabels.rpt"
            ElseIf OptCols(1).Value Then
                slRptName = "aflabship.rpt"
            Else
                '3 column - manf. Avery # 05150 & 05260  1" x 2 5/8"
                'CRpt1.ReportFileName = sgReportDirectory + "AfLabel3.rpt"
                slRptName = "aflabel3.rpt"
            End If
            
    
    '        CRpt1.Formulas(0) = "StartDate = Date(" + Format$(sStartDate, "yyyy") + "," + Format$(sStartDate, "mm") + "," + Format$(sStartDate, "dd") + ")"
    '        CRpt1.Formulas(1) = "EndDate = Date(" + Format$(sEndDate, "yyyy") + "," + Format$(sEndDate, "mm") + "," + Format$(sEndDate, "dd") + ")"
    '        CRpt1.Formulas(2) = "Contact = " & sContact
    '        CRpt1.Action = 1
    '        CRpt1.Formulas(0) = ""
    '        CRpt1.Formulas(1) = ""
    '        CRpt1.Formulas(2) = ""
    
            sgCrystlFormula1 = "Date(" + Format$(sStartDate, "yyyy") + "," + Format$(sStartDate, "mm") + "," + Format$(sStartDate, "dd") + ")" 'StartDate
            sgCrystlFormula2 = "Date(" + Format$(sEndDate, "yyyy") + "," + Format$(sEndDate, "mm") + "," + Format$(sEndDate, "dd") + ")" 'EndDate
            sgCrystlFormula3 = sContact 'Contact
             If optContact(0).Value Then             'no contact information
                sgCrystlFormula4 = "0"
            ElseIf optContact(1).Value Then       'use affidavit contact name
                sgCrystlFormula4 = "1"
            Else
                sgCrystlFormula4 = "2"
            End If
            
            slExportName = "Aflabel" & CStr(iLoop)
            
            If optRptDest(0).Value = True Then 'Display
                ilRptDest = 0
            ElseIf optRptDest(1).Value = True Then 'Print
                ilRptDest = 1
            Else
                ilRptDest = 2 'File
                'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
                ilExportType = cboFileType.ListIndex        '3-15-04
            End If
            If iLoop = 1 Then
                gUserActivityLog "E", sgReportListName & ": Prepass"
            End If
            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
        Next iLoop
    
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True

        Screen.MousePointer = vbDefault
        'If optRptDest(2).Value = True Then
        '    gMsgBox "Output Sent To: " & sOutput, vbInformation
        'End If
    Else
        gMsgBox "No mailing labels were found for " & lbcVehAff.Text & " " & sStartDate & " - " & sEndDate
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    gHandleError "", "frmLAbelRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmLabelRpt
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmLabelRpt
    gCenterForm frmLabelRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmLabelRpt.Caption = "Mailing Labels Report - " & sgClientName
    imChkListBoxIgnore = False
    'SQLQuery = "SELECT vef.vefName from vef WHERE ((vef.vefvefCode = 0 AND vef.vefType = 'C') OR vef.vefType = 'L' OR vef.vefType = 'A')"
    'SQLQuery = SQLQuery + " ORDER BY vef.vefName"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    grdVehAff.AddItem "" & rst(0).Value & ""
    '    rst.MoveNext
    'Wend
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = Format$(slDate, sgShowDateForm)
    CalOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = False     'current defaulted to display, disallow export type selectivity
    ilRet = gPopTitleNames()        'get all the title names
    'place titles in list box for selection
    For iLoop = 0 To UBound(tgTitleInfo) - 1 Step 1
        If iLoop = 0 Then
            lbcContacts.AddItem "None"
        End If
        lbcContacts.AddItem Trim$(tgTitleInfo(iLoop).sTitle)
        lbcContacts.ItemData(lbcContacts.NewIndex) = tgTitleInfo(iLoop).iCode
    Next iLoop
    lbcContacts.ListIndex = 0           'default to none
    lbcVehAff.Height = lbcContacts.Height + (lbcContacts.Top - lbcVehAff.Top)    'calc fullheight of vehicle box
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmLabelRpt = Nothing
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

Private Sub optContact_Click(Index As Integer)
    If Index < 2 Then           'no contact or using affidavit contact, dont show titles list box
        lbcContacts.Visible = False
        lacContact.Visible = False
        lbcVehAff.Height = lbcContacts.Height + (lbcContacts.Top - lbcVehAff.Top)    'calc fullheight of vehicle box
    Else
        lbcContacts.Visible = True
        lacContact.Visible = True
        lbcVehAff.Height = lacContact.Top - lbcVehAff.Top - 120
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

