VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrLibRpt 
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8160
   Begin VB.Frame frcOption 
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
      Height          =   3660
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   7575
      Begin VB.TextBox edcChangeDateTo 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox edcChangeDateFrom 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frcOldNew 
         Caption         =   "Show"
         Height          =   615
         Left            =   105
         TabIndex        =   10
         Top             =   2895
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optOldNew 
            Caption         =   "Current"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optOldNew 
            Caption         =   "History"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox ckcAllBuses 
         Caption         =   "All Buses"
         Height          =   255
         Left            =   5280
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox ckcAllSub 
         Caption         =   "All Subnames"
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   120
         Width           =   1455
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort By"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   1065
         Width           =   2535
         Begin VB.OptionButton optSortBy 
            Caption         =   "Library"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Bus Name"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox edcTo 
         Height          =   285
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   17
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox edcFrom 
         Height          =   285
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   15
         Top             =   270
         Width           =   1095
      End
      Begin VB.CheckBox ckcAllBusGroup 
         Caption         =   "All Bus Groups"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox ckcAllLibs 
         Caption         =   "All Libraries"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
      Begin VB.ListBox lbcBus 
         Height          =   1230
         ItemData        =   "EngrLibRpt.frx":0000
         Left            =   5280
         List            =   "EngrLibRpt.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1995
      End
      Begin VB.ListBox lbcBusGroup 
         Height          =   1230
         ItemData        =   "EngrLibRpt.frx":0004
         Left            =   3120
         List            =   "EngrLibRpt.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1995
      End
      Begin VB.ListBox lbcSubLib 
         Height          =   1230
         ItemData        =   "EngrLibRpt.frx":0008
         Left            =   5280
         List            =   "EngrLibRpt.frx":000A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   1995
      End
      Begin VB.ListBox lbcLibrary 
         Height          =   1230
         ItemData        =   "EngrLibRpt.frx":000C
         Left            =   3120
         List            =   "EngrLibRpt.frx":000E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label lacChangeDateTo 
         Caption         =   "To"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lacChangeDateFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lacChangeDates 
         Caption         =   "Enter change dates- "
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lacTo 
         Caption         =   "To"
         Height          =   255
         Left            =   690
         TabIndex        =   16
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lacFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   690
         TabIndex        =   14
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lacDates 
         Caption         =   "Dates-"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   585
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
      FormDesignHeight=   5670
      FormDesignWidth =   8160
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4455
      TabIndex        =   9
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   7
      Top             =   240
      Width           =   2685
   End
   Begin VB.Frame frcOutput 
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
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   2190
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2310
      End
   End
End
Attribute VB_Name = "EngrLibRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'*  EngrLibRpt - Create a report to show selective libraries by library name
'*               and or bus groups
'
'*
'*  Created September,  2004
'*
'*  Copyright Counterpoint Software, Inc.
'****************************************************************************
Option Explicit

'Dim WithEvents rstLibrary As ADODB.Recordset
Dim imLibChkListBoxIgnore As Integer        'All library check box flag
Dim imGroupChkListBoxIgnore As Integer      'all bus group check box flag
Dim imSubChkListBoxIgnore As Integer        'all subnames check box flag
Dim imBusChkListBoxIgnore As Integer        'all buses check box flag
Dim smBSEStamp As String
Dim smDBEStamp As String
Dim tmBSE() As BSE
Dim tmCTE As CTE                            'description table
Dim tmDBE() As DBE
Dim tmDHE As DHE
Dim tmAIE As AIE


Private Sub ckcAllBuses_Click()
Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long
    If imBusChkListBoxIgnore Then           'ignore doing anything to the list box entries
        Exit Sub
    End If
    If ckcAllBuses.Value = vbChecked Then     'if check box is on, select all entries in list box
        iValue = True
    Else
        iValue = False                      'if check box is off, deselect all entries in list box
    End If
    
    If lbcBus.ListCount > 0 Then         'at least 1 entries exists in check box
        imBusChkListBoxIgnore = True
        lRg = CLng(lbcBus.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcBus.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imBusChkListBoxIgnore = False
    End If
    

End Sub

Private Sub ckcAllBusGroup_Click()

Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long
Dim ilLoopGroups As Integer
Dim ilBGECode  As Integer
Dim ilRet As Integer


    If imGroupChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllBusGroup.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcBusGroup.ListCount > 0 Then
        imGroupChkListBoxIgnore = True
        lRg = CLng(lbcBusGroup.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcBusGroup.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imGroupChkListBoxIgnore = False
    End If
    
    mSetBuses           'set the bus name on based on the bus group selected
    
    
End Sub

Private Sub ckcAllLibs_Click()

Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long

    If imLibChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllLibs.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcLibrary.ListCount > 0 Then
        imLibChkListBoxIgnore = True
        lRg = CLng(lbcLibrary.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcLibrary.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imLibChkListBoxIgnore = False
    End If
    
    mSetSubNames      'gather and show/not show the associated subnames for the selected library

End Sub

Private Sub ckcAllSub_Click()
Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long

    If imSubChkListBoxIgnore Then           'ignore doing anything to the list box entries
        Exit Sub
    End If
    If ckcAllSub.Value = vbChecked Then     'if check box is on, select all entries in list box
        iValue = True
    Else
        iValue = False                      'if check box is off, deselect all entries in list box
    End If
    
    If lbcSubLib.ListCount > 0 Then         'at least 1 entries exists in check box
        imSubChkListBoxIgnore = True
        lRg = CLng(lbcSubLib.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSubLib.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imSubChkListBoxIgnore = False
    End If
End Sub

Private Sub cmdDone_Click()
    Unload EngrLibRpt
End Sub

Private Sub cmdReport_Click()

    Dim ilRet As Integer            'return error from subs/functions
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slRptName As String         'full report name of crystal .rpt
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim slSQLQuery As String        'formatting of sql query for selective libraries
    Dim ilLoop As Integer           'temp variable
    Dim slDate As String
    Dim slSQLFromDate As String     'user entered full from date for formatting sql call
    Dim slSQLToDAte As String       'user entered full to date for formatting sql call
    Dim llLoopLib As Long
    Dim slSQLDateQuery As String    'formttted sql string of user entered dates
    Dim slSqlSubQuery As String     'formattied sql string for subnames selection
    Dim slStr As String             'temp string handling
    Dim llLibCode As Long           'DHE library code
    Dim ilFound As Integer
    Dim ilDay As Integer
    Dim slDHEStamp As String
    Dim ilDheLoop As Integer
    Dim llDate As Long
    Dim llResult As Long
    Dim ilAIELoop As Integer
    Dim ilPass As Integer
    Dim slSQLChgDateFrom As String
    Dim slSQLChgDateTo As String
    Dim slAIEStamp As String
    Dim tlAIE() As AIE
    Dim tlDHE() As DHE
    Dim slRptType As String
    Dim ilVersion As Integer
    Dim llLoopTemp As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass


    If optRptDest(0).Value = True Then
       ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilRptDest = 2
        ilExportType = cboFileType.ListIndex
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    slSQLFromDate = gEditDateInput(edcFrom.text, "1/1/1970")   'check if from date is valid; if no date entered, set earliest possible
    If slSQLFromDate = "" Then  'if no returned date, its invalid
        edcFrom.SetFocus
        Exit Sub
    End If
    sgCrystlFormula2 = "Date(" + Format$(gAdjYear(slSQLFromDate), "yyyy") + "," + Format$(slSQLFromDate, "mm") + "," + Format$(slSQLFromDate, "dd") + ")"

    
    slSQLToDAte = gEditDateInput(edcTo.text, "12/31/2069")   'check if from date is valid; if no date entered, set latest possible
    If slSQLToDAte = "" Then  'if no returned date, its invalid
        edcTo.SetFocus
        Exit Sub
    End If
    sgCrystlFormula3 = "Date(" + Format$(gAdjYear(slSQLToDAte), "yyyy") + "," + Format$(slSQLToDAte, "mm") + "," + Format$(slSQLToDAte, "dd") + ")"

    If optSortBy(0).Value = True Then       'sort by library
        sgCrystlFormula4 = "'L'"
    Else
        sgCrystlFormula4 = "'B'"              'sort by Bus
    End If
    
    'if history, check change date range entered for validity
    If optOldNew(1).Value = True Then           'history
        slSQLChgDateFrom = gEditDateInput(edcChangeDateFrom.text, "1/1/1970")   'check if from date is valid; if no date entered, set earliest possible
        If slSQLChgDateFrom = "" Then  'if no returned date, its invalid
            edcChangeDateFrom.SetFocus
            Exit Sub
        End If
        sgCrystlFormula2 = "Date(" + Format$(gAdjYear(slSQLChgDateFrom), "yyyy") + "," + Format$(slSQLChgDateFrom, "mm") + "," + Format$(slSQLChgDateFrom, "dd") + ")"
    
        
        slSQLChgDateTo = gEditDateInput(edcChangeDateTo.text, "12/31/2069")   'check if from date is valid; if no date entered, set latest possible
        If slSQLChgDateTo = "" Then  'if no returned date, its invalid
            edcChangeDateTo.SetFocus
            Exit Sub
        End If
        sgCrystlFormula3 = "Date(" + Format$(gAdjYear(slSQLChgDateTo), "yyyy") + "," + Format$(slSQLChgDateTo, "mm") + "," + Format$(slSQLChgDateTo, "dd") + ")"
        sgCrystlFormula4 = "'L'"        'sortby - common formulas sent to reports
    End If
   
    Set rstLibrary = New Recordset
    gGeneraterstLibrary     'generate the ddfs for report
    
    rstLibrary.Open
      
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstLibrary, sgDBPath & "\library.ttx", True)
       
    If optOldNew(0).Value = True Then           'current
        slRptType = "Sum.rpt"
        'obtain all the valid library headers
        For llLoopLib = 0 To lbcLibrary.ListCount - 1
            If lbcLibrary.Selected(llLoopLib) Then          'test if user selected this entry
                llLibCode = lbcLibrary.ItemData(llLoopLib)
                'ilRet = gGetRec_DHE_DayHeaderInfo(llLibCode, "EngrLibRep-mcmdReport_click", tmDHE)
                ilRet = gGetRecs_DHE_DayHeaderInfoByLibrary(slDHEStamp, llLibCode, "EngrLibRpt: cmdReport", tlDHE())
    
                For ilDheLoop = LBound(tlDHE) To UBound(tlDHE) - 1
                    LSet tmDHE = tlDHE(ilDheLoop)
                    'Check this library header to see if it passes the date filters and subname filter
                    ilFound = mfilterSelectivity(slSQLFromDate, slSQLToDAte)
                   
                    If ilFound Then         'valid library header
                        mAddRstLibrary False, 0 'create a new entry to be printed
                    End If                  'ilfound: invalid dates or subname
                Next ilDheLoop          'next library for the same name
            End If                      'selected library
            slDHEStamp = ""             'force to reread with new library name
        Next llLoopLib                     'obtain next library selected
     
     'History
     Else
        slRptType = "Hist.rpt"
        'obtain the changes from Activity file to determine which changes to show history
        ilRet = gGetTypeOfRecs_AIE_ActiveInfo("DHE", slSQLChgDateFrom, slSQLChgDateTo, slAIEStamp, "EngrLibRpt", tlAIE())
        For ilAIELoop = LBound(tlAIE) To UBound(tlAIE) - 1
            LSet tmAIE = tlAIE(ilAIELoop)
            For ilPass = 1 To 2
                If ilPass = 1 Then      'get the current
                    llLibCode = tmAIE.lToFileCode
                Else                    'get the past
                    llLibCode = tmAIE.lFromFileCode
                End If
                ilRet = gGetRec_DHE_DayHeaderInfo(llLibCode, "EngrLibRpt: gGetRec_DHE_DayHeaderINfo- cmdReport", tmDHE)
                If ilPass = 1 Then
                    ilVersion = tmDHE.iVersion
                End If
                'test for selected library names
                For llLoopLib = 0 To lbcLibrary.ListCount - 1
                    If lbcLibrary.Selected(llLoopLib) Then          'test if user selected this entry
                        If lbcLibrary.ItemData(llLoopLib) = tmDHE.lDneCode Then
                        
                            ilFound = False
                            'Insure the subname found with this library should be included
                            If Not ckcAllSub.Value = vbChecked Then     'if not all checked, then build the sql query for the selected subnames
                                For llLoopTemp = 0 To lbcSubLib.ListCount - 1
                                       'if the subname is selected and it matches the subname in the processing library, ok to output this library
                                    If lbcSubLib.Selected(llLoopTemp) And (lbcSubLib.ItemData(llLoopTemp) = tmDHE.lDseCode) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next llLoopTemp
                            Else
                                ilFound = True               'valid library to include with this subname
                            End If
                            If ilFound Then
                                mAddRstLibrary True, ilVersion  'create a new entry to be printed
                                Exit For
                            End If
                        End If
                    End If
                Next llLoopLib
            Next ilPass
        Next ilAIELoop
     End If
    
    
    'debugging only
    'rstLibrary.MoveFirst
    'While Not rstLibrary.EOF
    '    slStr = rstLibrary.Fields("Name").Value
    '    slStr = slStr & "," & rstLibrary.Fields("SubName").Value
    '    rstLibrary.MoveNext
    'Wend
    
    'igRptSource = vbModeless       set this in minit
    
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstLibrary
    'Set fNewForm.Report = Appl.OpenReport(sgReportDirectory + slRptName & "Sum.rpt")
    'fNewForm.Report.Database.Tables(1).SetDataSource rstLibrary, 3
    'fNewForm.Show igRptSource
    
    Screen.MousePointer = vbDefault
    

    'rstLibrary.Close           'causes error when closed
    Set rstLibrary = Nothing
    If igRptSource = vbModal Then
        Unload EngrLibRpt
    End If
    
    
    Erase tlDHE, tlAIE
    Erase tmDBE
    Erase tmBSE
    
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in EngrLibRpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in EngrLibRpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrLibRpt
End Sub

Private Sub edcFrom_GotFocus()
    gCtrlGotFocus edcFrom
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrLibRpt
    gCenterForm EngrLibRpt
End Sub
Private Sub Form_Load()
Dim ilRet As Integer

On Error GoTo ErrHand:
    mInit
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set EngrLibRpt = Nothing
End Sub
Private Sub lbcBus_Click()
    If imBusChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllBuses.Value = vbChecked Then
        imBusChkListBoxIgnore = True
        ckcAllBuses.Value = False
        imBusChkListBoxIgnore = False
    End If
End Sub

Private Sub lbcBusGroup_Click()

    If imGroupChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllBusGroup.Value = vbChecked Then
        imGroupChkListBoxIgnore = True
        ckcAllBusGroup.Value = False
        imGroupChkListBoxIgnore = False
    End If
    
    mSetBuses           'set the bus name on based on the bus group selected
  
End Sub
Private Sub lbcLibrary_Click()

    If imLibChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllLibs.Value = vbChecked Then
        imLibChkListBoxIgnore = True
        ckcAllLibs.Value = False
        imLibChkListBoxIgnore = False
    End If
    
    mSetSubNames      'gather and show/not show the associated subnames for the selected library

    Exit Sub
End Sub

Private Sub lbcSubLib_Click()
    If imSubChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllSub.Value = vbChecked Then
        imSubChkListBoxIgnore = True
        ckcAllSub.Value = False
        imSubChkListBoxIgnore = False
    End If
End Sub

Private Sub optOldNew_Click(Index As Integer)
    If Index = 0 Then
        lacChangeDateFrom.Visible = False
        lacChangeDateTo.Visible = False
        edcChangeDateFrom.Visible = False
        edcChangeDateTo.Visible = False
        lacChangeDates.Visible = False
        
        lacDates.Visible = True
        lacFrom.Visible = True
        edcFrom.Visible = True
        lacTo.Visible = True
        edcTo.Visible = True
        frcSortBy.Visible = True
    Else
        lacChangeDates.Move 120, lacDates.Top
        lacChangeDateFrom.Move 120, lacChangeDates.Top + lacChangeDates.Height + 30
        edcChangeDateFrom.Move 1200, lacChangeDateFrom.Top
        lacChangeDateTo.Move 120, edcChangeDateFrom.Top + edcChangeDateTo.Height + 30
        edcChangeDateTo.Move 1200, lacChangeDateTo.Top
        
        lacDates.Visible = False
        lacFrom.Visible = False
        edcFrom.Visible = False
        lacTo.Visible = False
        edcTo.Visible = False
        
        lacChangeDateFrom.Visible = True
        lacChangeDateTo.Visible = True
        edcChangeDateFrom.Visible = True
        edcChangeDateTo.Visible = True
        lacChangeDates.Visible = True
        frcSortBy.Visible = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub
Private Sub mPopLibraryNames()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrDNEStamp, "EngrLibRpt-mPopulate Library Definition", tgCurrDNE())
    lbcLibrary.Clear
    For ilLoop = 0 To UBound(tgCurrDNE) - 1 Step 1
        lbcLibrary.AddItem Trim$(tgCurrDNE(ilLoop).sName)
        lbcLibrary.ItemData(lbcLibrary.NewIndex) = tgCurrDNE(ilLoop).lCode
    Next ilLoop
    Exit Sub
End Sub
Private Sub mInit()
Dim ilRet As Integer
    If igRptSource = vbModal Then       'if coming from a task, disallow to return to report list
                                        'need to return to the task
        cmdReturn.Enabled = False
    Else
        cmdReturn.Enabled = True
        igRptSource = vbModeless
    End If
    gPopExportTypes cboFileType  'populate the valid export types
    cboFileType.Enabled = False  'disable export file types until SAVE TO selected
    gChangeCaption frcOption     'show report name caption on selectivity box
    mPopLibraryNames                  'populate library names
    mPopBusGroup
    mPopBusNames
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", sgCurrLibDHEStamp, "EngrLibRpt-lbcLibrary_click", tgCurrLibDHE())

    Exit Sub
End Sub
Private Sub mPopBusGroup()
Dim ilRet As Integer
Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrLibRpt-mPopBusGroup", tgCurrBGE())
    lbcBusGroup.Clear
    For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
        lbcBusGroup.AddItem Trim$(tgCurrBGE(ilLoop).sName)
        lbcBusGroup.ItemData(lbcBusGroup.NewIndex) = tgCurrBGE(ilLoop).iCode
    Next ilLoop
    Exit Sub
End Sub
Public Sub mPopBusNames()
Dim ilRet As Integer
Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("G", sgCurrBDEStamp, "EngrLibRpt-mPopBusNames", tgCurrBDE())
    lbcBus.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        lbcBus.AddItem Trim$(tgCurrBDE(ilLoop).sName)
        lbcBus.ItemData(lbcBus.NewIndex) = tgCurrBDE(ilLoop).iCode
    Next ilLoop
    Exit Sub
End Sub
'
'
'           mPopSubLib - populate the subnames.  If user selected one library, populate
'           with only those subnames that reference the library name.
'           i.e. Paul Harvey has 3 libraries using Paul Harvey as the DHE library name
'                   Each of the subnames for the 3 libraries could be: Morning, AFternoon, Evening
'                if Paul Harvey selected, show subnames Mroning, Afternoon & evening
'           <input> ilAllFlag : 0 = get all subnames, not just a match for the library
'                             : non-0  get matches only
Public Sub mPopSubLib(ilAllFlag As Integer)
Dim ilRet As Integer
Dim llLoop As Integer
Dim llHeader As Long
Dim llIndex As Long
Dim llLoopDSE As Long
    'Get both current and history subnames
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrLibRpt-mPopDSE Day Subname", tgCurrDSE())
    lbcSubLib.Clear
     
    If ilAllFlag = 0 Then
        'all or more than 1 library selected, populate all the sublib names
        For llLoop = 0 To UBound(tgCurrDSE) - 1 Step 1
            lbcSubLib.AddItem Trim$(tgCurrDSE(llLoop).sName)
            lbcSubLib.ItemData(lbcSubLib.NewIndex) = tgCurrDSE(llLoop).lCode
        Next llLoop
    Else
        'single selection, only build in the subnames that match the library
        llIndex = lbcLibrary.ListIndex
        llHeader = lbcLibrary.ItemData(llIndex)
        For llLoop = 0 To UBound(tgCurrLibDHE) - 1 Step 1
            If tgCurrLibDHE(llLoop).lDneCode = llHeader And tgCurrLibDHE(llLoop).sState = "A" Then  '5-17-06
                For llLoopDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                    If tgCurrDSE(llLoopDSE).lCode = tgCurrLibDHE(llLoop).lDseCode Then
                        lbcSubLib.AddItem Trim$(tgCurrDSE(llLoopDSE).sName)
                        lbcSubLib.ItemData(lbcSubLib.NewIndex) = tgCurrDSE(llLoopDSE).lCode
                        Exit For
                    End If
                Next llLoopDSE
            End If
        Next llLoop
    End If
    
    Exit Sub
End Sub

Public Sub mSetSubNames()
Dim llIndex As Long
Dim llDheCode As Long
Dim ilRet As Integer

    'a library selected, find all subnames for the matching library only if a single selection;
    'otherwise, use all subnames
 
    If lbcLibrary.SelCount = 1 Then     'only one selected
        'llIndex = lbcLibrary.ListIndex
        'llDheCode = lbcLibrary.ItemData(llIndex)     'get library code to access all the matching libraries for the selection
        'get the matching current library headers
 
        'ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", sgCurrLibDHEStamp, "EngrLibRpt-lbcLibrary_click", tgCurrLibDHE())
        'populate list box of all subnames that reference this library
        mPopSubLib 1
        lbcSubLib.Visible = True
        ckcAllSub.Visible = True
        ckcAllSub.Value = vbUnchecked
    ElseIf lbcLibrary.SelCount = 0 Then         'none selected, check box turned off
        ckcAllSub.Value = vbUnchecked
        lbcSubLib.Clear
        ckcAllSub.Visible = True
        lbcSubLib.Visible = True
    Else                                        'more than 1 selected, use all subnames
        lbcSubLib.Visible = False       'more than 1 library selected, dont show the subnames , use all that belong to any library selected
        lbcSubLib.Clear
        mPopSubLib 0                    'populate all subnames
        ckcAllSub.Visible = False
        ckcAllSub.Value = vbChecked     'force event to select all
    End If
End Sub
'
'
'               mSetBuses - set the bus name selection on in the list box
'               based on the bus group selected
'
Public Sub mSetBuses()
Dim lRg As Long
Dim lRet As Long
Dim ilLoopGroups As Integer
Dim ilBGECode As Integer
Dim ilRet As Integer
Dim ilMatch As Integer
Dim ilLoop As Integer

    lRg = CLng(lbcBus.ListCount - 1) * &H10000 Or 0
    lRet = SendMessageByNum(lbcBus.hwnd, LB_SELITEMRANGE, False, lRg)   'turn off all the buses to reselect the ones in the groups selected
    ckcAllBuses.Value = vbUnchecked
    
    For ilLoopGroups = 0 To lbcBusGroup.ListCount - 1
        If lbcBusGroup.Selected(ilLoopGroups) Then
            ilBGECode = lbcBusGroup.ItemData(ilLoopGroups)
            ilRet = gGetRecs_BSE_BusSelGroup("G", smBSEStamp, ilBGECode, "lbcBusGroup_click", tmBSE())
        
            For ilMatch = 0 To UBound(tmBSE) - 1    'set selection for all bus names in this group
                For ilLoop = 0 To lbcBus.ListCount - 1
                    If tmBSE(ilMatch).iBdeCode = lbcBus.ItemData(ilLoop) Then
                        'force selection of the bus name
                        lbcBus.Selected(ilLoop) = True
                        Exit For
                    End If
                Next ilLoop
            Next ilMatch
        End If
    Next ilLoopGroups
End Sub
'
'           mFilterSelectivity - determine if this library has user selected dates,
'           correct subnames and bus groups
'           <input> slSQLFromDate - earliest date of library to retrieve
'                   slSQLToDate - latest date of library to retrieve


Private Function mfilterSelectivity(slSQLFromDate As String, slSQLToDAte As String) As Integer
Dim ilFoundSub As Integer
Dim llLoopTemp As Long
Dim ilValidDates As Integer

    ilFoundSub = False
     
    'Insure the subname found with this library should be included
    If Not ckcAllSub.Value = vbChecked Then     'if not all checked, then build the sql query for the selected subnames
        For llLoopTemp = 0 To lbcSubLib.ListCount - 1
            'if the subname is selected and it matches the subname in the processing library, ok to output this library
            If lbcSubLib.Selected(llLoopTemp) And (lbcSubLib.ItemData(llLoopTemp) = tmDHE.lDseCode) Then
                    ilFoundSub = True
                Exit For
            End If
        Next llLoopTemp
    Else
        ilFoundSub = True               'valid library to include with this subname
    End If
    
       
    ilValidDates = False
    'dhe enddate >= inputstart & dhe startdate <= inputend
    'If Format$(tmDHE.sEndDate, sgSQLDateForm) > Format$(slSQLFromDate, sgSQLDateForm) And Format$(tmDHE.sStartDate, sgSQLDateForm) <= Format$(slSQLToDAte, sgSQLDateForm) Then
    If (gDateValue(tmDHE.sEndDate) >= gDateValue(slSQLFromDate)) And (gDateValue(tmDHE.sStartDate) <= gDateValue(slSQLToDAte)) Then
        'insure the dates of this library should be included
        ilValidDates = True
    End If
    
    If ilFoundSub And ilValidDates And tmDHE.sCurrent <> "N" Then        'passed all tests, must be current version
        mfilterSelectivity = True
    Else
        mfilterSelectivity = False
    End If
End Function
'           Create a record in the active data source for printing crystal report
'          <input> ilHistory - true if history report (vs current)
'                  ilVersion - for History option only:  each pair (current &previous will have
'                               the current version # in it
'
'       Sorting for History report-  Major to minor:
'                           Library Name,
'                           original Library record (to keep all version together);
'                           Version # (desc) .  Each activity record contains the current & previous versions.
'                                        Each pair will ahve the version # of the current to keep them together.
'                           Sub-version # (desc)- true version # of the current or previous version
'                           Sequence - if more than 1 bus for a library, maintain the seq # to keep it with
'                                       its parent
'       '2-21-06 add ignore conflict flag
Public Sub mAddRstLibrary(ilHistory As Integer, ilVersion As Integer)
Dim ilDBE As Integer
Dim ilDNE As Integer
Dim ilRet As Integer
Dim ilDSE As Integer
Dim llDate As Long
Dim slHour As String
Dim ilBDE As Integer
Dim ilSequence As Integer
Dim ilLoop As Integer
Dim slStr As String
Dim ilBus As Integer
Dim ilFoundBus As Integer
Dim ilBusCode As Integer

    ilSequence = 0
    
    'determine the bus groups for the library
    ilRet = gGetRecs_DBE_DayBusSel(smDBEStamp, tmDHE.lCode, "EngrLibRpt- cmdReport for DBE", tmDBE())

    'determine the buses
    For ilDBE = 0 To UBound(tmDBE) - 1 Step 1
        If tmDBE(ilDBE).sType = "B" Then
            'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '    If tmDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                ilBDE = gBinarySearchBDE(tmDBE(ilDBE).iBdeCode, tgCurrBDE())
                If ilBDE <> -1 Then
                    'determine if this bus definition selected
                    ilFoundBus = False
                    For ilBus = 0 To lbcBus.ListCount - 1
                        If lbcBus.Selected(ilBus) Then
                            ilBusCode = lbcBus.ItemData(ilBus)
                            If tmDBE(ilDBE).iBdeCode = ilBusCode Then
                                ilFoundBus = True
                                Exit For
                            End If
                        End If
                    Next ilBus
                    If ilFoundBus Then
                        slStr = Trim$(tgCurrBDE(ilBDE).sName)
                        rstLibrary.AddNew
                        ilSequence = ilSequence + 1
                        rstLibrary.Fields("Bus") = slStr
        
                        'get the library name
                        For ilDNE = 0 To UBound(tgCurrDNE) - 1 Step 1
                            If tmDHE.lDneCode = tgCurrDNE(ilDNE).lCode Then
                                rstLibrary.Fields("Name") = Trim$(tgCurrDNE(ilDNE).sName)
                                Exit For
                            End If
                        Next ilDNE
                        
                        'get the library subname
                        For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                            If tmDHE.lDseCode = tgCurrDSE(ilDSE).lCode Then
                                rstLibrary.Fields("SubName") = Trim$(tgCurrDSE(ilDSE).sName)
                                Exit For
                            End If
                        Next ilDSE
                        
                        ilRet = gGetRec_CTE_CommtsTitle(tmDHE.lCteCode, "EngrLibRpt-cmdReport_click for CTE", tmCTE)
                        rstLibrary.Fields("Desc") = Trim$(tmCTE.sComment)        'Library description
                                                      
                        rstLibrary.Fields("StartDate") = tmDHE.sStartDate
                        rstLibrary.Fields("EndDate") = tmDHE.sEndDate
                                       
                        llDate = gDateValue(tmDHE.sStartDate)
                        rstLibrary.Fields("StartDateSort") = llDate     'date for sorting when same library names
                        
                        slStr = gDayMap(tmDHE.sDays)      'format days of the week
                        rstLibrary.Fields("Days") = Trim$(slStr)
                        
                        rstLibrary.Fields("StartTime") = Trim$(tmDHE.sStartTime)
                        rstLibrary.Fields("Length") = gLongToLength(tmDHE.lLength, True)
                        
                        'format the 24 hours.  Each byte of the 24 bytes represents hour 0-23.
                        'format as 1-4; 1,5,10; 0-24 etc
                        slHour = tmDHE.sHours
                        slStr = gHourMap(slHour)
                        rstLibrary.Fields("Hours") = Trim$(slStr)
                        
                        rstLibrary.Fields("State") = tmDHE.sState
                        rstLibrary.Fields("Version") = ilVersion            'for sorting on history option
                        rstLibrary.Fields("SubVersion") = tmDHE.iVersion    'for sorting on history option
                        rstLibrary.Fields("Sequence") = ilSequence      'keep multiple lines of same library header together
                    
                        rstLibrary.Fields("DHEHeaderCode") = tmDHE.lCode    '5-18-06
                        '2-21-06 add ignore conflict flag from library header
                        rstLibrary.Fields("IgnoreConflicts") = tmDHE.sIgnoreConflicts
                        If ilHistory Then                   'append the rest of the fields for history
                            rstLibrary.Fields("DateEntered") = Format$(tmAIE.sEnteredDate, sgSQLDateForm)
                            rstLibrary.Fields("TimeEntered") = Format$(tmAIE.sEnteredTime, sgSQLTimeForm)
                            rstLibrary.Fields("OrigAIECode") = tmAIE.lCode
                            rstLibrary.Fields("FileCode") = tmAIE.lOrigFileCode
                            
                            For ilLoop = LBound(tgCurrUIE) To UBound(tgCurrUIE) - 1
                                If tgCurrUIE(ilLoop).iCode = tmAIE.iUieCode Then
                                    rstLibrary.Fields("User") = Trim$(tgCurrUIE(ilLoop).sShowName)
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                End If
            'Next ilBDE
        End If
    Next ilDBE
End Sub
