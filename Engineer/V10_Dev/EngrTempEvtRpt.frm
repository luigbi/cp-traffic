VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrTempEvtRpt 
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
      Caption         =   "Template Events Report Selection"
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
      Top             =   1860
      Width           =   7575
      Begin VB.CheckBox ckcAllFields 
         Caption         =   "Show All Fields"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   2415
      End
      Begin VB.TextBox edcChangeDateTo 
         Height          =   285
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2865
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox edcChangeDateFrom 
         Height          =   285
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2535
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frcOldNew 
         Caption         =   "Show"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1575
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optOldNew 
            Caption         =   "History"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optOldNew 
            Caption         =   "Current"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox ckcAllSub 
         Caption         =   "All Subnames"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox ckcAllLibs 
         Caption         =   "All Templates"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.ListBox lbcSubLib 
         Height          =   1035
         ItemData        =   "EngrTempEvtRpt.frx":0000
         Left            =   4080
         List            =   "EngrTempEvtRpt.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2190
         Width           =   3075
      End
      Begin VB.ListBox lbcLibrary 
         Height          =   1230
         ItemData        =   "EngrTempEvtRpt.frx":0004
         Left            =   4080
         List            =   "EngrTempEvtRpt.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   390
         Width           =   3075
      End
      Begin VB.Label lacChangeDateTo 
         Caption         =   "To"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lacChangeDateFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   2550
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lacChangeDates 
         Caption         =   "Enter change dates- "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2250
         Visible         =   0   'False
         Width           =   1575
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
Attribute VB_Name = "EngrTempEvtRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'*  EngrTempEvtRpt - Create a report to show selective template events by their
'               template name, days or dates
'
'*
'*  Created: 05-15-06
'*
'*  Copyright Counterpoint Software, Inc.
'****************************************************************************
Option Explicit

'Dim WithEvents rstLibEvts As ADODB.Recordset
Dim imLibChkListBoxIgnore As Integer        'All library check box flag
Dim imGroupChkListBoxIgnore As Integer      'all bus group check box flag
Dim imSubChkListBoxIgnore As Integer        'all subnames check box flag
Dim imBusChkListBoxIgnore As Integer        'all buses check box flag
Dim tmDee As DEE            'day event image
Dim tmDHE As DHE            'day library header image
Dim smDEEStamp As String
Dim tlDEE() As DEE
Dim tlDHE() As DHE
Dim tmAIE As AIE

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
    Unload EngrTempEvtRpt
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
    Dim slSQLChgDateFrom As String
    Dim slSQLChgDateTo As String
    Dim ilDNE As Integer            'temp looping for DNE Day names table
    Dim slStr As String             'temp string handling
    Dim slHour As String
    Dim ilDSE As Integer
    Dim ilBDE As Integer
    Dim llLibCode As Long           'DHE library code
    Dim ilFound As Integer
    Dim ilDay As Integer
    Dim llDee As Long
    Dim ilANE As Integer
    Dim ilValidDay As Integer
    Dim slDHEStamp As String
    Dim ilDheLoop As Integer
    Dim llDate As Long
    Dim llResult As Long
    Dim slRptType As String
    Dim slAIEStamp As String
    Dim tlAIE() As AIE
    Dim ilAIELoop As Integer
    Dim ilPass As Integer
    Dim ilVersion As Integer
    Dim llLoopTemp As Integer
    Dim llLibEvtCode As Long
    Dim ilTempDays(0 To 6) As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim ilWeekDay As Integer

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
        slStartDate = slSQLChgDateFrom
        slEndDate = slSQLChgDateTo
    End If
        
    Set rstLibEvts = New Recordset
    gGeneraterstLibEvts
    rstLibEvts.Open
      
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstLibEvts, sgDBPath & "\libEvts.ttx", True)
       
   
    If optOldNew(0).Value = True Then           'current
        slRptType = "Det.rpt"
        If ckcAllFields.Value = vbUnchecked Then        '12-2-11
            slRptType = "DetShort.rpt"
        End If
        'obtain all the valid library headers
        For llLoopLib = 0 To lbcLibrary.ListCount - 1
            If lbcLibrary.Selected(llLoopLib) Then          'test if user selected this entry
                llLibCode = lbcLibrary.ItemData(llLoopLib)
                ilRet = gGetRecs_DHE_DayHeaderInfoByLibrary(slDHEStamp, llLibCode, "EngrTempEvtRpt: cmdReport", tlDHE())
                For ilDheLoop = LBound(tlDHE) To UBound(tlDHE) - 1
                    LSet tmDHE = tlDHE(ilDheLoop)
                
                    'Check this library header to see if it passes the date filters and subname filter
                    ilFound = mfilterSelectivity()
                   
                    If ilFound And tmDHE.sCurrent = "Y" Then
                        mAddRstLibEvts False, 0  'create a new entry to be printed
                    End If                  'invalid dates or subname
                Next ilDheLoop          'for ilDheLoop = LBound(tlDhe) to UBound(tlDhe)
            End If                      'selected library
    
            slDHEStamp = ""             'force to reread with new library name
        Next llLoopLib                     'obtain next library selected
    Else                                'history
        slRptType = "Hist.rpt"
        'obtain the changes from Activity file to determine which changes to show history
        ilRet = gGetTypeOfRecs_AIE_ActiveInfo("DEE", slSQLChgDateFrom, slSQLChgDateTo, slAIEStamp, "EngrLibRpt", tlAIE())
        For ilAIELoop = LBound(tlAIE) To UBound(tlAIE) - 1
            LSet tmAIE = tlAIE(ilAIELoop)
            For ilPass = 1 To 2
                If ilPass = 1 Then      'get the current
                    llLibEvtCode = tmAIE.lToFileCode
                Else                    'get the past
                    llLibEvtCode = tmAIE.lFromFileCode
                End If
                ReDim tlDEE(0 To 1) As DEE
                ilRet = gGetRec_DEE_DayEvent(llLibEvtCode, "EngrTempEvtRpt: cmdReport_click", tlDEE(0))
                'obtain the associated library header to this event
                llLibCode = tlDEE(0).lDheCode
                ilRet = gGetRec_DHE_DayHeaderInfo(llLibCode, "EngrTempEvtRpt: cmdReport_click", tmDHE)
                If ilPass = 1 Then
                    ilVersion = tmDHE.iVersion
                End If
                'see if the library name is one selected
                For llLoopLib = 0 To lbcLibrary.ListCount - 1
                    If lbcLibrary.Selected(llLoopLib) Then          'test if user selected this entry
                        llLibCode = lbcLibrary.ItemData(llLoopLib)
                        
                        If llLibCode = tmDHE.lDneCode Then
                            'Check this library header to see if it passes the date filters and subname filter
                            ilFound = mfilterSelectivity()
                            
                            If ilFound Then
                                mAddRstLibEvts True, ilVersion  'create a new entry to be printed
                            End If                  'invalid dates or subname
                            Exit For
                        End If
                    End If                      'selected library
            
                    slDHEStamp = ""             'force to reread with new library name
                Next llLoopLib                     'obtain next library selected
            Next ilPass
        Next ilAIELoop
    End If                              'optOldNew(0).Value = True
    
    
    'debugging only
    'rstLibEvts.MoveFirst
    'While Not rstLibEvts.EOF
    '    slStr = rstLibEvts.Fields("Name").Value
    '    slStr = slStr & "," & rstLibEvts.Fields("EvStartTime").Value
    '    slStr = slStr & "," & rstLibEvts.Fields("Version").Value
    '    rstLibEvts.MoveNext
    'Wend
    
    'igRptSource = vbModeless       set in minit to return to caller
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, Trim(slRptName) & Trim(slRptType), slExportName, rstLibEvts
    'Set fNewForm.Report = Appl.OpenReport(sgReportDirectory + slRptName & "Sum.rpt")
    'fNewForm.Report.Database.Tables(1).SetDataSource rstLibEvts, 3
    'fNewForm.Show igRptSource
    
    Screen.MousePointer = vbDefault
    

    'rstLibEvts.Close           'causes error when closed
    Set rstLibEvts = Nothing
    If igRptSource = vbModal Then
        Unload EngrTempEvtRpt
    End If
    
    
    Erase tlDEE
    Erase tlDHE
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in User Rpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in User Rpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrTempEvtRpt
End Sub
Private Sub edcChangeDateFrom_GotFocus()
    gCtrlGotFocus edcChangeDateFrom
End Sub


Private Sub edcChangeDateTo_GotFocus()
  gCtrlGotFocus edcChangeDateTo
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrTempEvtRpt
    gCenterForm EngrTempEvtRpt
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
    Set EngrTempEvtRpt = Nothing
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
    Else
        lacChangeDateFrom.Move 120, lacChangeDates.Top + lacChangeDates.Height + 30
        edcChangeDateFrom.Move 1200, lacChangeDateFrom.Top
        lacChangeDateTo.Move 120, edcChangeDateFrom.Top + edcChangeDateTo.Height + 30
        edcChangeDateTo.Move 1200, lacChangeDateTo.Top
       
        lacChangeDateFrom.Visible = True
        lacChangeDateTo.Visible = True
        edcChangeDateFrom.Visible = True
        edcChangeDateTo.Visible = True
        lacChangeDates.Visible = True
  
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

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "T", sgCurrDNEStamp, "EngrTempEvtRpt-mPopulate Library Definition", tgCurrDNE())
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
    'populate the current libraries
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "T", sgCurrLibDHEStamp, "EngrTempEvtRpt-mInit DHE", tgCurrLibDHE())
    'populate the current event types
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrTempEvtRpt-mInit ETE", tgCurrETE())
    'populate the current bus controls
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrTempEvtRpt-mInit BusCCE", tgCurrBusCCE())
    'populate audio controls
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrTempEvtRpt-mInit Audio CCE", tgCurrAudioCCE())
    'populate the current start time types
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrTempEvtRpt-mInit TTE StartType", tgCurrStartTTE())
    'poulate the current end time types
    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrTempEvtRpt-mInit TTE EndType", tgCurrEndTTE())
    'populate material types
    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrTempEvtRpt-mInit MTE Material Types", tgCurrMTE())
    'populate Audio Names
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrTempEvtRpt-mInit ANE Audio Names", tgCurrANE())
    'populate Audio Sources
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrTempEvtRpt-mInit ASE Audio Source Names", tgCurrASE())
    'populate relays
    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrLibDef-mInit Relays RNE", tgCurrRNE())
    'populate silence codes
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrLibDef-mInit Silence SCE", tgCurrSCE())
    'populate netcue names
    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrLibDef-mInit Netcue NNE", tgCurrNNE())
    'follow names
    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrLibDef-mInit Follow Names FNE", tgCurrFNE())

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
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrTempEvtRpt-mPopDSE Day Subname", tgCurrDSE())
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
            If tgCurrLibDHE(llLoop).lDneCode = llHeader Then
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
 
        'ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", sgCurrLibDHEStamp, "EngrTempEvtRpt-lbcLibrary_click", tgCurrLibDHE())
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
'           mFilterSelectivity - determine if this library has user selected dates,
'           correct subnames and bus groups
Private Function mfilterSelectivity() As Integer
Dim ilFoundSub As Integer
Dim llLoopTemp As Long
Dim ilValidDates As Integer
Dim ilLoop As Integer
Dim ilValidDay As Integer       'day of week check

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
    
   
    If ilFoundSub Then         'passed all tests
        mfilterSelectivity = True
    Else
        mfilterSelectivity = False
    End If
End Function

'          Create a record in the active data source for printing crystal report
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
'       2-21-06 add ignore conflicts flag
Public Sub mAddRstLibEvts(ilHistory As Integer, ilVersion As Integer)
Dim ilRet As Integer            'return error from subs/functions
Dim ilLoop As Integer           'temp variable
Dim slDate As String
Dim llLoopLib As Long
Dim ilDNE As Integer            'temp looping for DNE Day names table
Dim slStr As String             'temp string handling
Dim slHour As String
Dim ilDSE As Integer
Dim ilBDE As Integer
Dim llLibCode As Long           'DHE library code
Dim ilFound As Integer
Dim ilDay As Integer
Dim llDee As Long
Dim ilANE As Integer
Dim ilValidDay As Integer
Dim ilDheLoop As Integer
Dim llDate As Long
Dim llResult As Long
Dim ilSequence As Integer

        On Error GoTo ErrHand
       
        ilSequence = 0
        If Not ilHistory Then           'not history, get all day events for the library
                                        'rather than just the change made
            'gather the events for this library
            ilRet = gGetRecs_DEE_DayEvent(smDEEStamp, tmDHE.lCode, "EngrLibEvts-cmdReport", tlDEE())
        End If
        For llDee = LBound(tlDEE) To UBound(tlDEE) - 1
            LSet tmDee = tlDEE(llDee)
                      
            rstLibEvts.AddNew
            ilSequence = ilSequence + 1

            'get the library name
            For ilDNE = 0 To UBound(tgCurrDNE) - 1 Step 1
                If tmDHE.lDneCode = tgCurrDNE(ilDNE).lCode Then
                    rstLibEvts.Fields("Name") = Trim$(tgCurrDNE(ilDNE).sName)
                    Exit For
                End If
            Next ilDNE
              
            'get the library subname
            rstLibEvts.Fields("SubName") = ""
            For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                If tmDHE.lDseCode = tgCurrDSE(ilDSE).lCode Then
                    rstLibEvts.Fields("SubName") = Trim$(tgCurrDSE(ilDSE).sName)
                    Exit For
                End If
            Next ilDSE
              
            rstLibEvts.Fields("DescCteCode") = tmDHE.lCteCode       'Library description
            rstLibEvts.Fields("StartDate") = Trim$(tmDHE.sStartDate)
            llDate = gDateValue(tmDHE.sStartDate)
            rstLibEvts.Fields("StartDateSort") = llDate     'date for sorting when same library names
            
            rstLibEvts.Fields("EndDate") = Trim$(tmDHE.sEndDate)
              
            rstLibEvts.Fields("StartTime") = Trim$(tmDHE.sStartTime)
            
            rstLibEvts.Fields("Length") = gLongToLength(tmDHE.lLength, True)
                             
            rstLibEvts.Fields("State") = tmDHE.sState
            
            '2-21-06 add ignore conflict flag from library header
            rstLibEvts.Fields("IgnoreConflicts") = tmDHE.sIgnoreConflicts
            rstLibEvts.Fields("DHEHeaderCode") = tmDHE.lCode
        
            'Determine event type ETE(program or avail)
            For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
                If tmDee.iEteCode = tgCurrETE(ilLoop).iCode Then
                    rstLibEvts.Fields("EventType") = Trim$(tgCurrETE(ilLoop).sCategory)  'p = pgm , a = avail
                    Exit For
                End If
            Next ilLoop
            
            'Setup reference for Bus references
            rstLibEvts.Fields("EvBusDeeCode") = tmDee.lCode
                      
            'event start time displacement
            rstLibEvts.Fields("EvStartTime") = gLongToStrLengthInTenth(tmDee.lTime, False)
            'event start time for sorting
            rstLibEvts.Fields("EvStartTimeSort") = tmDee.lTime
            
            'Event Type (TimeType TTE, type = S)
            rstLibEvts.Fields("EvStartType") = ""
            For ilLoop = 0 To UBound(tgCurrStartTTE) - 1 Step 1
                If tmDee.iStartTteCode = tgCurrStartTTE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvStarttype") = Trim$(tgCurrStartTTE(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            
            'Fixed Type (Y/N)
            rstLibEvts.Fields("EvFix") = tmDee.sFixedTime
            
            'Event End type (tte, type = "E"
            rstLibEvts.Fields("EvEndType") = ""
            For ilLoop = 0 To UBound(tgCurrEndTTE) - 1 Step 1
                If tmDee.iEndTteCode = tgCurrEndTTE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvEndType") = Trim$(tgCurrEndTTE(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            
            'Event duration
            rstLibEvts.Fields("EvDur") = gLongToStrLengthInTenth(tmDee.lDuration, True)
            
            'Days
            slStr = gDayMap(tmDee.sDays)
            rstLibEvts.Fields("EvDays") = Trim$(slStr)
            
            'Hours
            slStr = gHourMap(tmDee.sHours)
            rstLibEvts.Fields("EvHours") = Trim$(slStr)
            
            'Event Material Type (MTE)
            rstLibEvts.Fields("EvMatType") = ""
            For ilLoop = 0 To UBound(tgCurrMTE) - 1 Step 1
                If tmDee.iMteCode = tgCurrMTE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvMatType") = Trim$(tgCurrMTE(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            
            'primary audio source (ANE)
            rstLibEvts.Fields("EvAudName1") = ""
            'For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
            '    If tmDee.iAudioAseCode = tgCurrASE(ilLoop).iCode Then
                ilLoop = gBinarySearchASE(tmDee.iAudioAseCode, tgCurrASE())
                If ilLoop <> -1 Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tgCurrASE(ilLoop).iPriAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tgCurrASE(ilLoop).iPriAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            rstLibEvts.Fields("EvAudName1") = Trim$(tgCurrANE(ilANE).sName)
                    '        Exit For
                        End If
                    'Next ilANE
            '        Exit For
                End If
            'Next ilLoop
            
            'Primary Audio Source item
            rstLibEvts.Fields("EvItem1") = Trim$(tmDee.sAudioItemID)
            
            'primary audio source control CCE
            rstLibEvts.Fields("EvCtl1") = ""
            For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If tmDee.iAudioCceCode = tgCurrAudioCCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvCtl1") = Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
            'Backup audio source (ANE)
            rstLibEvts.Fields("EvAudName2") = ""
            'For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
            '    If tmDee.iAudioAseCode = tgCurrASE(ilLoop).iCode Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    'If tgCurrASE(ilLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
                    '    If tmDee.iBkupAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tmDee.iBkupAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            rstLibEvts.Fields("EvAudName2") = Trim$(tgCurrANE(ilANE).sName)
                    '        Exit For
                        End If
                    'Next ilANE
            '        Exit For
            '    End If
            'Next ilLoop
                             
            'backup audio source control CCE
            rstLibEvts.Fields("EvCtl2") = ""
            For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If tmDee.iBkupCceCode = tgCurrAudioCCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvCtl2") = Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop

            'protection audio source (ANE)
            rstLibEvts.Fields("EvAudName3") = ""
            'For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
            '    If tmDee.iProtAneCode = tgCurrASE(ilLoop).iCode Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    'If tgCurrASE(ilLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
                    '    If tmDee.iProtAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tmDee.iProtAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            rstLibEvts.Fields("EvAudName3") = Trim$(tgCurrANE(ilANE).sName)
                    '        Exit For
                        End If
                    'Next ilANE
            '        Exit For
            '    End If
            'Next ilLoop
            
            'Protection Audio Source item
            rstLibEvts.Fields("EvItem3") = Trim$(tmDee.sProtItemID)
            
            'Protection audio source control CCE
            rstLibEvts.Fields("EvCtl3") = ""
            For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                If tmDee.iProtCceCode = tgCurrAudioCCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvCtl3") = Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
            'Relay # 1 of 2(RNE)
            rstLibEvts.Fields("EvRelay1") = ""
            'For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If tmDee.i1RneCode = tgCurrRNE(ilLoop).iCode Then
                ilLoop = gBinarySearchRNE(tmDee.i1RneCode, tgCurrRNE())
                If ilLoop <> -1 Then
                    rstLibEvts.Fields("EvRelay1") = Trim$(tgCurrRNE(ilLoop).sName)
            '        Exit For
                End If
            'Next ilLoop
            
            'Relay # 2 of 2(RNE)
            rstLibEvts.Fields("EvRelay2") = ""
            'For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If tmDee.i2RneCode = tgCurrRNE(ilLoop).iCode Then
                ilLoop = gBinarySearchRNE(tmDee.i2RneCode, tgCurrRNE())
                If ilLoop <> -1 Then
                    rstLibEvts.Fields("EvRelay2") = Trim$(tgCurrRNE(ilLoop).sName)
            '        Exit For
                End If
            'Next ilLoop
            
            'Follow
            rstLibEvts.Fields("EvFollow") = ""
            For ilLoop = 0 To UBound(tgCurrFNE) - 1 Step 1
                If tmDee.iFneCode = tgCurrFNE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvFollow") = Trim$(tgCurrFNE(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            
            'Silence Time in tenths of seconds
            slStr = gLongToLength(tmDee.lSilenceTime, False)     'gLongToStrLengthInTenth(tmDee.lSilenceTime, False)
            rstLibEvts.Fields("EvTime") = Trim(slStr)
            'Silence code 1 or 4
            rstLibEvts.Fields("EvSilence1") = ""
            For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
                If tmDee.i1SceCode = tgCurrSCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvSilence1") = Trim$(tgCurrSCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
            'Silence code 2 or 4
            rstLibEvts.Fields("EvSilence2") = ""
            For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
                If tmDee.i2SceCode = tgCurrSCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvSilence2") = Trim$(tgCurrSCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
             'Silence code 3 or 4
            rstLibEvts.Fields("EvSilence3") = ""
            For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
                If tmDee.i3SceCode = tgCurrSCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvSilence3") = Trim$(tgCurrSCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
            'Silence code 4 or 4
            rstLibEvts.Fields("EvSilence4") = ""
            For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
                If tmDee.i4SceCode = tgCurrSCE(ilLoop).iCode Then
                    rstLibEvts.Fields("EvSilence4") = Trim$(tgCurrSCE(ilLoop).sAutoChar)
                    Exit For
                End If
            Next ilLoop
            
            'net cue 1 of 2
            rstLibEvts.Fields("EvNetCue1") = ""
            'For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If tmDee.iStartNneCode = tgCurrNNE(ilLoop).iCode Then
                ilLoop = gBinarySearchNNE(tmDee.iStartNneCode, tgCurrNNE())
                If ilLoop <> -1 Then
                    rstLibEvts.Fields("EvNetCue1") = Trim$(tgCurrNNE(ilLoop).sName)
            '        Exit For
                End If
            'Next ilLoop
            
            'net cue 2 of 2
            rstLibEvts.Fields("EvNetCue2") = ""
            'For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If tmDee.iStartNneCode = tgCurrNNE(ilLoop).iCode Then
                ilLoop = gBinarySearchNNE(tmDee.iEndNneCode, tgCurrNNE())
                If ilLoop <> -1 Then
                    rstLibEvts.Fields("EvNetCue2") = Trim$(tgCurrNNE(ilLoop).sName)
            '        Exit For
                End If
            'Next ilLoop
            
            'Title 1 of 2
            rstLibEvts.Fields("EvTitle1CteCode") = tmDee.l1CteCode
            
             'Title 2 of 2
            rstLibEvts.Fields("EvTitle2CteCode") = tmDee.l2CteCode
            
            rstLibEvts.Fields("EvABCCustomFields") = ""
            If sgClientFields = "A" Then            'abc client
                If Trim$(tmDee.sABCFormat) <> "" Or Trim$(tmDee.sABCPgmCode) <> "" Or Trim$(tmDee.sABCXDSMode) <> "" Or Trim$(tmDee.sABCRecordItem) <> "" Then
                    rstLibEvts.Fields("EvABCCustomFields") = "ABC: Format-" & Trim$(tmDee.sABCFormat) & ", Pgm Code-" & Trim$(tmDee.sABCPgmCode) & ", XDS Mode-" & Trim$(tmDee.sABCXDSMode) & ", Record Item-" & Trim$(tmDee.sABCRecordItem)
                End If
            End If
            
            rstLibEvts.Fields("Version") = ilVersion            'for sorting on history option
            rstLibEvts.Fields("SubVersion") = tmDHE.iVersion    'for sorting on history option
            rstLibEvts.Fields("Sequence") = ilSequence      'keep multiple lines of same library header together
            If ilHistory Then                   'append the rest of the fields for history
                rstLibEvts.Fields("DateEntered") = Format$(tmAIE.sEnteredDate, sgSQLDateForm)
                rstLibEvts.Fields("TimeEntered") = Format$(tmAIE.sEnteredTime, sgSQLTimeForm)
                rstLibEvts.Fields("OrigAIECode") = tmAIE.lCode
                rstLibEvts.Fields("FileCode") = tmAIE.lOrigFileCode
                
                For ilLoop = LBound(tgCurrUIE) To UBound(tgCurrUIE) - 1
                    If tgCurrUIE(ilLoop).iCode = tmAIE.iUieCode Then
                        rstLibEvts.Fields("User") = Trim$(tgCurrUIE(ilLoop).sShowName)
                        Exit For
                    End If
                Next ilLoop
            End If
        Next llDee
        Exit Sub
        
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in User Rpt-cmdReport:  mAddRstLibEvts "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in User Rpt-cmdReport: mAddRstLibEvts "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub
