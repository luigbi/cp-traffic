VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmWebLogImportRpt 
   Caption         =   "Web Log Import"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffWebLogImportRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   7125
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3510
      Top             =   1080
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5385
      FormDesignWidth =   7125
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   4
      Top             =   255
      Width           =   1935
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
      Height          =   3030
      Left            =   240
      TabIndex        =   1
      Top             =   1785
      Width           =   6705
      Begin V81Affiliate.CSI_Calendar CalFromDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
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
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalToDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
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
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.Label LabTo 
         Caption         =   "End:"
         Height          =   240
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   570
      End
      Begin VB.Label labFrom 
         Caption         =   "Date- Start:"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      Top             =   720
      Width           =   1935
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
         ItemData        =   "AffWebLogImportRpt.frx":08CA
         Left            =   1050
         List            =   "AffWebLogImportRpt.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmWebLogImportRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hmFrom As Integer

Private Type IMPORTLOGDATA
    lSeq As Long
    sGenDate As String
    sGenTime As String
    sImportDateTime As String
    sImportDateForSort As String
    sUser As String
    sVehicle As String
    sStation As String
    sAdvt As String
    sProd As String
    sDateAired As String
    sTimeAired As String
    sLen As String
    sISCI As String
    sMsg As String
    sATTCode As String
    sASTCode As String
End Type

Private tmImportLogData As IMPORTLOGDATA

Private Const IMPORTLOGDATE = 1     'import log, position to date of export
Private Const IMPORTLOGTIME = 12    'import log, position to time of export
Private Const IMPORTLOGUSER = 29    'import log, position to user
'constant index to the spot data
Private Const ATTCODEINDEX = 1
Private Const VEHICLEINDEX = 2  '1
Private Const STATIONINDEX = 3  '2
Private Const ADVTINDEX = 4
Private Const PRODINDEX = 5
Private Const ASTCODEINDEX = 14
Private Const DATEAIREDINDEX = 15
Private Const TIMEAIREDINDEX = 16
Private Const STATUSINDEX = 17
Private Const LENINDEX = 10
Private Const ISCIINDEX = 12
Private Const MSGINDEX = 22  '21
Private Sub cmdReport_Click()

    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim slRptName As String         'full report name of crystal .rpt
    Dim slFromFile As String
    Dim llFromDate As Long
    Dim sFromDate As String
    Dim llToDate As Long
    Dim sToDate As String
    Dim slGenDate As String
    Dim slGenTime As String
    Dim ilOk As Integer
    Dim dFWeek As Date

    On Error GoTo ErrHand
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    gUserActivityLog "S", sgReportListName & ": Prepass"
    'guide gets to see key codes
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        sgCrystlFormula1 = 1
    Else
        sgCrystlFormula1 = 0
    End If
    slRptName = "AfWebImportLog.rpt"
    slExportName = "AfWebImportLog"
    
    
    'gUserActivityLog "E", sgReportListName & ": Prepass"
    
    slGenDate = Format$(gNow(), "m/d/yyyy")
    slGenTime = Format$(gNow(), sgShowTimeWSecForm)

    If CalFromDate.Text = "" Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFromDate.SetFocus
        Exit Sub
    End If
    
    sFromDate = CalFromDate.Text
    
    llFromDate = DateValue(gAdjYear(sFromDate))
    sToDate = CalToDate.Text
    If Trim$(sToDate) = "" Then         'no end date enterd, make same as from date
        sToDate = sFromDate
    End If
    llToDate = DateValue(gAdjYear(sToDate))
    
    If gIsDate(sFromDate) = False Or (Len(Trim$(sFromDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFromDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sToDate) = False Or (Len(Trim$(sToDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalToDate.SetFocus
        Exit Sub
    End If
    
    dFWeek = CDate(sFromDate)
    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
 
    dFWeek = CDate(sToDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula3 = False            'assume txt file exists
    slFromFile = sgDBPath & "Messages\WebImportLog.Txt"
    
    
    Screen.MousePointer = vbHourglass
    ilOk = mReadFile(slFromFile, llFromDate, llToDate, slGenDate, slGenTime)
    If Not ilOk Then
        'error with web import log file, does not exist in message folder
        gMsg = "WebImportLog.Txt does not exist in " & sgDBPath & "Messages folder"
        gMsgBox gMsg, vbCritical
        sgCrystlFormula3 = True         'file does not exist
        Screen.MousePointer = vbDefault
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
        
    SQLQuery = "Select * from smr "
    SQLQuery = SQLQuery + " where ( smrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"

    'dan todo change for rollback
    'frmCrystal.gCrystlReports "", ilExportType, ilRptDest, slRptName, slExportName, True
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName

    'remove all the records just printed
    SQLQuery = "DELETE FROM smr "
    SQLQuery = SQLQuery & " WHERE (smrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' " & "and smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "frmWebLogImportRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebLogImportRpt-cmdReport"
End Sub

Private Sub cmdDone_Click()
    Unload frmWebLogImportRpt

End Sub


Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmWebLogImportRpt
    
End Sub
Private Sub Form_Load()

        
    frmWebLogImportRpt.Caption = "Web Import Log Report - " & sgClientName

End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmWebLogImportRpt
    gCenterForm frmWebLogImportRpt
    gPopExportTypes cboFileType
    cboFileType.Enabled = True

End Sub
Private Sub Form_Initialize()
    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmWebLogImportRpt = Nothing

End Sub

'****************************************************************
'*                                                              *
'*      Procedure Name:mReadFile                                *
'*      <input>   slFromFile - full path and file name          *
'                 of web import filename (webimportlog.txt      *
'                 llFromDate - user requested start date        *
'                 llToDAte - user requested end date            *
'*                                                              *
'*                                                              *
'*                                                              *
'****************************************************************
Private Function mReadFile(slFromFile As String, llFromDate As Long, llToDate As Long, slGenDate As String, slGenTime As String) As Integer
    
    Dim ilRet As Integer
    Dim slLine As String
    Dim slStr As String
    Dim ilEof As Integer
    Dim slDate As String
    Dim ilDatesToProcess() As Integer
    Dim ilHowManyDates As Integer
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim slTempMsg As String
    Dim ilStartPos As Integer
    Dim ilErrorMsgPos As Integer
    Dim llline As Long
    
    'ilRet = 0
    On Error GoTo mReadFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        mReadFile = False
        Exit Function
    End If
        
    ilHowManyDates = (llToDate - llFromDate) + 1
    ReDim ilDatesToProcess(0 To ilHowManyDates) As Integer
     
    tmImportLogData.sGenDate = slGenDate
    tmImportLogData.sGenTime = slGenTime
    tmImportLogData.lSeq = 0
    Do
        ilRet = 0
        'On Error GoTo mReadFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                'see if this is a valid entry based on date
                'positions 1 -10 is date:  mm-dd-yyyy
                slDate = Mid$(slLine, IMPORTLOGDATE, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 3, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 6, 4)
                'test for valid date
                If gIsDate(slDate) = True Then          'must be a record that doesnt start with a date in position 1 (vb messages)
                    'tmImportLogData.sImportDateTime = ""
                    tmImportLogData.sUser = ""
                    tmImportLogData.sLen = ""
                    tmImportLogData.sVehicle = ""
                    tmImportLogData.sStation = ""
                    tmImportLogData.sAdvt = ""
                    tmImportLogData.sProd = ""
                    tmImportLogData.sDateAired = ""
                    tmImportLogData.sTimeAired = ""
                    tmImportLogData.sISCI = ""
                   
                    slTempMsg = ""
                    llDate = gDateValue(slDate)
                    If llDate >= llFromDate And llDate <= llToDate Then                 'passed the user requested date filter
                        ilIndex = llDate - llFromDate
                        '
                        'There are 3 main types of records to look for:  The starting Export information record, Any Spot Warning messages, and hard errors that prevent export from running
                        '
                        If (InStr(1, slLine, "Starting", vbBinaryCompare) > 0) Or (InStr(1, slLine, "Importing spots.", vbBinaryCompare) > 0) Then          'found starting event for the day
                            ilDatesToProcess(ilIndex) = True
                            'format the start of export time to mm/dd/yy (vs mm-dd-yyyy)
                            tmImportLogData.sImportDateTime = Mid$(slLine, IMPORTLOGDATE, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 3, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 8, 2) & " at " & Mid$(slLine, IMPORTLOGTIME, 10)                'start date and time of import
                            ilStartPos = InStr(IMPORTLOGUSER, slLine, " ", vbBinaryCompare)        'find blank after text "User:"
                            If ilStartPos > 0 Then
                                tmImportLogData.sUser = Mid$(slLine, IMPORTLOGUSER, (ilStartPos - IMPORTLOGUSER))         'get the user name
                            Else
                                tmImportLogData.sUser = "Unknown"
                            End If
                            tmImportLogData.sMsg = "** Starting Web Import Process **"
                            tmImportLogData.sImportDateForSort = Trim$(Str(llDate))
                            Do While Len(tmImportLogData.sImportDateForSort) < 5
                                tmImportLogData.sImportDateForSort = "0" & tmImportLogData.sImportDateForSort
                            Loop
                            'mWriteImportRecord
                        ElseIf InStr(1, slLine, "Warning:", vbBinaryCompare) > 0 Then       'probably ast no match
                            'tmImportLogData.sImportDateTime = Mid$(slLine, IMPORTLOGDATE, 10) & " at " & Mid$(slLine, IMPORTLOGTIME, 10)          'start date and time of import
                            ilStartPos = InStr(IMPORTLOGUSER, slLine, " ", vbBinaryCompare)         'get past "User:" text
                            
                            If ilStartPos > 0 Then
                                tmImportLogData.sUser = Mid$(slLine, IMPORTLOGUSER, (ilStartPos - IMPORTLOGUSER))     'get user name
                            Else
                                tmImportLogData.sUser = "Unknown"
                            End If
                            ilStartPos = InStr(IMPORTLOGUSER, slLine, "Warning:", vbBinaryCompare)     'find blank after the username, then comes the message.  Determine of spot warning message
                            If ilStartPos > 0 Then
                                slTempMsg = Mid$(slLine, ilStartPos + 9)
                            End If
                            
                            ilRet = gParseItem(slTempMsg, ATTCODEINDEX, ",", tmImportLogData.sATTCode)
                            ilRet = gParseItem(slTempMsg, VEHICLEINDEX, ",", tmImportLogData.sVehicle)
                            ilRet = gParseItem(slTempMsg, STATIONINDEX, ",", tmImportLogData.sStation)
                            ilRet = gParseItem(slTempMsg, ADVTINDEX, ",", tmImportLogData.sAdvt)
                            ilRet = gParseItem(slTempMsg, PRODINDEX, ",", tmImportLogData.sProd)
                            ilRet = gParseItem(slTempMsg, ASTCODEINDEX, ",", tmImportLogData.sASTCode)
                            ilRet = gParseItem(slTempMsg, DATEAIREDINDEX, ",", slStr)                   'date is format yyyy-mm-dd
                            tmImportLogData.sDateAired = Mid$(slStr, 6, 2) & "/" & Mid$(slStr, 9, 2) & "/" & Mid$(slStr, 3, 2)
                            ilRet = gParseItem(slTempMsg, TIMEAIREDINDEX, ",", tmImportLogData.sTimeAired)
                            ilRet = gParseItem(slTempMsg, LENINDEX, ",", tmImportLogData.sLen)
                            ilRet = gParseItem(slTempMsg, ISCIINDEX, ",", tmImportLogData.sISCI)
                            ilRet = gParseItem(slTempMsg, MSGINDEX, ",", tmImportLogData.sMsg)
                            ilRet = gParseItem(slTempMsg, STATUSINDEX, ",", slStr)
                            If Val(slStr) = 1 Or Val(slStr) = 4 Or Val(slStr) = 5 Then      'not aired
                                'blank out the actual date & time
                                tmImportLogData.sDateAired = ""
                                tmImportLogData.sTimeAired = ""
                            End If
                            mWriteImportRecord
                        ElseIf InStr(1, slLine, "Error", vbBinaryCompare) > 0 Or InStr(1, slLine, "error", vbBinaryCompare) > 0 Then
                            'tmImportLogData.sImportDateTime = Mid$(slLine, IMPORTLOGDATE, 10) & " at " & Mid$(slLine, IMPORTLOGTIME, 10)          'start date and time of import
                            'format the start of export time to mm/dd/yy (vs mm-dd-yyyy)
                            tmImportLogData.sImportDateTime = Mid$(slLine, IMPORTLOGDATE, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 3, 2) & "/" & Mid$(slLine, IMPORTLOGDATE + 8, 2) & " at " & Mid$(slLine, IMPORTLOGTIME, 10)                'start date and time of import
                            ilStartPos = InStr(IMPORTLOGUSER, slLine, " ", vbBinaryCompare)         'get past "User:" text
                            ilErrorMsgPos = ilStartPos                                              'start of the error text, just past the username
                            If ilStartPos > 0 Then
                                tmImportLogData.sUser = Mid$(slLine, IMPORTLOGUSER, (ilStartPos - IMPORTLOGUSER))     'get user name
                            Else
                                tmImportLogData.sUser = "Unknown"
                            End If
                            ilStartPos = InStr(IMPORTLOGUSER, slLine, "Error:", vbBinaryCompare)     'look for any hard errors that occurred
                            If ilStartPos > 0 Then
                                tmImportLogData.sMsg = "Call CSI," & Mid$(slLine, ilErrorMsgPos)
                                slStr = Mid$(slLine, ilErrorMsgPos)         'line contains the text "error".  Look for ":" representing the delimeter used to end the message to show
                                    ilStartPos = InStr(1, slStr, ".", vbBinaryCompare)      'get the entire string containing the error message
                                    If ilStartPos > 0 Then
                                        tmImportLogData.sMsg = "Call CSI," & Mid$(slStr, 1, ilStartPos)
                                    'Else
                                     '   tmImportLogData.sMsg = "Call CSI, General error has occured"
                                    End If
                            Else
                                ilStartPos = InStr(IMPORTLOGUSER, slLine, "error", vbBinaryCompare)     'look for any hard errors that occurred
                                If ilStartPos > 0 Then
                                    slStr = Mid$(slLine, ilErrorMsgPos)         'line contains the text "error".  Look for ":" representing the delimeter used to end the message to show
                                    ilStartPos = InStr(ilErrorMsgPos, slStr, ":", vbBinaryCompare)      'get the entire string containing the error message
                                    If ilStartPos > 0 Then
                                        tmImportLogData.sMsg = "Call CSI," & Mid$(slStr, 1, ilStartPos - 1)
                                    Else
                                        tmImportLogData.sMsg = "Call CSI, General error has occured"
                                    End If
                                Else
                                    tmImportLogData.sMsg = Mid$(slLine, ilErrorMsgPos)
                                End If
                              
                            End If
                            mWriteImportRecord
                        End If      'instr
                    End If          'lldate >= llfromdate and llDate <= lltodate
                End If              'gisdate(sldate)
            End If                  'asc(slline) = 26            'eof
        End If                      '(len(slline) > 0
    Loop Until ilEof
    Close hmFrom
    If ilRet <> 0 Then
        mReadFile = False
    Else
        mReadFile = True
    End If
    MousePointer = vbDefault
    Exit Function
mReadFileErr:
    ilRet = Err.Number
    Resume Next
End Function
'       mWRiteImportRecord
'       Create the prepass record from what was parsed from the Web Import Log
'       to send to crystal
'       Common parameters to update are in tmImportLogData
'    sImportDateTime As String
'    sUser As String
'    sVehicle As String
'    sStation As String
'    sAdvt As String
'    sProd As String
'    sDateAired As String
'    sTimeAired As String
'    sLen As String
'    sISCI As String
'    sMsg As String
Private Sub mWriteImportRecord()
Dim SQLQuery As String

        'remove illegal characters from description fields
        tmImportLogData.sVehicle = gRemoveIllegalChars(tmImportLogData.sVehicle)
        tmImportLogData.sUser = gRemoveIllegalChars(tmImportLogData.sUser)
        tmImportLogData.sAdvt = gRemoveIllegalChars(tmImportLogData.sAdvt)
        tmImportLogData.sProd = gRemoveIllegalChars(tmImportLogData.sProd)
        tmImportLogData.sISCI = gRemoveIllegalChars(tmImportLogData.sISCI)
        tmImportLogData.sMsg = gRemoveIllegalChars(tmImportLogData.sMsg)
        
        tmImportLogData.sVehicle = gFixQuote(tmImportLogData.sVehicle)
        tmImportLogData.sUser = gFixQuote(tmImportLogData.sUser)
        tmImportLogData.sAdvt = gFixQuote(tmImportLogData.sAdvt)
        tmImportLogData.sProd = gFixQuote(tmImportLogData.sProd)
        tmImportLogData.sISCI = gFixQuote(tmImportLogData.sISCI)
        tmImportLogData.sMsg = gFixQuote(tmImportLogData.sMsg)

        tmImportLogData.lSeq = tmImportLogData.lSeq + 1
        If InStr(tmImportLogData.sMsg, "Agreement") > 0 Then            '5-22-12 show the agreement code if anymessage dealing with agreements
            tmImportLogData.sMsg = tmImportLogData.sMsg & "," & tmImportLogData.sATTCode
        End If
        If InStr(tmImportLogData.sMsg, "AST ") > 0 Then            '5-22-12 show the AST code if anymessage dealing with AST
            tmImportLogData.sMsg = tmImportLogData.sMsg & "," & tmImportLogData.sASTCode
        End If
        
        SQLQuery = "INSERT INTO smr (smrGenDate, smrGenTime, "      'gen date & time
        SQLQuery = SQLQuery & "smrMSARank, "                        'import date as decimal string for sorting
        SQLQuery = SQLQuery & "smrCallLetters, "                    'Call letters
        SQLQuery = SQLQuery & "smrDMAMarket, "                      'Date Aired
        SQLQuery = SQLQuery & "smrMSAMarket, "                      'Time aired
        SQLQuery = SQLQuery & "smrState, "                          'spot length
        SQLQuery = SQLQuery & "smrFormat, "                         'Date & Time Export started
        SQLQuery = SQLQuery & "smrOwner, "                          'user name
        SQLQuery = SQLQuery & "smrString1, "                        'vehicle
        SQLQuery = SQLQuery & "smrString2, "                        'advt
        SQLQuery = SQLQuery & "smrString3, "                        'prod
        SQLQuery = SQLQuery & "smrString4, "                        'isci
        SQLQuery = SQLQuery & "smrString5, "                        'message
        SQLQuery = SQLQuery & "smrSeqNo) "                          'seq no
       
        SQLQuery = SQLQuery & " Values ( '" & Format$(tmImportLogData.sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(tmImportLogData.sGenTime, False))))) & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sImportDateForSort & "', "
        SQLQuery = SQLQuery & "'" & tmImportLogData.sStation & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sDateAired & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sTimeAired & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sLen & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sImportDateTime & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sUser & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sVehicle & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sAdvt & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sProd & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sISCI & "',"
        SQLQuery = SQLQuery & "'" & tmImportLogData.sMsg & "',"
        SQLQuery = SQLQuery & tmImportLogData.lSeq
        SQLQuery = SQLQuery & ")"

        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmWebLogImportRpt-WriteImportRecord"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebLogImportRpt-mWriteImportRecord"
    Exit Sub
End Sub

