VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmSendReport 
   Caption         =   "Crystal Reports"
   ClientHeight    =   4335
   ClientLeft      =   2055
   ClientTop       =   2355
   ClientWidth     =   6945
   Icon            =   "AffSendReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6945
   Begin VB.TextBox txtWaitForReport 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1905
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   885
      Left            =   1290
      TabIndex        =   0
      Top             =   2460
      Visible         =   0   'False
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   1561
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Left            =   5505
      Top             =   2850
   End
End
Attribute VB_Name = "frmSendReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private tmRQF As RQF
Private hmRQF As Integer
Private hmRFF As Integer
Dim imRecLen As Integer
Dim bmReportDone As Boolean
Const ERRORLOG = "AffErrorLog.txt"
'Const MYINTERFACEVERSION As String = "/Version1.0"
'Const DEBUGMODE = "/D"
Enum ReportDone
    NotDone
    WithErrors
    Success
End Enum
Private Sub Form_Load()
    Me.Left = -5000
    Me.Height = 5
    Me.Width = 5
    Me.Caption = "AffiliateForReport"
    Timer1.Interval = 5
    Timer1.Enabled = True   'mMakeNetReport in timer so can unload if need to.
    Screen.MousePointer = vbHourglass
End Sub
Public Sub mMakeNetReport()
    Dim slCommandLine As String
    Dim ilRet As Integer
    Dim ilCounter As Integer
    Dim ilReportResult As ReportDone   'batch
    'Dim slPathToExe As String
    Dim blRet As Boolean
On Error GoTo errorbox

    hmRQF = CBtrvTable(1)
    ilRet = btrOpen(hmRQF, "", sgDBPath & "rqf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> 0 Then
        gMsgBox "Couldn't open RQF Table", vbInformation + vbOKOnly, "Error"
        gLogMsg "Couldn't open RQF Table", ERRORLOG, False
    End If
    ilRet = mOpenRFFTable
    If ilRet <> 0 Then
        gMsgBox "Couldn't open RFF Table", vbInformation + vbOKOnly, "Error"
        gLogMsg "Couldn't open RFF Table", ERRORLOG, False
    End If
    'Dan M 3/23/10 Made global to open csiNetReporter
'    'Dan M 12/17/09  run csiNetReporterAlternate IF 'test' is in the folder name.
'    slCommandLine = mBuildAlternateAsNeeded
'    slPathToExe = slCommandLine
'    'slCommandLine = sgExeDirectory & "csinetreporter.exe "
'    'Dan M Debug only turned on for special internal guide. 12/10/09
'    If (Len(sgSpecialPassword) = 4) Then
'        slCommandLine = slCommandLine & DEBUGMODE & " "
'    End If
'    'slCommandLine = "csinetreporter.exe \D " & MYINTERFACEVERSION & " "
'    slCommandLine = slCommandLine & MYINTERFACEVERSION & " "
'    If LenB(sgStartupDirectory) Then
'        slCommandLine = slCommandLine & " """ & sgStartupDirectory & """ "
'    End If
    If ogReport.OkToMoveToFirstReport(True) Then
        ogReport.Reports.MoveFirst
        ilCounter = 0
        Do While Not ogReport.Reports.EOF
            mLoadRQFTableBasic Trim(ogReport.Reports!Name)
            mLoadRQFWithOutput
            tmRQF.lEnteredTime = tmRQF.lEnteredTime + ilCounter
            mRecordToTables ogReport.Reports!code
            slCommandLine = slCommandLine & " " & tmRQF.lCode
            ilCounter = ilCounter + 1
            ogReport.Reports.MoveNext
        Loop
        'Dan M 3/23/10 made global and added setting ilreportresult
       ' gShellAndWait slCommandLine
        ilReportResult = NotDone
        blRet = gCallNetReporter(Normal, slCommandLine)
        If Not blRet Then
            gMsgBox "Cannot find Report module with " & slCommandLine, vbOKOnly + vbExclamation, "Error"
            bgReportModuleRunning = False
            Exit Sub
        End If

'On Error GoTo errornoexe
'        Shell slCommandLine 'batch
'        bgReportModuleRunning = True
On Error GoTo errorbox
        Do Until ilReportResult <> NotDone 'batch
            Sleep 5000  '10000
            Refresh
            DoEvents
            ilReportResult = mReportDone
        Loop
            ' if report done successfully send a msgbox to user saying successful.
        If ilReportResult = Success Then 'batch
        ' If mReportDone = Success Then
            mMsgBoxDisplayed
        End If
    Else
        MsgBox "No data exists for that report", vbExclamation + vbOKOnly
    End If
   ' Screen.MousePointer = vbDefault
    Exit Sub
'errornoexe:
'    gMsgBox "Cannot find Report module at " & slPathToExe, vbOKOnly + vbExclamation, "Error"
'    bgReportModuleRunning = False
'    Exit Sub
errorbox:
    gMsgBox "mMakeNetReport had an error: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    gLogMsg "mMakeNetReport had an error: " & Err.Description, ERRORLOG, False
End Sub
Private Function mBuildAlternateAsNeeded() As String
    Dim ilRet As Integer
    Dim slStr As String
Dim slTempPath As String
If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
    slTempPath = sgExeDirectory & "csinetreporteralternate.exe "
Else
    slTempPath = sgExeDirectory & "csinetreporter.exe "
End If
    ilRet = 0
    On Error GoTo FileErr
    slStr = FileDateTime(slTempPath)
    On Error GoTo 0
    If ilRet = 1 Then
        If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
            slTempPath = "csinetreporteralternate.exe "
        Else
            slTempPath = "csinetreporter.exe "
        End If
    End If
mBuildAlternateAsNeeded = slTempPath
    Exit Function
FileErr:
    ilRet = 1
    Resume Next
End Function
Private Sub mMsgBoxDisplayed()
    Select Case tmRQF.sOutputType
'        Case "S"
'            MsgBox "Export Complete"
        Case "P"
            MsgBox "Printing Complete"
    End Select
End Sub
Private Function mReportDone() As ReportDone
    Dim ilRecLen As Integer
    Dim ilRet As Integer
    
    ilRecLen = Len(tmRQF)
    ilRet = btrGetEqual(hmRQF, tmRQF, ilRecLen, tmRQF.lCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If tmRQF.sCompleted = "Y" Then
        mReportDone = Success
    ElseIf tmRQF.sCompleted = "E" Then
    'this value is set by csiNetReporter2 if an error happened.  Currently, traffic/affiliate use shell and wait and don't do anything with this value.
        mReportDone = WithErrors
    ElseIf bmReportDone Or InStr(1, txtWaitForReport.Text, "Q", vbTextCompare) > 0 Then
    'Dan 12/18/09 csiNetReporter can call and change textbox if it is forced to quit.
        mReportDone = WithErrors
    Else
        mReportDone = NotDone
    End If

'    Dim ilRecLen As Integer
'    Dim ilRet As Integer
'
'    ilRecLen = Len(tmRQF)
'    ilRet = btrGetEqual(hmRQF, tmRQF, ilRecLen, tmRQF.lCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'    If tmRQF.sCompleted = "Y" Then
'        mReportDone = Success
'    ElseIf tmRQF.sCompleted = "E" Then
'    'this value is set by  csiNetReporter2 if an error happened.  Currently, traffic/affiliate use shell and wait and don't do anything with this value.
'        mReportDone = WithErrors
'    Else
'        mReportDone = NotDone
'    End If
End Function


Private Sub mLoadRQFTableBasic(slReportName As String)
    Dim slDate As String
    Dim slTime As String
    'dan 1/21/10 gnow may backdate(as in demo database) and don't want that
    'slDate = Format$(gNow(), "m/d/yy")
    slDate = Format$(Now(), "m/d/yy")

    tmRQF.lEnteredDate = gDateValue(slDate)
   ' slTime = Format$(gNow(), "h:mm:ssAM/PM")
    slTime = Format$(Now(), "h:mm:ssAM/PM")
    tmRQF.lEnteredTime = gTimeToLong(slTime, False)
    
    With tmRQF
        .sReportName = slReportName
        .sReportType = "A"
        .sRunMode = "C"
        .lCode = 0
        .sCompleted = "N"
        .lConnection = 1
        If Not ogReport Is Nothing Then
            If ogReport.Connect = ADO Then
                .lConnection = 2
            ElseIf ogReport.Connect = Native Then
                .lConnection = 0
            End If
        End If
        .sUserName = sgUserName
      '7-29-09 unused
        .sPriority = "N"
        .sRunType = "N"
        .sReportSource = "N" 'prepass?
        .sDisposition = "E"
        .sLastDateRun = "0"
        .iRunDay = 0
        .lPrePassDate = 0
        .lPrePassTime = 0
        .sUnused = "0"
    End With

End Sub
Private Sub mLoadRQFWithOutput()
    Select Case ogReport.OutputType
        Case Printerfile
            tmRQF.sOutputType = "P"
            tmRQF.iPrintCopies = ogReport.PrintCopies
        Case exportfile
            tmRQF.sOutputType = "S"
            tmRQF.sOutputFileName = ogReport.OutputFileName
            tmRQF.iOutputSaveType = ogReport.OutputSaveType
        Case Else
            tmRQF.sOutputType = "D"
    End Select
End Sub
Private Sub mRecordToRQFTable()
    Dim ilRecLen As Integer
    Dim ilRet As Integer
    
    ilRecLen = Len(tmRQF)
    ilRet = btrInsert(hmRQF, tmRQF, ilRecLen, 0)
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    If ilRet <> 0 Then
        Err.Raise 2805, "RecordToRQFTable", "Couldn't write to RQF table, btrieve error " & ilRet
    End If
End Sub
Private Function mOpenRFFTable() As Integer
    Dim ilRet As Integer
    
    hmRFF = CBtrvTable(1)
    ilRet = btrOpen(hmRFF, "", sgDBPath & "rff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    mOpenRFFTable = ilRet
End Function
Private Sub mRecordToRFFTable(Index As Integer, RqfKey As Long, Optional blExtended As Boolean = False)
    Dim ilRet As Integer
    
    If blExtended Then
        tgRffExtended(Index).lRqfCode = RqfKey
        ilRet = btrInsert(hmRFF, tgRffExtended(Index), imRecLen, 0)
    Else
        tgRff(Index).lRqfCode = RqfKey
        ilRet = btrInsert(hmRFF, tgRff(Index), imRecLen, 0)
    End If
    If ilRet <> 0 Then
        Err.Raise2807 , "Record to RFF table", "Couldn't write to rff table."
    End If

End Sub
Private Sub mRecordToTables(ByVal ilTempCode As Integer)
    Dim c As Integer
    Dim ilNumberOfRffRecords As Integer
    
    mRecordToRQFTable
    imRecLen = Len(tgRff(0))
    mCleanRff tmRQF.lCode
    'tgrff = 0 when nothing has been written to it.
    If UBound(tgRff) <> 0 Then
        ilNumberOfRffRecords = UBound(tgRff) - 1
        For c = 0 To ilNumberOfRffRecords 'UBound(tgrff) - 1
            If tgRff(c).lRqfCode = ilTempCode And tgRff(c).lCode = 0 Then
                mRecordToRFFTable c, tmRQF.lCode
                If tgRff(c).lExtendExists = 1 Then       'value too long for table?
                    mRffExtendForeignKey c
                End If
            End If
        Next c
        If UBound(tgRffExtended) > 0 Then
            For c = 0 To UBound(tgRffExtended) - 1
                tgRffExtended(c).iSequenceNumber = tgRffExtended(c).iSequenceNumber + ilNumberOfRffRecords
                mRecordToRFFTable c, tmRQF.lCode, True
            Next c
        End If
    End If
    ' the real code replaces the temporary code for deletion purposes.
    ogReport.Reports!code = tmRQF.lCode
End Sub
Private Sub mCleanRff(llRqfCode As Long)
    Dim tlRffKey1 As RFFKEY1
    Dim ilRet As Integer
    Dim tlRff As RFF
    Dim ilRffLen As Integer
    Dim olMyFileSystem As FileSystemObject
    
    ilRffLen = Len(tlRff)
    tlRffKey1.lRqfCode = llRqfCode
    tlRffKey1.iSequenceNumber = 0
    tlRffKey1.sType = "A"
    ilRet = btrGetGreaterOrEqual(hmRFF, tlRff, ilRffLen, tlRffKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While ilRet = BTRV_ERR_NONE And tlRff.lRqfCode = llRqfCode
        If tlRff.sType = "A" Then 'remove recordset
            Set olMyFileSystem = New FileSystemObject
            If olMyFileSystem.FileExists(tlRff.sFormulaValue) Then
                olMyFileSystem.DeleteFile tlRff.sFormulaValue
            End If
        End If
        ilRet = btrDelete(hmRFF)
        ilRet = btrGetNext(hmRFF, tlRff, ilRffLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop

End Sub
Private Sub mRffExtendForeignKey(ilForeignIndex As Integer)
    Dim ilForeignKey As Integer
    Dim c As Integer
    
    ilForeignKey = tgRff(ilForeignIndex).lCode
    If UBound(tgRffExtended) > 0 Then
        For c = 0 To UBound(tgRffExtended) - 1
            If tgRffExtended(c).lRffCode = ilForeignIndex Then
                tgRffExtended(c).lRffCode = ilForeignKey
            End If
        Next c
    Else
        gMsgBox "Problem in Report, mRffExtendForeignKey:  can't find rest of formula", vbOKOnly + vbExclamation, "Error"
        gLogMsg "Problem in Report, mRffExtendForeignKey:  can't find rest of formula", ERRORLOG, False
    End If
End Sub
Private Sub mCloseTables()
    Dim ilRet As Integer
    
    ilRet = btrClose(hmRQF)
    btrDestroy hmRQF
    ilRet = btrClose(hmRFF)
    btrDestroy hmRFF

End Sub
Private Sub mDeleteRecords()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim c As Integer
    
    ilRecLen = Len(tmRQF)
   ' If ogReport.Reports.EOF <> ogReport.Reports.BOF Then
    If ogReport.OkToMoveToFirstReport(True) Then
        ogReport.Reports.MoveFirst
        Do While Not ogReport.Reports.EOF
            tmRQF.lCode = ogReport.Reports!code
            ilRet = btrGetEqual(hmRQF, tmRQF, ilRecLen, tmRQF.lCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            ilRet = btrDelete(hmRQF)
            If ilRet <> 0 Then
               ' MsgBox "Couldn't delete RQF records in report.frm", vbInformation + vbOKOnly
            End If
            mCleanRff ogReport.Reports!code
            ogReport.Reports.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
    mDeleteRecords
    mCloseTables
    Erase tgRffExtended
    Erase tgRff
    If Not ogReport Is Nothing Then
        Set ogReport = Nothing
    End If
    Set frmSendReport = Nothing
End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    mMakeNetReport
    Unload Me

End Sub




Private Sub txtWaitForReport_Change()
bmReportDone = True
End Sub
