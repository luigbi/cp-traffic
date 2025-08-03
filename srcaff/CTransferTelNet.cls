VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransferTelNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'use this to model other processes
'change loadIni, connect disconnect, mProcessmain
Implements ITransfer
Private Const MYCHOICE As String = "Telnet"
Private Const DEFAULTDAYS As Integer = 30
Private Const DEFAULTFOLDER As String = "Transferred"
Private Const SHOWDAYS As Integer = -2
Private Const KICKERLOGIN As String = "LOGIN"
Private Const KICKERPASSWORD As String = "PASSWORD"
Private Const KICKERCONFIRMDEFAULT As String = ".compel.sys"
Private Const KICKERERRORDEFAULT As String = "bash:"
Private smError As String
Private smAddress As String
Private imPort As Integer
Private smUsername As String
Private smPassword As String
Private smFromPath As String
Private smSavePath As String
Private bmIsLoaded As Boolean
Private ilDays As Integer
Private smName As String
Private bmNoPassword As Boolean
Private smExt As String
Private myFiles As ADODB.Recordset
Private smStatus As String
Private bmConnected As Boolean
Private bmStartPaused As Boolean
Private omFile As FileSystemObject
Private smEarlyDate As String
'6239
Private smWrapStart As String
Private smWrapEnd As String
Private smKickerConfirmation As String
Private smKickerError As String

Private Sub Class_Initialize()
    Set myFiles = gPrepRecordset
    bmConnected = False
    bmStartPaused = False
    Set omFile = New FileSystemObject
    smEarlyDate = NODATE
End Sub

Private Sub Class_Terminate()
    If Not myFiles Is Nothing Then
        If (myFiles.State And adStateOpen) <> 0 Then
            myFiles.Close
        End If
        Set myFiles = Nothing
    End If
    Set omFile = Nothing
End Sub

Private Property Let ITransfer_DateStored(slToday As String)
    smEarlyDate = DateAdd("d", SHOWDAYS, slToday)
End Property
Private Property Get ITransfer_DateStored() As String
    ITransfer_DateStored = DateAdd("d", -SHOWDAYS, smEarlyDate)
End Property
Private Property Get ITransfer_NextFile() As String
    Dim slRet As String
    
On Error GoTo ERRORBOX
    slRet = ""
    If Not myFiles Is Nothing Then
        myFiles.Filter = "isProcessed = False"
        If Not myFiles.EOF Then
            slRet = Trim$(myFiles!Filename)
        End If
    End If
Cleanup:
    ITransfer_NextFile = slRet
    Exit Property
ERRORBOX:
    slRet = ""
    GoTo Cleanup
End Property
Private Property Get ITransfer_StartPaused() As Boolean
    ITransfer_StartPaused = bmStartPaused
End Property
Private Property Get ITransfer_StatusMessage() As String
    ITransfer_StatusMessage = smStatus
End Property
Private Property Get ITransfer_ExtensionToFind() As String
    ITransfer_ExtensionToFind = smExt
End Property
Private Property Let ITransfer_ExtensionToFind(slExt As String)
    If Mid(slExt, 1, 1) <> "." Then
        slExt = "." & slExt
    End If
    smExt = slExt
End Property
Private Property Let ITransfer_IsNotPasswordProtected(blValue As Boolean)
    bmNoPassword = blValue
End Property
Private Property Get ITransfer_IsNotPasswordProtected() As Boolean
    ITransfer_IsNotPasswordProtected = bmNoPassword
End Property
Private Property Let ITransfer_Name(slValue As String)
    smName = slValue
End Property
Private Property Get ITransfer_Name() As String
    ITransfer_Name = smName
End Property
Private Property Get ITransfer_Address() As String
    ITransfer_Address = smAddress
End Property
Private Property Get ITransfer_DaysToSave() As Integer
    ITransfer_DaysToSave = ilDays
End Property
Private Property Get ITransfer_ErrorMessage() As String
    ITransfer_ErrorMessage = smError
End Property
'Private Property Get ITransfer_ExportFromPath() As String
'    ITransfer_ExportFromPath = smFromPath
'End Property
Private Property Get ITransfer_SaveFolder() As String
    ITransfer_SaveFolder = smSavePath
End Property
Private Property Get ITransfer_IsLoaded() As Boolean
    ITransfer_IsLoaded = bmIsLoaded
End Property
'Private Property Get ITransfer_Mode() As TransferMode
'    ITransfer_Mode = emMode
'End Property

Private Property Get ITransfer_FilesProcessed(blRefresh As Boolean) As Integer
    Dim ilCount As Integer
    Dim slFileName As String
    Dim slDate As String
    
    ilCount = 0
    If blRefresh Then
        smError = ""
        'start fresh
        If Not myFiles Is Nothing Then
            If (myFiles.State And adStateOpen) <> 0 Then
                myFiles.Close
            End If
            Set myFiles = Nothing
        End If
        Set myFiles = gPrepRecordset
        If Len(smExt) = 0 Then
            smExt = ".txt"
        End If
        slFileName = Dir(smSavePath & "*" & smExt)
        Do While slFileName <> ""
            If InStr(1, slFileName, "me") > 0 Then
                slFileName = slFileName
            End If
            slDate = mGetDateProcessed(slFileName)
            If DateDiff("d", smEarlyDate, slDate) > -1 Then
                myFiles.Filter = "FileName = '" & slFileName & "' AND date = '" & CDate(slDate) & "'"
                If myFiles.EOF Then
                    myFiles.AddNew Array("FileName", "Date", "Status", "isProcessed"), Array(slFileName, CDate(slDate), FSPROCESSED, True)
                End If
            End If
            slFileName = Dir$
        Loop
    End If
    If Not myFiles Is Nothing Then
        myFiles.Filter = "isProcessed = True"
        ilCount = myFiles.RecordCount
    End If
    ITransfer_FilesProcessed = ilCount
   
End Property
Private Property Get ITransfer_FilesWaiting(blRefresh As Boolean) As Integer
    Dim ilCount As Integer
    Dim slFileName As String
    Dim slDate As String
    
    ilCount = 0
    If blRefresh Then
        smError = ""
        If myFiles Is Nothing Then
            Set myFiles = gPrepRecordset
        End If
        If Len(smExt) = 0 Then
            smExt = ".txt"
        End If
        slFileName = Dir(smFromPath & "*" & smExt)
        slDate = mGetDate(slFileName)
        Do While slFileName <> ""
            myFiles.Filter = "FileName = '" & slFileName & "' AND date = '" & CDate(slDate) & "'"
            If myFiles.EOF Then
                myFiles.AddNew Array("FileName", "Date", "Status", "isProcessed"), Array(slFileName, CDate(slDate), FSWAITING, False)
            ElseIf myFiles!isProcessed Then
                myFiles!isProcessed = False
                myFiles!Status = FSWAITING
                myFiles!Date = CDate(slDate)
            End If
            slFileName = Dir$
        Loop
    End If
    If Not myFiles Is Nothing Then
        myFiles.Filter = "Status = 'Problem'"
        If Not myFiles.EOF Then
            smError = smError & "Problem with file."
        End If
        myFiles.Filter = "isProcessed = False"
        ilCount = myFiles.RecordCount
    End If
    If ilCount = 0 And bmConnected Then
        mDisConnect
    End If
    
    ITransfer_FilesWaiting = ilCount
End Property
Private Function ITransfer_LoadIni(slIniPath As String) As Boolean
    Dim blRet As Boolean
    Dim slRet As String
    Dim ilRet As Integer
    'change as needed.  some don't need port.
    'do not quit if don't find DaysToSaveOld,  LogPath, or SaveFolder
    'removed mode until will use.
    blRet = True
    If Len(smName) = 0 Then
        smName = MYCHOICE
    End If
    smError = smName & "'s ini is missing these values:"
    'required
    gLoadFromIni smName, "ExportAddress", slIniPath, smAddress
    If smAddress = NOTFOUND Then
        blRet = False
        'use this to signal that this isn't being used!
        smError = smAddress
    Else
        gLoadFromIni smName, "ExportPort", slIniPath, slRet
        If slRet <> NOTFOUND Then
            imPort = CInt(slRet)
        Else
            blRet = False
            smError = smError & " 'Port'"
        End If
'        gLoadFromIni smName, "ExportMode", slIniPath, slRet
'        If slRet <> NOTFOUND Then
'            If UCase(slRet) = "TEL" Then
'                emMode = TelNet
'            ElseIf UCase(slRet) = "FTP" Then
'                emMode = Ftp
'            ElseIf UCase(slRet) = "WEB" Then
'                emMode = WebService
'            Else
'                blRet = False
'            End If
'        Else
'            blRet = False
'            smError = smError & " 'ExportMode'"
'        End If
        gLoadFromIni smName, "ExportFromPath", slIniPath, smFromPath
        If smFromPath = NOTFOUND Then
            blRet = False
            smError = smError & " 'ExportFromPath'"
        Else
            smFromPath = gSetPathEndSlash(smFromPath, False)
        End If
        If Dir(smFromPath, vbDirectory) = vbNullString Then
            smError = smError & "ExportFromPath " & smFromPath & " does not exist."
            blRet = False
        End If
        'may be required
        gLoadFromIni smName, "ExportPassword", slIniPath, smPassword
        If smPassword = NOTFOUND Then
            If bmNoPassword Then
                smPassword = ""
            Else
                blRet = False
                smError = smError & " 'ExportPassword'"
            End If
        End If
        gLoadFromIni smName, "ExportUserName", slIniPath, smUsername
        If smUsername = NOTFOUND Then
            If bmNoPassword Then
                smUsername = ""
                smError = smError & " 'ExportUserName'"
            Else
                blRet = False
            End If
        End If
        '6239
        gLoadFromIni smName, "WrapStart", slIniPath, smWrapStart
        If smWrapStart = NOTFOUND Then
           smWrapStart = ""
        End If
        gLoadFromIni smName, "WrapEnd", slIniPath, smWrapEnd
        If smWrapEnd = NOTFOUND Then
           smWrapEnd = ""
        End If
        gLoadFromIni smName, "Echo", slIniPath, slRet
        If slRet = NOTFOUND Then
           bgEcho = True
        ElseIf UCase(slRet) = "FALSE" Then
            bgEcho = False
        Else
            bgEcho = True
        End If

        gLoadFromIni smName, "DaysToSaveOld", slIniPath, slRet
        If slRet <> NOTFOUND Then
            ilDays = CInt(slRet)
        Else
            ilDays = DEFAULTDAYS
        End If
        gLoadFromIni smName, "exportSavePath", slIniPath, smSavePath
        If smSavePath = NOTFOUND Then
           smSavePath = DEFAULTFOLDER
        Else
            smSavePath = gSetPathEndSlash(smSavePath, False)
        End If
        gLoadFromIni smName, "Confirmation", slIniPath, smKickerConfirmation
        If smKickerConfirmation = NOTFOUND Then
           smKickerConfirmation = KICKERCONFIRMDEFAULT
        End If
        gLoadFromIni smName, "Error", slIniPath, smKickerError
        If smKickerError = NOTFOUND Then
           smKickerError = KICKERERRORDEFAULT
        End If
        gLoadFromIni smName, "ExportStartPaused", slIniPath, slRet
        If UCase(slRet) = "TRUE" Then
            bmStartPaused = True
        Else
            bmStartPaused = False
        End If
On Error GoTo CANTCREATE
        If Dir(smSavePath, vbDirectory) = vbNullString Then
            ilRet = 0
            MkDir smSavePath
            If (ilRet <> 0) And (ilRet <> 75) Then
                smError = smError & "ExportSaveFolder " & smSavePath & " does not exist."
                blRet = False
            End If
        End If
    End If
    bmIsLoaded = blRet
    ITransfer_LoadIni = blRet
    Exit Function
CANTCREATE:
    ilRet = Err.Number
    Resume Next
End Function
Private Function ITransfer_Move(blMarkAsUnread As Boolean) As Boolean
On Error GoTo ERRORBOX
    Dim blRet As Boolean
    Dim slFileName As String
    Dim slAdditional As String
    
    blRet = True
    smError = ""
    If blMarkAsUnread Then
        slAdditional = "UNREAD"
    Else
        slAdditional = ""
    End If
    If Not myFiles.EOF Then
On Error GoTo ERRORBOX
        slFileName = Trim$(myFiles!Filename)
        FileCopy smFromPath & slFileName, smSavePath & slAdditional & slFileName
        Kill smFromPath & slFileName
    Else
        blRet = False
        smError = "Problem. There is no file to 'Move'."
    End If
Cleanup:
    ITransfer_Move = blRet
    Exit Function
ERRORBOX:
    smError = "Could not delete file " & slFileName
    myFiles!Status = FSPROBLEM
    'myFiles.Update
    blRet = False
    GoTo Cleanup
End Function

Private Function ITransfer_Process() As Boolean
    Dim blRet As Boolean
    Dim slProcessResult As String
    
    smError = ""
    slProcessResult = FSPROCESSED
    blRet = False
    If Not myFiles Is Nothing Then
        myFiles.Filter = "isProcessed = False " 'AND Status <> 'Connect Issue'"
        If Not myFiles.EOF Then
On Error GoTo ERRORAGAIN
            DoEvents
            If mProcessMain() Then
                blRet = True
            Else
                blRet = False
                ' this will mark connect so don't rerun this time.
                If InStr(1, smError, "connect", vbTextCompare) > 0 Then
                    slProcessResult = FSNOCONNECT
                Else
                    slProcessResult = FSPROBLEM
                End If
            End If
            smStatus = smStatus & " " & Trim$(myFiles!Filename) & "::" & slProcessResult
            myFiles!isProcessed = True
            myFiles!Status = slProcessResult
           ' myFiles.Update
        End If
    End If
Cleanup:
    ITransfer_Process = blRet
    Exit Function
ERRORAGAIN:
    smError = smName & " problem 'Process':" & Err.Description
    blRet = False
    slProcessResult = FSPROBLEM
    Resume Next
End Function
Private Sub ITransfer_FillGrid(myFileGrid As MSHierarchicalFlexGridLib.IMSHFlexGrid)
    Dim c As Integer
    
    With myFileGrid
        If Not myFiles Is Nothing Then
            myFiles.Filter = adFilterNone
            gGrid_Clear myFileGrid, True
            If Not myFiles.BOF And Not myFiles.EOF Then
                myFiles.MoveFirst
                mGridFillWithRows myFileGrid, myFiles.RecordCount
                c = 0
                Do While Not myFiles.EOF
                    c = c + 1
                    If c < .Rows Then
                        .TextMatrix(c, FILESINDEXFILENAME) = Trim$(myFiles!Filename)
                        .TextMatrix(c, FILESINDEXDATE) = Format(myFiles!Date, "m/d/yyyy")
                        .TextMatrix(c, FILESINDEXTIME) = Format(myFiles!Date, "h:nn am/pm")
                        .TextMatrix(c, FILESINDEXSTATUS) = myFiles!Status
                    End If
                    myFiles.MoveNext
                Loop
            End If
        End If
    End With
End Sub
Private Function ITransfer_Connect() As Boolean
    'connect, username, password...as needed
    Dim blRet As Boolean
    Dim blDisconnect As Boolean
   ' Dim blEchoError As Boolean
    
    blDisconnect = False
    smError = ""
    sgTelNetReturn = ""
   ' blEchoError = False
    blRet = True
    If Not bmConnected Then
        With frmTransfer.ttcControl
            .TermType = "NVT|TTY"
'On Error GoTo ERRECHO
            .Connect smAddress, imPort
'            If blEchoError Then
'                bgEcho = False
'                .Connect smAddress, imPort
'            End If
 On Error GoTo ERRORBOX
            DoEvents
            If mTestTelNetReturn("") Then
                If sgTelNetReturn = "Error" Then
                    blRet = False
                    smError = "Could not connect." & sgTelNetReturn
                Else
                    If mTestTelNetReturn("") Then
                        If Not bmNoPassword Then
                            If InStr(1, sgTelNetReturn, KICKERLOGIN, vbTextCompare) > 0 Then
                                .SendData smUsername & vbCrLf
                                If Not mTestTelNetReturn(smUsername) Then
                                    blRet = False
                                    blDisconnect = True
                                    smError = "Connected, but server didn't return Username."
                                Else
                                    .SendData smPassword & vbCrLf
                                    If Not mTestTelNetReturn() Then
                                        blRet = False
                                        blDisconnect = True
                                        smError = "Connected, but server didn't return Password."
                                    Else
                                         If Not mTestTelNetReturn(smKickerConfirmation) Then
                                            blRet = False
                                            blDisconnect = True
                                            smError = "Connected, but server rejected login information."
                                        End If
                                    End If
                                End If
                            Else
                                blRet = False
                                blDisconnect = True
                                smError = "Could not connect.  Server did not ask for login information"
                            End If
                        End If
                    Else
                    
                    End If
                End If 'error?
            Else
                blRet = False
                smError = "Could not connect.  No response from server."
            End If
            If blDisconnect Then
                .DisConnect
            End If
        End With
    End If
    bmConnected = blRet
    ITransfer_Connect = blRet
    Exit Function
'ERRECHO:
'    blEchoError = True
'    Resume Next
ERRORBOX:
    ITransfer_Connect = False
    smError = "ITransfer_Connect: " & Err.Description
End Function
Private Function ITransfer_DisConnect() As Boolean

    ITransfer_DisConnect = mDisConnect()

End Function
Private Function mDisConnect() As Boolean
    Dim blRet As Boolean

 On Error GoTo ERRORBOX
    blRet = True
    If bmConnected Then
        With frmTransfer.ttcControl
            .DisConnect
        End With
        bmConnected = False
    End If
    mDisConnect = blRet
    Exit Function
ERRORBOX:
    mDisConnect = False
    smError = smName & "-Disconnect: " & Err.Description
End Function
Private Function mGetDate(slName As String) As String
    Dim slRet As String
    Dim myFile As File
    
    slRet = NODATE
On Error GoTo ERRORBOX
    If omFile.FileExists(smFromPath & slName) Then
        Set myFile = omFile.GetFile(smFromPath & slName)
        slRet = myFile.DateCreated
    End If
Cleanup:
    mGetDate = slRet
    Set myFile = Nothing
    Exit Function
ERRORBOX:
    slRet = NODATE
    smError = "Problem with mGetDate"
End Function
Private Function mGetDateProcessed(slName As String) As String
    Dim slRet As String
    Dim myFile As File
    
    slRet = NODATE
On Error GoTo ERRORBOX
    If omFile.FileExists(smSavePath & slName) Then
        Set myFile = omFile.GetFile(smSavePath & slName)
        slRet = myFile.DateCreated
    End If
Cleanup:
    mGetDateProcessed = slRet
    Set myFile = Nothing
    Exit Function
ERRORBOX:
    slRet = NODATE
    smError = "Problem with mGetDateProcessed"
End Function
Private Sub mGridFillWithRows(grdCtrl As MSHFlexGrid, llFillNoRow As Long)
    Dim llRows As Long
    Dim llCols As Long

    With grdCtrl
        For llRows = .FixedRows To .FixedRows + llFillNoRow Step 1
            If llRows >= .Rows Then
                .AddItem ""
            End If
            .Row = llRows
            For llCols = 0 To .Cols - 1 Step 1
                .TextMatrix(llRows, llCols) = ""
                .Col = llCols
                .CellBackColor = LIGHTYELLOW
            Next llCols
        Next llRows
        'do last row.
        .Row = .Rows - 1
         For llCols = 0 To .Cols - 1 Step 1
            .Col = llCols
            .CellBackColor = LIGHTYELLOW
        Next llCols
       
    End With
End Sub

Private Function mProcessMain() As Boolean
    'change as needed
    'recordset already done before here
    Dim blRet As Boolean
    Dim olFiles As FileSystemObject
    Dim olThisFile As TextStream
    Dim slLine As String
    Dim blThisLineOk As Boolean
    Dim slToday As String
    Dim slDateFromLine As String
    Dim slTestLine As String
    
 On Error GoTo ERRORBOX
    blRet = False
    If Not bmConnected Then
        smStatus = ""
        If Not ITransfer_Connect Then
            GoTo Cleanup
        End If
    End If
    slToday = Format(Now(), "yymmdd")
    DoEvents
    Set olFiles = New FileSystemObject
    If olFiles.FileExists(smFromPath & Trim(myFiles!Filename)) Then
        Set olThisFile = olFiles.OpenTextFile(smFromPath & Trim(myFiles!Filename))
        blRet = True
        Do
            slLine = olThisFile.ReadLine
            slLine = Trim$(slLine)
            'playlist will add to line, but want to test without addtion
            slTestLine = slLine
            If Len(slLine) = 0 Then
                blThisLineOk = False
            Else
                blThisLineOk = True
                If UCase(smExt) = ".WEG" Then
                    If InStr(1, slLine, "DELPLAYLIST") > 0 And InStr(1, slLine, slToday) > 0 Then
                        smError = smError & " A deletion could not be sent for " & Trim(myFiles!Filename) & " because of date. "
                        blRet = False
                        blThisLineOk = False
                    End If
'                       '6239
                    If Len(smWrapStart) > 0 And Len(smWrapEnd) > 0 Then
                         slLine = smWrapStart & " """ & slLine & """ " & smWrapEnd
                    ElseIf Len(smWrapStart) = 0 And Len(smWrapEnd) > 0 Then
                         slLine = " """ & slLine & """ " & smWrapEnd
                    ElseIf Len(smWrapEnd) = 0 And Len(smWrapStart) > 0 Then
                        slLine = " """ & slLine & """ " & smWrapEnd
                    End If
                   ' End If
                End If
            End If
            DoEvents
            If blThisLineOk Then
                frmTransfer.ttcControl.Echo True
                frmTransfer.ttcControl.SendData slLine & vbCrLf
                If Not mTestTelNetReturn(slTestLine, True) Then
                    smError = "Error with line in " & smFromPath & myFiles!Filename & ":" & slLine & " " & smError
                    blRet = False
                End If
            End If
            slLine = ""
        Loop Until olThisFile.AtEndOfStream
    Else
        smError = smError & " " & smName & "-mProcessMain: couldn't find: " & smFromPath & myFiles!Filename
    End If
Cleanup:
    Set olFiles = Nothing
On Error Resume Next
    olThisFile.Close
On Error GoTo 0
    Set olThisFile = Nothing
    mProcessMain = blRet
    Exit Function
ERRORBOX:
    mProcessMain = False
    smError = smError & " : " & Err.Description
End Function
Private Function mTestTelNetReturn(Optional slDesiredReturn As String = "", Optional blTestError As Boolean = False) As Boolean
    Dim llCount As Long
    Dim ilPos As Integer
    Dim slError As String
    Dim ilEnd As Integer
    
    slError = ""
    bgCantClose = True
    llCount = 0
    Do While sgTelNetReturn = "" And llCount < 50
        Sleep 1000
        DoEvents
        llCount = llCount + 1
    Loop
    If sgTelNetReturn <> "" Then
        If slDesiredReturn <> "" Then
            If InStr(1, UCase(sgTelNetReturn), UCase(slDesiredReturn)) = 0 Then
                mTestTelNetReturn = False
            Else
                ilPos = InStr(1, sgTelNetReturn, smKickerError) 'bash:
                If ilPos = 0 Then
                    mTestTelNetReturn = True
                Else
                    mTestTelNetReturn = False
                    ilEnd = InStr(ilPos, sgTelNetReturn, vbCr)
                    If ilEnd < Len(sgTelNetReturn) And ilEnd > 0 Then
                        slError = Mid(sgTelNetReturn, ilPos, ilEnd - ilPos)
                    Else
                        slError = Mid(sgTelNetReturn, ilPos)
                    End If
                    smError = slError & vbCrLf & smError
                End If
            End If
        Else
            mTestTelNetReturn = True
        End If
    Else
        mTestTelNetReturn = False
    End If
    bgCantClose = False
End Function
