VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'depends on gNow(),sgUserName  Change MessageFolder,ImportFolder,ExportFolder for traffic or affiliate
' Using the logger:
'    Set myFacts = New CLogger
'    With myFacts
'         'CreateLogName adds "Log" and date and '.txt'
'        .LogPath = .CreateLogName(this.MessageFolder & "IDCError")
'        .WriteXml slSend, slReturn
        'True-- do you want date and time added to facts?
'        WriteFacts "Now writing fact" , True
        'True-- do you want time added to warning?
'        .WriteWarning "Could not read values from xml.ini", True
        'True- just time (date/time or time always added to error  True-stack trace?
'       .WriteError "Test by Dan",True,True
'    End With
' if the logpath cannot be found or created, no file is created
' myFile is public so it is usable wherever logger is used.
'Warning
':::::: WARNING:::::::
'3:41:26 PM   Could not read values from xml.ini.
':::::::::::::::::::
'Error
'3:29:14 PM User: Guide. ERROR: test by Dan
'Fact
'3/15/13 3:24:14 PM Now writing fact
'xml
':::12/15/1999 9:48:48 AM Begin Transmission:::
'  SENT
' soap envelope here
'  RETURN MESSAGE
' return soap envelope
'::: End Transmission :::

Public myFile As FileSystemObject
Public BlockUserName As Boolean
Private bmIsLog As Boolean
Private smLogPath As String
Private smError As String
'surround each line with :::3/1/13 12:48:45 PM Begin Transmission:::  and "Send" and "Return, as needed.
Private bmIsXml As Boolean
'surround with ::::WARNING:::
Private bmIsWarning As Boolean
'how many days before create new?  -1 is never 0 is on every 'start' otherwise, number of days. Default is never
Private bmIsDateTime As Boolean
'6539  needs to know if new file to write header
Private bmIsNew As Boolean
Private Const LOGDATE As String = "mm-dd-yy"
Private Const CLEANTEXT As String = "LOG_"
Private Const CLEANEXPORT As Integer = 60
Private Const CLEANIMPORT As Integer = 60
'7635 changed from 7 7/20/15 changed from 60 days
Private Const CLEANMESSAGE As Integer = 90
'cleaning up folders
Private bmCleanTested As Boolean
Private imCleanFolder As FoldersToClean
Public Enum FoldersToClean
    None = 0
    messages = 1
    exports = 2
    imports = 3
    All = 4
End Enum
Public Property Get MessageFolder() As String
    'affiliate
    MessageFolder = sgMsgDirectory
    'traffic
    'MessageFolder = sgdbpath & "Messages\"
End Property
Public Property Get ImportFolder() As String
    'affiliate
    ImportFolder = sgImportDirectory
    'traffic
    'ImportFolder = sgImportPath
End Property
Public Property Get ExportFolder() As String
    'affiliate
    ExportFolder = sgExportDirectory
    'traffic
    'ExportFolder = sgExportPath
End Property
Public Property Get UserName() As String
    UserName = sgUserName
End Property
Public Property Let CleanThisFolder(FolderToClean As FoldersToClean)
    imCleanFolder = FolderToClean
    bmCleanTested = False
End Property
Public Property Get CleanThisFolder() As FoldersToClean
    CleanThisFolder = imCleanFolder
End Property

Public Property Get isNew() As Boolean
    isNew = bmIsNew
End Property
Public Property Get isLog() As Boolean
    isLog = bmIsLog
End Property
Public Property Get LogPath() As String
    LogPath = smLogPath
End Property
Public Property Let LogPath(slvalue As String)
On Error GoTo ERRORBOX
    '8886
    'If Dir(slvalue) > "" Then
    If myFile.FILEEXISTS(slvalue) Then
        smLogPath = slvalue
        bmIsLog = True
        Exit Property
    End If
    ' create file
    myFile.CreateTextFile slvalue, True
    '8886
    'If Dir(slvalue) > "" Then
    If myFile.FILEEXISTS(slvalue) Then
        smLogPath = slvalue
        bmIsLog = True
        bmIsNew = True
    Else
        smLogPath = ""
        bmIsLog = False
    End If
    Exit Property
ERRORBOX:
    bmIsLog = False
    smLogPath = ""
    smError = Err.Description
End Property
Public Property Get ErrorMessage()
    ErrorMessage = smError
End Property
Public Sub WriteXml(slSend As String, slReturn As String)
    bmIsXml = True
    WriteToFile slSend, slReturn
    bmIsXml = False
End Sub
Public Sub WriteWarning(slWarning As String, Optional blWithTime As Boolean = False)
    bmIsWarning = True
    If blWithTime Then
        bmIsDateTime = True
    End If
    WriteToFile slWarning, , blWithTime
    bmIsDateTime = False
    bmIsWarning = False
End Sub
Public Sub WriteError(slError As String, Optional blWithTimeNoDate As Boolean = False, Optional blIsStackTrace As Boolean = True)
    Dim slLine As String
    
    slLine = ""
    '7481 moved
'    If Len(UserName) > 0 Then
'        slLine = "User: " & UserName & "."
'    End If
'    slLine = slLine & " ERROR: " & slError
    sgTmfStatus = "E"
    slLine = " ERROR: " & slError
    bmIsDateTime = True
    WriteToFile slLine, , blWithTimeNoDate, blIsStackTrace
    bmIsDateTime = False
End Sub
Public Sub WriteFacts(slFact As String, Optional blWithDateTime As Boolean = False)
    If blWithDateTime Then
        bmIsDateTime = True
    End If
    WriteToFile slFact
    bmIsDateTime = False
End Sub
Public Sub CleanFolder(Optional slOverrideFolder As String = "", Optional ilOverrideDays As Integer = 0)
    '8866 lost passing optional slName
    Dim slFolder As String
    Dim ilDays As Integer
    
    If bmCleanTested And Len(slOverrideFolder) = 0 And ilOverrideDays = 0 Then
        Exit Sub
    End If
    slFolder = ""
    ilDays = 0
    If Len(slOverrideFolder) > 0 Then
        slFolder = slOverrideFolder
    ElseIf imCleanFolder = All Then
        If ilOverrideDays > 0 Then
            mCleanFolders Me.MessageFolder, ilOverrideDays
            mCleanFolders Me.ImportFolder, ilOverrideDays
            mCleanFolders Me.ExportFolder, ilOverrideDays
           
        Else
            mCleanFolders Me.MessageFolder, CLEANMESSAGE
            mCleanFolders Me.ImportFolder, CLEANIMPORT
            mCleanFolders Me.ExportFolder, CLEANEXPORT
        End If
    ElseIf imCleanFolder = None Then
        Exit Sub
    Else
        Select Case imCleanFolder
            Case messages
                slFolder = Me.MessageFolder
                ilDays = CLEANMESSAGE
            Case exports
                slFolder = Me.ExportFolder
                ilDays = CLEANEXPORT
            Case imports
                slFolder = Me.ImportFolder
                ilDays = CLEANIMPORT
        End Select
     End If
     If ilOverrideDays > 0 Then
        ilDays = ilOverrideDays
    End If
    mCleanFolders slFolder, ilDays
    bmCleanTested = True
End Sub
Public Function IsLogFile(slName As String) As Boolean
    '    slFileName = Dir(slFolder & slType & CLEANTEXT & "??-??-??.txt")
    Dim blRet As Boolean
    Dim ilPos As Integer
    Dim ilPosHyphens As Integer
    Dim ilLength As Integer
    Dim slHyphen As String

    blRet = False
    slHyphen = "-"
    ilPos = InStr(1, slName, CLEANTEXT, vbTextCompare)
    If ilPos > 0 Then
        '30
        ilLength = Len(slName)
        '27
        ilPosHyphens = InStrRev(slName, ".txt")
        '.txt at end of line
        If ilPosHyphens + 3 = ilLength Then
            'now - or underscore
            '24 + 6 6 is where last - should be
            If InStrRev(slName, slHyphen, ilPosHyphens) + 6 = ilLength Then
                blRet = True
            Else
                slHyphen = "_"
                If InStrRev(slName, slHyphen, ilPosHyphens) + 6 = ilLength Then
                    blRet = True
                End If
            End If
            'now we know they hyphen type. Reuse 'ilPosHyphens'. Let's continue
            If blRet Then
                'from 15 to 19
                ilPos = ilPos + Len(CLEANTEXT)
                'ilposhyphens is still start of .txt
                If ilPos < ilPosHyphens Then
                    '21 test for hypyhen before the hyphen we previously found
                    ilPosHyphens = InStrRev(slName, slHyphen, ilPosHyphens - 5)
                    If ilPosHyphens + 9 = ilLength Then
                        If ilPosHyphens = ilPos + 2 Then
                            blRet = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    IsLogFile = blRet
End Function
Public Sub CleanFile(slPath As String, ilDays As Integer)
    Dim myCurrent As file
    Dim slCurrentDate As String
    
On Error GoTo ERRORBOX
    If myFile.FILEEXISTS(slPath) Then
        If ilDays = 0 Then
            myFile.DeleteFile slPath
        Else
            Set myCurrent = myFile.GetFile(slPath)
            slCurrentDate = myCurrent.DateCreated
            If DateDiff("d", slCurrentDate, Now()) > ilDays Then ' was gNow()
                myFile.DeleteFile slPath
            End If
        End If
    End If
Cleanup:
    Set myCurrent = Nothing
    Exit Sub
ERRORBOX:
    smError = "Error deleting " & slPath
    GoTo Cleanup
End Sub

Private Sub mCleanFolders(slFolder As String, ilDays As Integer)
    '8886 removed slType As String,
   ' Dim slFileName As String
    Dim myCurrent As file
    Dim slCurrentDate As String
    Dim olFolder As Folder

On Error GoTo ERRORBOX
    '8886
    ' don't need to add "*" after sltype
'    If Len(slType) < 1 Then
'        slType = "*"
'    Else
'        slType = slType & "*"
'    End If
'    slFileName = Dir(slFolder & slType & CLEANTEXT & "??-??-??.txt")
'    Do While slFileName > ""
'        If myFile.FILEEXISTS(slFolder & slFileName) Then
'            Set myCurrent = myFile.GetFile(slFolder & slFileName)
'            slCurrentDate = myCurrent.DateCreated
'            If DateDiff("d", slCurrentDate, Now()) > ilDays Then
'                myFile.DeleteFile slFolder & slFileName
'            End If
'        End If
'       slFileName = Dir()
'    Loop
'    '8108
'    slFileName = Dir(slFolder & slType & CLEANTEXT & "??_??_??.txt")
'    Do While slFileName > ""
'        If myFile.FILEEXISTS(slFolder & slFileName) Then
'            Set myCurrent = myFile.GetFile(slFolder & slFileName)
'            slCurrentDate = myCurrent.DateCreated
'            If DateDiff("d", slCurrentDate, Now()) > ilDays Then
'                myFile.DeleteFile slFolder & slFileName
'            End If
'        End If
'       slFileName = Dir()
'    Loop
'    If Len(slType) < 1 Then
'        slType = "*"
'    Else
'        slType = slType
'    End If
    If myFile.FolderExists(slFolder) Then
        Set olFolder = myFile.GetFolder(slFolder)
        For Each myCurrent In olFolder.Files
            If Me.IsLogFile(myCurrent.Name) Then
                slCurrentDate = myCurrent.DateCreated
                If DateDiff("d", slCurrentDate, Now()) > ilDays Then
                    myFile.DeleteFile myCurrent
                End If
            End If
        Next
    End If
Cleanup:
    Set myCurrent = Nothing
    Exit Sub
ERRORBOX:
    smError = "Error in mCleanFolders deleting from" & slFolder & " " & Err.Description
    GoTo Cleanup
End Sub
Private Sub WriteToFile(ByVal slSend As String, Optional slReturn As String = "", Optional blIsTimeNoDate As Boolean = False, Optional blIsStackTrace As Boolean = False)
    Dim blRet As Boolean
    Dim olFile As TextStream
    Dim slDateTime As String
    
 On Error GoTo ERRORBOX
    If blIsTimeNoDate Then
        slDateTime = Format(TIME$(), "Long Time")
    Else
        slDateTime = gNow()
    End If
    If bmIsLog Then
        Set olFile = myFile.OpenTextFile(smLogPath, ForAppending, True)
        If bmIsXml Then
            olFile.WriteLine ":::" & slDateTime & " Begin Transmission:::"
            If Len(slReturn) > 0 Then
                olFile.WriteLine "  SENT "
            End If
'        Else
        '7548
        ElseIf Not BlockUserName Then
            '7481
            If Len(UserName) > 0 Then
                slSend = "User: " & UserName & ". " & slSend
            End If
        End If
        If bmIsDateTime Then
            slSend = slDateTime & " " & slSend
        End If
        If bmIsWarning Then
            olFile.WriteLine "::::::WARNING:::::::"
        End If
        olFile.WriteLine slSend
        If bmIsWarning Then
            olFile.WriteLine ":::::::::::::::::::"
            olFile.WriteBlankLines 1
        End If
        If Len(slReturn) > 0 Then
            If bmIsXml Then
                olFile.WriteLine "  RETURN MESSAGE"
            End If
            If bmIsDateTime Then
        
            End If
            olFile.WriteLine slReturn
        End If
        If bmIsXml Then
            olFile.WriteLine "::: End Transmission :::"
            olFile.WriteBlankLines 1
        End If
        If blIsStackTrace Then
            mSaveStackTrace olFile
        End If

        olFile.Close
        Set olFile = Nothing
    End If
    Exit Sub
ERRORBOX:
    smError = "Error in CLogger-WriteToFile: " & Err.Description
End Sub
Public Function CreateLogName(slPathAndName As String) As String
    Dim ilPos As Integer
    Dim slLog As String
On Error GoTo ERRORBOX
    slLog = ""
    CreateLogName = ""
    ilPos = InStrRev(slPathAndName, ".")
    If ilPos > 0 Then
        slLog = Mid(slPathAndName, 1, ilPos - 1)
    Else
        slLog = slPathAndName
    End If
    ilPos = InStrRev(slPathAndName, "\")
    slLog = slLog & "Log_" & Format(gNow(), LOGDATE) & ".txt"
    CreateLogName = slLog
    Exit Function
ERRORBOX:
    CreateLogName = ""
End Function
Public Function FileNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    'Remove " and '
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    FileNameFilter = slName
    
End Function
Public Function CsvSafe(slString As String) As String
'trim, replace " with ' and surround with ""
    Dim slRet As String
    
    slRet = Trim$(slString)
    slRet = gFixDoubleQuote(slRet)
    slRet = gAddQuotes(slRet)
    CsvSafe = slRet
End Function
Private Sub Class_Initialize()
    Clear True
End Sub
Private Sub Clear(blError As Boolean)
    bmIsNew = False
    bmIsLog = False
    smLogPath = ""
    bmIsXml = False
    bmIsDateTime = False
    bmIsWarning = False
    bmCleanTested = False
    imCleanFolder = None
    BlockUserName = False
    If blError Then
        smError = ""
    End If
    Set myFile = New FileSystemObject
End Sub

Private Sub mSaveStackTrace(olFile As TextStream)
    Dim ilTo As Integer
    Dim ilLoop As Integer
    Dim ilTotalLen As Integer
    Dim slMethodName As String
    Dim slLine As String
On Error GoTo ERRORBOX
    slLine = ""
    If Not olFile Is Nothing Then
        ' Verify there is at least one item on the call stack. Otherwise don't print anything.
        ilTotalLen = 0
        For ilLoop = LBound(sgCallStack) To UBound(sgCallStack)
            slMethodName = sgCallStack(ilLoop)
            ilTotalLen = ilTotalLen + Len(slMethodName)
        Next ilLoop
        If ilTotalLen < 1 Then
            Exit Sub
        End If
        For ilLoop = LBound(sgCallStack) To UBound(sgCallStack)
            slMethodName = sgCallStack(ilLoop)
            If Len(slMethodName) > 0 Then
                slLine = slLine & "," & slMethodName
            End If
        Next ilLoop
        olFile.WriteLine "Stack Trace:" & slLine
        olFile.WriteBlankLines 1
    End If
    Exit Sub
ERRORBOX:
End Sub

