Attribute VB_Name = "Transfer"
Option Explicit
Public bgEcho As Boolean

'copied from other areas of affiliate:
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize As Long, ByVal lpFileName$)
Public Const GRIDSCROLLWIDTH = 270
Public Const LIGHTYELLOW = &HC0FFFF '&HBFFFFF '&H80FFFF '&HBFFFFF
Public DARKGREEN As Long ' = RGB(0, 128, 0)      'rgb(0,128,0)
Public Const LIGHTGREEN = &H80FF80
Public Const LIGHTBLUE = &HFDFFD7
Public Const GRAY = &HC0C0C0
Public Const ORANGE = 33023 '&H80FF
Public Const GREEN = 49152
Public Const BROWN = 128
'end copy
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long
'FTP
Type CSIFTPINFO
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sSendFolder As String * 128
   sRecvFolder As String * 128
   sServerDstFolder As String * 128
   sServerSrcFolder As String * 128
   sLogPathName As String * 128
End Type
 
Type CSIFTPSTATUS
   iState As Integer    ' 0=Complete, 1=Busy.
   iStatus As Integer   ' 0=Success, 1=Errors occured
   iJobCount As Integer ' The # of files yet to process.
   lLastError As Long   ' Contains the results of GetLastError if an error occurs.
End Type
 
Type CSIFTPERRORINFO
    sInfo As String * 1024
    sFileThatFailed As String * 128
End Type
 
Type CSIFTPFILELISTING
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sPathFileMask As String * 128
   sSavePathFileName As String * 128
   sLogPathName As String * 128
   nTotalFiles As Integer
End Type
'Type IDCWEBSETTINGS
'Type CSIFTPSETTINGS
'    sUrl As String
'    '\mp3files. Used later to build path and mask
'    sPathToFiles As String
'    myFTP As CSIFTPFILELISTING
'    bFindFile As Boolean
'End Type
'Public tgCsiFtpFileListing As CSIFTPFILELISTING


    
Declare Function csiFTPInit Lib "CSI_Utils.dll" (ByRef FTPInfo As CSIFTPINFO) As Integer
Declare Function csiFTPRenameFile Lib "CSI_Utils.dll" (ByVal szCurrentFileName$, ByVal szNewFileName$) As Integer
Declare Function csiFTPGetFileListing Lib "CSI_Utils.dll" (ByRef FTPFileListing As CSIFTPFILELISTING) As Integer
Declare Function csiFTPGetStatus Lib "CSI_Utils.dll" (ByRef FTPStatus As CSIFTPSTATUS) As Integer
Declare Function csiFTPGetError Lib "CSI_Utils.dll" (ByRef FTPErrorInfo As CSIFTPERRORINFO) As Integer
Declare Function csiFTPFileFromServer Lib "CSI_Utils.dll" (ByVal slFileName$) As Integer
'End FTP

Public Const NOTFOUND = "Not Found"
Public Const NODATE As Date = "1/1/1970"
Public Const FSPROCESSED As String = "Processed"
Public Const FSWAITING As String = "Unprocessed"
Public Const FSPROBLEM As String = "Problem"
Public Const FSNOCONNECT As String = "Connect Issue"
Public Const FSPENDING As String = "Pending"
'Public Enum TransferMode
'    TelNet
'    Ftp
'    WebService
'    none
'End Enum
Public bgCantClose As Boolean
Public sgStartupDirectory As String
Public sgDbPath As String
Public sgExeDirectory As String
'for logging
Public sgUserName As String
Public sgImportDirectory As String
Public sgExportDirectory As String
Public sgMsgDirectory As String
Public igExportSource As Integer
Public sgCallStack(0 To 9) As String

Public myIPump As ITransfer
Public myIPumpImport As ITransfer
Public myGeneric As ITransfer
''must change
'Public myIDC As ITransfer
''must change
'Public myMarketron As ITransfer

Public sgTelNetReturn As String

Public Const FILESINDEXSTATUS As Integer = 3
Public Const FILESINDEXFILENAME As Integer = 0
Public Const FILESINDEXDATE As Integer = 1
Public Const FILESINDEXTIME As Integer = 2
Public Const LFCR = vbLf & vbCr

Public Sub gWriteIni(Filename As String, Section As String, Key As String, Value As String)
  WritePrivateProfileString Section, Key, Value, Filename
End Sub
Public Function gPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "FileName", adChar, 50
            .Append "Date", adDate
            .Append "Status", adChar, 20
            .Append "isProcessed", adBoolean
        End With
    myRs.Open
    myRs.Sort = "status asc, Date asc"
    Set gPrepRecordset = myRs
End Function
'copied from other areas of affiliate
Public Sub gGrid_Clear(grdCtrl As MSHFlexGrid, ilFillRows As Integer)
    
'
'   grdCtrl (I)-  Grid Control name
'   ilFillRows (I)- True=Fill Grid with blank rows; False=Only have one blank row
'
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    
    If ilFillRows Then
        llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - 2
    Else
        llFillNoRow = 0
    End If
    llRows = grdCtrl.Rows
    Do While llRows >= grdCtrl.FixedRows + llFillNoRow + 1
        If llRows <= grdCtrl.FixedRows + 1 Then
            Exit Do
        End If
        grdCtrl.RemoveItem llRows - 1
        llRows = llRows - 1
    Loop
    If ilFillRows Then
        gGrid_FillWithRows grdCtrl
    Else
        llRows = grdCtrl.FixedRows
        If llRows >= grdCtrl.Rows Then
            Do While llRows >= grdCtrl.Rows
                grdCtrl.AddItem ""
            Loop
'        Else
'            llRows = grdCtrl.Rows
'            For llCols = 0 To grdCtrl.Cols - 1 Step 1
'                grdCtrl.TextMatrix(llRows, llCols) = ""
'            Next llCols
        End If
    End If
    llRows = grdCtrl.FixedRows
    Do While llRows < grdCtrl.Rows
        For llCols = 0 To grdCtrl.Cols - 1 Step 1
            grdCtrl.TextMatrix(llRows, llCols) = ""
        Next llCols
        llRows = llRows + 1
    Loop
End Sub
Public Sub gGrid_IntegralHeight(grdCtrl As MSHFlexGrid)
'    If grdCtrl.Rows > 0 Then
'        If (grdCtrl.Height - 15) Mod grdCtrl.RowHeight(grdCtrl.FixedRows) <> 0 Then
'            'grdHistory.Height = ((grdHistory.Height \ grdHistory.RowHeight(1)) + 1) * grdHistory.RowHeight(1) + 15
'            grdCtrl.Height = (grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows)) * grdCtrl.RowHeight(grdCtrl.FixedRows) + 15
'        End If
'    End If
    Dim llHeight As Long
    Dim llRow As Long
    
    llHeight = 0
    If grdCtrl.FixedRows > 0 Then
        For llRow = 1 To grdCtrl.FixedRows Step 1
            llHeight = llHeight + grdCtrl.RowHeight(llRow - 1)
        Next llRow
    End If
    Do While llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows) < grdCtrl.Height
        llHeight = llHeight + grdCtrl.RowHeight(grdCtrl.FixedRows)
    Loop
    grdCtrl.Height = llHeight + 15
End Sub

Public Sub gGrid_AlignAllColsLeft(grdCtrl As MSHFlexGrid)
    Dim ilCol As Integer
    
    For ilCol = 0 To grdCtrl.Cols - 1 Step 1
        grdCtrl.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol

End Sub

Public Sub gGrid_FillWithRows(grdCtrl As MSHFlexGrid)
    Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    
    llFillNoRow = grdCtrl.Height \ grdCtrl.RowHeight(grdCtrl.FixedRows) - grdCtrl.FixedRows - 1
    
    For llRows = grdCtrl.FixedRows To grdCtrl.FixedRows + llFillNoRow Step 1
        Do While llRows >= grdCtrl.Rows
            grdCtrl.AddItem ""
            For llCols = 0 To grdCtrl.Cols - 1 Step 1
                grdCtrl.TextMatrix(llRows, llCols) = ""
            Next llCols
        Loop
    Next llRows
End Sub

Public Function gGrid_DetermineRowCol(grdCtrl As MSHFlexGrid, X As Single, Y As Single) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llColLeftPos As Long
    
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        If grdCtrl.RowIsVisible(llRow) Then
            If (Y >= grdCtrl.RowPos(llRow)) And (Y <= grdCtrl.RowPos(llRow) + grdCtrl.RowHeight(llRow) - 15) Then
                llColLeftPos = grdCtrl.ColPos(0)
                For llCol = 0 To grdCtrl.Cols - 1 Step 1
                    If grdCtrl.ColWidth(llCol) > 0 Then
                        If (X >= llColLeftPos) And (X <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
                            grdCtrl.Row = llRow
                            grdCtrl.Col = llCol
                            gGrid_DetermineRowCol = True
                            Exit Function
                        End If
                        llColLeftPos = llColLeftPos + grdCtrl.ColWidth(llCol)
                    End If
                Next llCol
            End If
        End If
    Next llRow
    gGrid_DetermineRowCol = False
    Exit Function
End Function
Public Function gSetPathEndSlash(ByVal slInPath As String, ilAdjDrivePath As Integer) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If Right$(slPath, 1) <> "\" Then
        slPath = slPath + "\"
    End If
'    If ilAdjDrivePath Then
'        slPath = gAdjustDrivePath(slPath)
'    End If
    gSetPathEndSlash = slPath
End Function
Public Function gXmlIniPath(Optional slName As String = "") As String
'modified!
    Dim oMyFileObj As FileSystemObject
    Dim slIniPath As String
    
    If slName = "" Then
        slName = "xml.ini"
    End If
    Set oMyFileObj = New FileSystemObject
    '4/14/12: Changed order of search to: Startup; database; exe From: database; exe; startup
    '         the reason for the change to to handle the case were two xml.ini definded and one placed
    '         in the database and the other placed into an ini folder.
    slIniPath = oMyFileObj.BuildPath(sgStartupDirectory, slName)
    If oMyFileObj.FileExists(slIniPath) Then
        gXmlIniPath = slIniPath
    Else
        slIniPath = oMyFileObj.BuildPath(sgDbPath, slName)
        If oMyFileObj.FileExists(slIniPath) Then
            gXmlIniPath = slIniPath
        Else
            slIniPath = oMyFileObj.BuildPath(sgExeDirectory, slName)
            If Not oMyFileObj.FileExists(slIniPath) Then
                slIniPath = vbNullString
            End If
            gXmlIniPath = slIniPath
        End If
    End If
    Set oMyFileObj = Nothing
    
End Function
Public Function gLoadFromIni(Section As String, Key As String, slPath As String, sValue As String) As Boolean
    'no change
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    
    sValue = "Not Found"
    gLoadFromIni = False
    If Dir(slPath) > "" Then
        BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slPath)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                sValue = Left(sBuffer, BytesCopied)
                gLoadFromIni = True
            End If
        End If
    End If 'slPath not valid?
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function
'changed for this module only!
Function gNow() As String
    Dim slDate As String
    Dim slDateTime As String
    Dim ilPos As Integer

    If slDate = "" Then
        gNow = Now
    Else
        slDateTime = Trim$(Now)
        ilPos = InStr(1, slDateTime, " ", 1)
        If ilPos > 0 Then
            gNow = slDate & Mid$(slDateTime, ilPos)
        Else
            gNow = slDate
        End If
    End If
End Function
Public Sub gCenterStdAlone(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2 + 115 '+ Screen.Height \ 10
End Sub
Public Function gFixQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = "'" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixQuote = sOutStr
End Function

Public Function gFixDoubleQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = """" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixDoubleQuote = sOutStr
End Function
Public Function gAddQuotes(slInStr As String) As String
    gAddQuotes = """" & slInStr & """"
End Function

'end copy
