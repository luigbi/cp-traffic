Attribute VB_Name = "modXML"
'******************************************************
'*  modXML - various global declarations for importing
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
'7496
Public sgAudioExtension As String
'ttp 5457
Type XDIGITALSTATIONINFO
    sCallLetters As String
    sBand As String
    sSiteId As String
    sFrequency As String
    sOwnership As String
    sAddress As String
    sAddress2 As String
    sCity As String
    sState As String
    sZip As String
    sContactName As String
    sEmail As String
    sPhone As String
    sCell As String
End Type
'ttp 5589
Type XDIGITALAGREEMENTINFO
    sCode As String
    sStation As String
    sStartDate As String
    sEndDate As String
    sSiteId As String
    sProgramCode As String
    sProgramName As String
'6725
    sNetworkId As String
    sStatus As String
End Type
'6741
Type XDIGITALVEHICLEINFO
    iNetworkID As Integer
    sName As String
    iCode As Integer
End Type
Type XFRESPLITSDFCODE
    lSdfCode As Long
    lLogDate As Long
    lLogTime As Long
End Type
Type IDCWEBSETTINGS
    sUrl As String
    '\mp3files. Used later to build path and mask
    sPathToFiles As String
    myFTP As CSIFTPFILELISTING
    bFindFile As Boolean
    'if above is true, stop if don't find?
   ' bMustFile As Boolean
End Type
 '6581
Public Const XMLERRORFILE As String = "XMLErrorResponse.txt"
Declare Function csiFTPGetFileListing Lib "CSI_Utils.dll" (ByRef FTPFileListing As CSIFTPFILELISTING) As Integer
Public Function gXDStationContact(ilShttCode As Integer, slStationInfo As XDIGITALSTATIONINFO) As Boolean
'5457
    'return true/false..false if error
    'also return slStationInfo for contact
    'rule: returns contact for station where 'aff e-mail' is checked, and station set for XDigital.  Which contact to write if many?  order by last name, first name.  If no name, then by email.
    
    Dim rstContact As ADODB.Recordset
    Dim blFound As Boolean
    Dim blRet As Boolean
    
    blFound = False
    'not an error
    blRet = True
    With slStationInfo
        .sContactName = ""
        .sEmail = ""
        .sPhone = ""
    End With
    If ilShttCode > 0 Then
 On Error GoTo errbox
         SQLQuery = "select arttFirstName,arttLastName,arttPhone,arttEmail from artt inner join shtt on arttShttCode =shttCode" & _
        " WHERE arttWebEMail = 'Y' and shttUsedForXDigital = 'Y'  " & _
        " AND arttShttcode = " & ilShttCode & " ORDER BY arttLastName, arttFirstName, arttEmail"
        Set rstContact = gSQLSelectCall(SQLQuery)
        If Not rstContact.EOF Then
        'ttp 5618
            rstContact.Filter = "arttLastName <> ''"
            'first, any with a name
            If Not rstContact.EOF Then
                With rstContact
                    slStationInfo.sEmail = .Fields("arttemail").Value
                    slStationInfo.sContactName = gXMLNameFilter(.Fields("arttFirstName").Value) & " " & gXMLNameFilter(.Fields("arttLastName").Value)
                    slStationInfo.sPhone = .Fields("arttPhone").Value
                End With
                blFound = True
            End If
            If Not blFound Then
                rstContact.Filter = adFilterNone
                With rstContact
                    slStationInfo.sEmail = .Fields("arttemail").Value
                    slStationInfo.sContactName = gXMLNameFilter(.Fields("arttFirstName").Value) & " " & gXMLNameFilter(.Fields("arttLastName").Value)
                    slStationInfo.sPhone = .Fields("arttPhone").Value
                End With
            End If
        End If
Cleanup:
        If Not rstContact Is Nothing Then
            If (rstContact.State And adStateOpen) <> 0 Then
                rstContact.Close
            End If
            Set rstContact = Nothing
        End If
   End If
    gXDStationContact = blRet
    Exit Function
errbox:
    gHandleError "", "ModXml-gXDStationContact"
    blRet = False
End Function
Public Function gIsSiteXDStation() As Boolean
    Dim rstCount As ADODB.Recordset
    Dim blRet As Boolean
    
    blRet = False
On Error GoTo ErrHandler
    SQLQuery = "SELECT count(*) as amount FROM Site Where siteCode = 1 AND siteStationToXDS = 'Y'"
    Set rstCount = gSQLSelectCall(SQLQuery)
    If rstCount!amount > 0 Then
        blRet = True
    End If
Cleanup:
    If Not rstCount Is Nothing Then
        If (rstCount.State And adStateOpen) <> 0 Then
            rstCount.Close
        End If
        Set rstCount = Nothing
    End If
    gIsSiteXDStation = blRet
    Exit Function
ErrHandler:
    gHandleError "", "gIsSiteXDStation"
    blRet = False
    GoTo Cleanup
    
End Function

Public Function gXmlIniPath(Optional blTestExistence As Boolean = False) As String
'dan M 11/01/2010
'search data folder, then exe folder; if not there, assume it's in sgDirectory folder. Use blTestExistence if don't want to assume, and return "" if doesn't exist.
    Dim oMyFileObj As FileSystemObject
    Dim slIniPath As String
    Set oMyFileObj = New FileSystemObject
    '4/14/12: Changed order of search to: Startup; database; exe From: database; exe; startup
    '         the reason for the change to to handle the case were two xml.ini definded and one placed
    '         in the database and the other placed into an ini folder.
'    slIniPath = oMyFileObj.BuildPath(sgDBPath, "xml.ini")
'    If oMyFileObj.FileExists(slIniPath) Then
'        gXmlIniPath = slIniPath
'    Else
'        slIniPath = oMyFileObj.BuildPath(sgExeDirectory, "xml.ini")
'        If oMyFileObj.FileExists(slIniPath) Then
'            gXmlIniPath = slIniPath
'        Else
'            slIniPath = oMyFileObj.BuildPath(sgStartupDirectory, "xml.ini")
'            If blTestExistence Then
'                If Not oMyFileObj.FileExists(slIniPath) Then
'                    slIniPath = vbNullString
'                End If
'            End If
'            gXmlIniPath = slIniPath
'        End If
'    End If
    slIniPath = oMyFileObj.BuildPath(sgStartupDirectory, "xml.ini")
    If oMyFileObj.FILEEXISTS(slIniPath) Then
        gXmlIniPath = slIniPath
    Else
        slIniPath = oMyFileObj.BuildPath(sgDBPath, "xml.ini")
        If oMyFileObj.FILEEXISTS(slIniPath) Then
            gXmlIniPath = slIniPath
        Else
            slIniPath = oMyFileObj.BuildPath(sgExeDirectory, "xml.ini")
            If Not oMyFileObj.FILEEXISTS(slIniPath) Then
                slIniPath = vbNullString
            End If
            gXmlIniPath = slIniPath
        End If
    End If
    Set oMyFileObj = Nothing
    
End Function

Public Function gXMLNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilStartPos As Integer
    Dim ilFound As Integer
    
    slName = slInName
    'Remove " and '
    ilStartPos = 1
    Do
        ilFound = False
        ilPos = InStr(ilStartPos, slName, "&", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&amp;" & Mid$(slName, ilPos + 1)
            ilStartPos = ilPos + Len("&amp;")
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&lt;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&gt;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&apos;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&quot;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    gXMLNameFilter = Trim$(slName)
End Function
Public Function gLoadFromIni(slSection As String, slKey As String, slPath As String, slValue As String) As Boolean
    'get values from ini file.
    'I-slSection ("XDigital"),  slKey("Host"), slPath (path to ini file)
    'o-slValue. Value from ini file. If not found, = "Not Found"
    'o- boolean.  True if found.
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    
    slValue = "Not Found"
    gLoadFromIni = False
    '8886
    'If Dir(slPath) > "" Then
    If gFileExist(slPath) = FILEEXISTS Then
        BytesCopied = GetPrivateProfileString(slSection, slKey, "Not Found", sBuffer, 128, slPath)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                slValue = Left(sBuffer, BytesCopied)
                gLoadFromIni = True
            End If
        End If
    End If 'slPath not valid?
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function

Public Function gIsNull(slXmlStatus As String) As Boolean
'5896
    Dim blRet As Boolean
    Dim slNull As String
    
    blRet = True
    If LenB(Trim$(slXmlStatus)) > 0 Then
        slNull = Mid(slXmlStatus, 1, 1)
        If slNull <> Chr(0) Then
            blRet = False
        End If
    End If
    gIsNull = blRet
End Function
Public Function gGetXmlErrorFile(slError As String) As String
'6807 send to Jeff
    'return slError for why not returning.
    Dim slRet As String
        
    slRet = ""
    slError = ""
On Error GoTo ERRORBOX
    slRet = sgMsgDirectory & gGetComputerName() & sgUserName & XMLERRORFILE
    gGetXmlErrorFile = slRet
    Exit Function
ERRORBOX:
    slError = Err.Description
    gGetXmlErrorFile = ""
End Function
Public Function gDeleteFile(slFile As String) As Boolean
    '6808 error on kill is that it's read only
    Dim ilRet As Integer
    Dim slDateTime As String
    
    gDeleteFile = True
    'On Error GoTo ERREXIST
    ilRet = 0
    'slDateTime = FileDateTime(slFile)
    ilRet = gFileExist(slFile)
    If ilRet = 0 Then
        On Error GoTo ERRFILE
        Kill slFile
    End If
    Exit Function
'ERREXIST:
'    ilRet = 1
'    Resume Next
ERRFILE:
    gDeleteFile = False
End Function

Public Function gParseXml(sllines As String, slName As String, llCurrentRow As Long) As String
    Dim slStartElement As String
    Dim slEndElement As String
    Dim slValue As String
    Dim ilPos As Long
    Dim ilEndPos As Long
    Dim ilLength As Integer
    Dim ilStart As Long
    
    If llCurrentRow = 0 Then
        llCurrentRow = 1
    End If
    slStartElement = "<" & slName & ">"
    slEndElement = "</" & slName & ">"
    ilPos = InStr(llCurrentRow, sllines, slStartElement)
    If ilPos > 0 Then
        ilStart = ilPos + Len(slStartElement)
        ilEndPos = InStr(llCurrentRow, sllines, slEndElement)
        ilLength = ilEndPos - ilStart
        slValue = Mid(sllines, ilStart, ilLength)
        gParseXml = gUnencodeXmlData(slValue)
    Else
        slStartElement = "<" & slName
        slEndElement = ">"
        ilPos = InStr(llCurrentRow, sllines, slStartElement)
        If ilPos > 0 Then
            ilStart = ilPos + Len(slStartElement)
            ilEndPos = InStr(llCurrentRow + ilPos, sllines, slEndElement)
            ilLength = ilEndPos - ilStart
            slValue = Mid(sllines, ilStart, ilLength)
            gParseXml = gUnencodeXmlData(slValue)
        Else
            gParseXml = vbNullString
        End If
    End If
End Function
Private Function gUnencodeXmlData(slData As String) As String
    Dim slRet As String
    If InStr(1, slData, "&") > 0 Then
        slRet = Replace(slData, "&lt;", "<")
        slRet = Replace(slRet, "&gt;", ">")
        slRet = Replace(slRet, "&amp;", "&")
        slRet = Replace(slRet, "&apos;", "`")
        slRet = Replace(slRet, "&quot;", """")
        gUnencodeXmlData = slRet
    Else
        gUnencodeXmlData = slData
    End If
End Function

