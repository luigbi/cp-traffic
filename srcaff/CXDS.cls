VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dan todo:  fix mclear and going into xml.ini
'depends on gsetpathendslash, and getinivalues full of dependencies
' use mSetSoap before 'sending'
Dim myService As CWebService
Private smErrorMessage As String
Private smUrl As String
Private smRet As String
'7539 to help testing
Private smImportPath As String
'Private smGuids() As String
Private ImPortNumber As Integer

Private smSecurityToken As String
Private smTransmissionID As String
Private smStartDate As String
Private smEndDate As String
Private smUnitID As String
Private smSiteID As String
Private smTestReturnValue As String
Public ImportMethod As XDSImportMethod
Public UseSecure As Boolean
Public ImportPage As String
Public Enum XDSImportMethod
    NoOffAirMonitoring 'TrackEvents
    OffAirMonitoring 'TrackRegionalEvents
    AsPlayed  'TrackPlayEvents
End Enum
'properties that 'pass-through'
'NameSpace
'LogPath
'isLog
'isTest
Public Property Let Port(ilValue As Integer)
    ImPortNumber = ilValue
End Property
Public Property Let ImportPath(slvalue As String)
    smImportPath = gSetPathEndSlash(slvalue, False)
End Property
Public Property Let isTest(blValue As Boolean)
    myService.isTest = blValue
End Property
Public Property Get isTest() As Boolean
    isTest = myService.isTest
End Property
Public Property Get ReturnValue() As String
    ReturnValue = smRet
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = smErrorMessage
End Property
Public Property Get NameSpace() As String
    NameSpace = myService.NameSpace
End Property
Public Property Let LogPath(slvalue As String)
    myService.LogPath = slvalue
End Property
Public Property Get isLog() As Boolean
    isLog = myService.isLog
End Property
Public Property Get SoapUrl() As String
    SoapUrl = smUrl
End Property
Public Property Let SoapUrl(slUrl As String)
    smErrorMessage = ""
    '7539 add later
    smUrl = slUrl
End Property
Public Property Let SecurityToken(slvalue As String)
    smSecurityToken = slvalue
End Property
Public Property Let TransmissionID(slvalue As String)
    smTransmissionID = slvalue
End Property
Public Property Let UnitID(slvalue As String)
    smUnitID = slvalue
End Property
Public Property Let SiteID(slvalue As String)
    smSiteID = slvalue
End Property
Public Property Let startDate(slDate As String)
    'make safe
    smStartDate = slDate
End Property
Public Property Let endDate(slDate As String)
    'make safe
    smEndDate = slDate
End Property
Public Property Let TestReturnValue(slvalue As String)
    smTestReturnValue = slvalue
End Property
Private Sub Class_Initialize()
    Set myService = New CWebService
    With myService
        .NameSpace = "xnms.xdigital.com/"
    End With
    Clear True
End Sub
Private Sub Class_Terminate()
    Set myService = Nothing
'    If Not myRsStations Is Nothing Then
'        If (myRsStations.State And adStateOpen) <> 0 Then
'            myRsStations.Close
'        End If
'        Set myRsStations = Nothing
'    End If
'    Erase smGuids
End Sub

Public Function CreateLogName(slPathAndName As String)
    CreateLogName = myService.CreateLogName(slPathAndName)
End Function
Public Sub Clear(blAll As Boolean)
    smErrorMessage = ""
    smStartDate = ""
    smEndDate = ""
    smUnitID = ""
    smSiteID = ""
    ImportMethod = OffAirMonitoring
    If blAll Then
        ImportPage = ""
        smImportPath = ""
        ImPortNumber = 80
        UseSecure = False
        smSecurityToken = ""
        smTransmissionID = ""
    End If
End Sub
Public Sub LogStart()
    'simple message, and tests to see if need new file
    myService.LogStart
End Sub
Public Sub LogEnd()
    myService.LogEnd
End Sub
Public Function Proxy(slUrl As String, slPort As String, blUseSecure As Boolean, slTestProxyUrl As String) As Boolean
    Dim blRet As Boolean
    
    blRet = True
    myService.SecureProxy = blUseSecure
    myService.Proxy slUrl, slPort
    If Len(slTestProxyUrl) > 0 Then
        blRet = myService.TestConnection("http://" & slTestProxyUrl)
    End If
    Proxy = blRet
End Function
Public Function ImportTimes(slFileName As String) As Boolean
    Dim blRet As Boolean
    Dim ilRet As Integer
    Dim slTemp As String
    Dim slTempMore As String
    Dim hlImportFile As Integer
    Dim slXmlValue As String
    Dim slMethod As String
    Dim slUseThisReturn As String
    Dim ilPos As Integer
    
    blRet = False
    If mIsComplete() Then
        If Len(smImportPath) > 0 Then
            mSetSoap ImportPage
    On Error GoTo ERRORBOX
            Select Case ImportMethod
                Case XDSImportMethod.AsPlayed
                    slMethod = "TrackPlayEvents"
                Case XDSImportMethod.NoOffAirMonitoring
                    slMethod = "TrackEvents"
                Case Else
                    slMethod = "TrackRegionalEvents"
            End Select
            With myService
                .MethodName = slMethod
                slTemp = "<SecurityToken TransmissionID='225'>" & smSecurityToken & "</SecurityToken>"
                If Len(smTransmissionID) > 0 Then
                    slTempMore = .BuildXmlElement("TransmissionID", smTransmissionID)
                End If
                If Len(smUnitID) > 0 Then
                    slTempMore = slTempMore & .BuildXmlElement("UnitID", smUnitID)
                End If
                If Len(smSiteID) > 0 Then
                    slTempMore = slTempMore & .BuildXmlElement("SiteID", smSiteID)
                End If
                If Len(smStartDate) > 0 Then
                    slTempMore = slTempMore & .BuildXmlElement("StartDate", smStartDate)
                End If
                If Len(smEndDate) > 0 Then
                    slTempMore = slTempMore & .BuildXmlElement("EndDate", smEndDate)
                End If
                slTemp = slTemp & .BuildXmlElement("filter", slTempMore)
                .Body = slTemp
            End With
            'what do I read as result?
            mSend "SiteID"
            If Len(smErrorMessage) = 0 Then
                If InStr(UCase(smRet), "TEST") > 0 Then
                    blRet = True
                    slUseThisReturn = smTestReturnValue
                    TestCreateReturnValue (2)
                Else
                    blRet = True
                    slUseThisReturn = myService.ReturnValue
                End If
            Else
                'not an error.  didn't return anything
                ilPos = InStr(1, myService.ReturnValue, "<Events />")
                If ilPos > 0 Then
                    blRet = True
                    smErrorMessage = ""
                    slUseThisReturn = myService.ReturnValue
                Else
                    smErrorMessage = myService.ReturnValue
                End If
            End If
            'now we parse and write
            If blRet Then
                If Not gDeleteFile(smImportPath & slFileName) Then
                    blRet = False
                    smErrorMessage = smErrorMessage & " Could not delete file '" & slFileName & "'"
                    GoTo Cleanup
                End If
                ilRet = 0
On Error GoTo OpenFileError
                'hlImportFile = FreeFile
                'Open smImportPath & slFileName For Output Lock Write As hlImportFile
                ilRet = gFileOpen(smImportPath & slFileName, "Output Lock Write", hlImportFile)
                If ilRet <> 0 Then
                    smErrorMessage = smErrorMessage & " Could not create file '" & slFileName & "'"
                    Close #hlImportFile
                    blRet = False
                End If
                Print #hlImportFile, slUseThisReturn
                Close hlImportFile
            End If
        Else
            smErrorMessage = "ImportPath cannot be empty."
        End If
    End If
Cleanup:
    ImportTimes = blRet
    Exit Function
OpenFileError:
    ilRet = Err
    Resume Next
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Problem in 'ImportTimes': " & Err.Description
    GoTo Cleanup
End Function
Public Function GetIniValues() As Boolean
    'dual provider?  must have section [BackOffice] with all values correct
    Dim blRet As Boolean
    Dim slIniPath As String
    Dim slRet As String
    Dim slSection As String
    Dim slProxyUrl As String
    Dim slProxyPort As String
    Dim slProxyTestUrl As String
    Dim blUseSecure As Boolean
    Dim blUseProxySecure As Boolean
    Dim slErrorTemp As String
    
On Error GoTo ERRORBOX
    blRet = False
    slProxyUrl = ""
    slProxyPort = ""
    slProxyTestUrl = ""
    blUseSecure = False
    blUseProxySecure = False
    slIniPath = gXmlIniPath(True)
    If LenB(slIniPath) = 0 Then
        slErrorTemp = "Xml.ini doesn't exit. 'Import XDS' cannot be activated."
        GoTo Cleanup
    End If
    'go through hoops because of possible 'dual provider'
        'this is importpage
    slSection = "XDigital"
    gLoadFromIni slSection, "WebServiceURL", slIniPath, slRet
    If slRet = "Not Found" Then
        slSection = "BackOffice"
        gLoadFromIni slSection, "WebServiceURL", slIniPath, slRet
        If slRet = "Not Found" Then
            slErrorTemp = "Cannot find 'WebServiceURL' in Xml.ini'  Make sure section [XDigital] or sectin [BackOffice] exists. 'Import XDS' cannot be activated."
            GoTo Cleanup
        End If
    End If
    ImportPage = slRet
    gLoadFromIni slSection, "ProxyServer", slIniPath, slRet
    If slRet <> "Not Found" Then
        slProxyUrl = slRet
        'must have port defined also
        gLoadFromIni slSection, "ProxyPort", slIniPath, slRet
        If slRet <> "Not Found" Then
            slProxyPort = slRet
            gLoadFromIni slSection, "ProxyTestURL", slIniPath, slRet
            If slRet <> "Not Found" Then
                slProxyTestUrl = slRet
            End If
            gLoadFromIni slSection, "UseSecureProxy", slIniPath, slRet
            If UCase(slRet) = "TRUE" Then
                blUseProxySecure = True
            End If
        Else
            slErrorTemp = "Xml.ini Cannot use 'ProxyServer' without also using 'ProxyPort'. 'Import XDS' cannot be activated."
            GoTo Cleanup
        End If
        If Not Proxy(slProxyUrl, slProxyPort, blUseProxySecure, slProxyTestUrl) Then
            slErrorTemp = "Couldn't use defined proxy in xml.ini. 'Import XDS' cannot be activated."
            GoTo Cleanup
        End If
    End If
    gLoadFromIni slSection, "UseSecure", slIniPath, slRet
    If UCase(slRet) = "TRUE" Then
        UseSecure = True
    End If
    gLoadFromIni slSection, "Host", slIniPath, slRet
    If slRet = "Not Found" Or Len(slRet) = 0 Then
        slErrorTemp = "Couldn't find 'Host' in xml.ini. 'Import XDS' cannot be activated."
        GoTo Cleanup
    End If
    smUrl = slRet
    'Check for test mode
    If StrComp(smUrl, "Test", vbTextCompare) = 0 Then
        myService.isTest = True
    End If
    gLoadFromIni slSection, "SecurityToken", slIniPath, slRet
    If slRet = "Not Found" Or Len(slRet) = 0 Then
        slErrorTemp = "Couldn't read 'SecurityToken' in xml.ini-" & slSection & ". 'Import XDS' cannot be activated."
        GoTo Cleanup
    End If
    smSecurityToken = slRet
    blRet = True
Cleanup:
    If Not blRet Then
        Clear True
        smErrorMessage = slErrorTemp
    End If
    GetIniValues = blRet
    Exit Function
ERRORBOX:
    blRet = False
    slErrorTemp = slErrorTemp & " Problem in 'GetIniValues': " & Err.Description
    GoTo Cleanup
End Function
Private Function mIsComplete() As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    If Len(smUrl) = 0 Then
        smErrorMessage = smErrorMessage & "Soap Url cannot be empty."
        blOk = False
    End If
    If Len(smSecurityToken) = 0 Then
        smErrorMessage = smErrorMessage & "Security token cannot be empty."
        blOk = False
    End If
    If Len(smTransmissionID) = 0 And Len(smSiteID) = 0 And Len(smUnitID) = 0 And Len(smStartDate) = 0 And Len(smEndDate) = 0 Then
        smErrorMessage = smErrorMessage & "At least one 'filter' must be created."
        blOk = False
    End If
    mIsComplete = blOk
End Function
Private Function mUrlSafe(slUrl As String) As Boolean
    Dim blRet As Boolean
    
    blRet = False
    If Len(slUrl) > 0 Then
        If InStr(1, slUrl, "ttp:", vbTextCompare) > 0 Then 'And InStr(1, slUrl, "mx/orders") > 0 Then
            blRet = True
        ElseIf UCase(slUrl) = "TEST" Then
            blRet = True
        End If
    End If
    mUrlSafe = blRet
End Function
Private Sub mSetSoap(slPage As String)
    Dim slStart As String
    
    If UseSecure Then
        slStart = "https://"
    Else
        slStart = "http://"
    End If
    myService.WebServiceAddress = slStart & smUrl & ":" & ImPortNumber & slPage
End Sub

Private Sub mSend(slReturnField As String)
    Dim slvalue As String
    
    With myService
        If .Send() Then
            If .isTest Then
                slvalue = .ReturnValue
            Else
                slvalue = .Read(slReturnField)
            End If
            If Len(slvalue) = 0 Then
                smErrorMessage = .ErrorMessage
                smRet = ""
            Else
                smRet = slvalue
            End If
        Else
            smErrorMessage = .ErrorMessage
        End If
    End With
End Sub

Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function

Private Function mUnencodeXmlData(slData As String) As String
    Dim slRet As String
    If InStr(1, slData, "&") > 0 Then
        slRet = Replace(slData, "&lt;", "<")
        slRet = Replace(slRet, "&gt;", ">")
        slRet = Replace(slRet, "&amp;", "&")
        slRet = Replace(slRet, "&apos;", "`")
        slRet = Replace(slRet, "&quot;", """")
        mUnencodeXmlData = slRet
    Else
        mUnencodeXmlData = slData
    End If
End Function
Public Sub TestCreateReturnValue(NumberEvents As Integer, Optional UnitOrSerial As String, Optional Program As String = "", Optional Site As String = "", Optional Cue As String = "")
    Dim slRet As String
    Dim slTemp As String
    Dim strEnvelope As String
    Dim slMethod As String
    Dim c As Integer
    
    Select Case ImportMethod
        Case XDSImportMethod.AsPlayed
            slMethod = "TrackPlayEvents"
        Case XDSImportMethod.NoOffAirMonitoring
            slMethod = "TrackEvents"
        Case Else
            slMethod = "TrackRegionalEvents"
    End Select
    If Len(UnitOrSerial) = 0 Then
        UnitOrSerial = 1
    End If
    If Len(Program) = 0 Then
        Program = 2
    End If
    If Len(Site) = 0 Then
        Site = 4
    End If
    If Len(Cue) = 0 Then
        Cue = 3
    End If
    strEnvelope = ""
    If NumberEvents > 0 Then
        slTemp = ""
        If ImportMethod = AsPlayed Then
            For c = 1 To NumberEvents
                slTemp = slTemp & "<Event SerialNumber = """ & UnitOrSerial & """ Port = ""portslot1""><ProgramNumber>" & Program & "</ProgramNumber><SiteID>" & Site & "</SiteID><Timestamp>12/14/1999 11:22:03 AM</Timestamp><Watermark>yes</Watermark></Event>"
            Next c
            slRet = myService.BuildXmlElement("Events", slTemp)
            slRet = myService.BuildXmlElement(slMethod & "Result", slRet)
            strEnvelope = "<?xml version=""1.0"" encoding = ""utf-8""?><soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            strEnvelope = strEnvelope & "<soap:Body><" & slMethod & "Response" & " xmlns=""" & myService.NameSpace & """ >"
            strEnvelope = strEnvelope & slRet & "</" & slMethod & "Response" & "></soap:Body></soap:Envelope>"
        Else
            For c = 0 To NumberEvents
                slTemp = slTemp & "<Event unitID=""" & UnitOrSerial & """><SiteID>" & Site & "</SiteID><ProgramCode>" & Program & "</ProgramCode><Cue>" & Cue & "</Cue><National>DickISCI</National><ISCI>DanISCI</ISCI><Status>Insert</Status>"
                slTemp = slTemp & "<AirTime>12/14/1999 11:22:03 AM</AirTime><ReportTime>12/14/1999 1:30:14 PM</ReportTime><SerialNumber>5</SerialNumber><Index>6</Index><Watermark>Yes</Watermark></Event>"
            Next c
            slRet = myService.BuildXmlElement("Events", slTemp)
            slRet = myService.BuildXmlElement(slMethod & "Result", slRet)
            strEnvelope = "<?xml version=""1.0"" encoding = ""utf-8""?><soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            strEnvelope = strEnvelope & "<soap:Body><" & slMethod & "Response" & " xmlns=""" & myService.NameSpace & """ >"
            strEnvelope = strEnvelope & slRet & "</" & slMethod & "Response" & "></soap:Body></soap:Envelope>"
        End If
    End If
    smTestReturnValue = strEnvelope
End Sub
'Private Sub mFillGuidArray()
'    myService.ReadMultiple "AffidavitGUID", smGuids
'End Sub

