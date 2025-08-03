VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEOExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' depends on: gXmlIniPath and more!

Dim myService As CWebService
Private smErrorMessage As String
Private smUrl As String
Private smRet As String
Private smUserName As String
Private smPassword As String
Private smSchemaNameSpace As String
'will be 1 more than used  8278 remove 512
Dim smBody() As String ' * 512
Dim lmBodyCount As Long
Dim lmBodyMultiplier As Long
Const ChunkSize As Integer = 50
'must match smBody above!
Const STRINGSIZE As Integer = 512
Const WEBSERVICEPAGE As String = "Orders/VendorOrderServices/VendorOrderServices.asmx"
'properties that 'pass-through'
'NameSpace
'LogPath
'isLog
'isTest
'isProxy
Public Property Get isProxy() As Boolean
    isProxy = myService.isProxy
End Property
Public Property Let isTest(blValue As Boolean)
    myService.isTest = blValue
End Property
Public Property Get isTest() As Boolean
    isTest = myService.isTest
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
'Public Property Let SoapUrl(slUrl As String)
'    smErrorMessage = ""
'    If mUrlSafe(slUrl) Then
'        smUrl = slUrl
'        mSetSoap
'    Else
'        smErrorMessage = "The url is not valid.  Export not sent."
'        smUrl = ""
'    End If
'End Property

Public Property Get IsStation(slStation As String, slBand As String) As Boolean
    IsStation = mTestStation(slStation, slBand)
End Property
Public Property Get UserName() As String
    UserName = smUserName
End Property
Public Property Let UserName(ByVal slUserName As String)
    smUserName = slUserName
End Property
Public Property Get Password() As String
    Password = smPassword
End Property
Public Property Let Password(ByVal slPassword As String)
    smPassword = slPassword
End Property

Private Sub Class_Initialize()
    Clear False
    Set myService = New CWebService
    With myService
        .NameSpace = "www.Marketron.com/webservices" '"datacast.com/PMWSDL"
    End With
    smSchemaNameSpace = "www.Marketron.com/schemas"
    If mGetXmlIniValues() Then
        mSetSoap
    End If
End Sub
Public Function CreateLogName(slPathAndName As String)
    CreateLogName = myService.CreateLogName(slPathAndName)
End Function
Public Sub Clear(blOrderOnly As Boolean)
    If Not blOrderOnly Then
        smErrorMessage = ""
        smUrl = ""
        smRet = ""
        smUserName = ""
        smPassword = ""
   End If
    ReDim smBody(0 To ChunkSize)
    lmBodyCount = 0
    lmBodyMultiplier = 1
End Sub
Public Function AddToBody(slAdd As String) As Boolean
    Dim blRet As Boolean
    '8278
    On Error GoTo ERRORBOX
    blRet = True
    If Len(slAdd) > STRINGSIZE Then
        blRet = False
    Else
        If lmBodyCount > ChunkSize * lmBodyMultiplier Then
            lmBodyMultiplier = lmBodyMultiplier + 1
            ReDim Preserve smBody(0 To ChunkSize * lmBodyMultiplier)
        End If
        smBody(lmBodyCount) = slAdd
        lmBodyCount = lmBodyCount + 1
    End If
    AddToBody = blRet
    Exit Function
ERRORBOX:
    smErrorMessage = "addToBody issue"
    ReDim smBody(0)
    lmBodyCount = 0
    lmBodyMultiplier = 0
    AddToBody = False
End Function
Public Function SendOrder() As Boolean
    Dim blRet As Boolean
    Dim llIndex As Long
    Dim llUpper As Long
    
    smErrorMessage = ""
    blRet = False
    If Len(smUrl) > 0 Then
        llUpper = lmBodyCount - 1
        If llUpper > -1 Then
            With myService
                .MethodName = "Submit804Order"
                .Body = mBuildAuthorization()
                If Len(.Body) > 0 Then
                    For llIndex = 0 To llUpper
                        .Body = .Body & Trim$(smBody(llIndex))
                    Next llIndex
                End If
                .Send
                If Len(smErrorMessage) = 0 Then
                    If IsNumeric(smRet) Then
                        If smRet <> "0" Then
                            blRet = True
                        End If
                    ElseIf myService.isTest Then
                        blRet = True
                    End If
                End If
            End With
        Else
            smErrorMessage = "must create the body before sending."
        End If
    Else
        smErrorMessage = "Url not set!"
    End If
    SendOrder = blRet
End Function

Public Function Proxy(slUrl As String, slPort As String, blUseSecure As Boolean, slTestProxyUrl As String) As Boolean
    Dim blRet As Boolean
    
    blRet = True
    myService.SecureProxy = blUseSecure
    myService.Proxy slUrl, slPort
    If myService.isProxy Then
        If Len(slTestProxyUrl) > 0 Then
            blRet = myService.TestConnection("http://" & slTestProxyUrl)
        End If
    Else
        blRet = False
    End If
    Proxy = blRet
End Function
Public Sub LogStart()
    'simple message, and tests to see if need new file
    myService.LogStart
End Sub
Public Sub LogEnd()
    myService.LogEnd
End Sub

'Private Function mUrlSafe(slUrl As String) As Boolean
'    Dim blRet As Boolean
'
'    blRet = False
'    If Len(slUrl) > 0 Then
'        If InStr(1, slUrl, "http", vbTextCompare) > 0 And InStr(1, slUrl, "VendorOrderServices") > 0 Then
'        blRet = True
'        ElseIf InStr(1, slUrl, "TEST", vbTextCompare) > 0 Then
'            blRet = True
'        End If
'    End If
'    mUrlSafe = blRet
'End Function
Private Sub mSetSoap()
    myService.WebServiceAddress = smUrl
End Sub
Private Function mTestStation(slStation As String, slBand As String) As Boolean
    Dim blRet As Boolean
    Dim slSend As String
    
    
    If Len(smUrl) > 0 Then
        'will set error message to "", test username and password and write error messages if not.
        slSend = mBuildAuthorization()
        If Len(slSend) > 0 Then
            With myService
                .MethodName = "IsStationActive"
                slSend = slSend & .BuildXmlElement("callLetters", slStation) & .BuildXmlElement("band", slBand) & .BuildXmlElement("forOrderType", "Network")
                .Body = slSend
            End With
            mSend "Success"
            If Len(smErrorMessage) = 0 Then
                If smRet = "Success" Then
                    blRet = True
                ElseIf myService.isTest Then
                    blRet = True
                End If
            End If
        End If
    Else
        smErrorMessage = "Url not set!"
    End If
    mTestStation = blRet
End Function
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
Private Function mBuildAuthorization() As String
    Dim slRet As String
    
    slRet = ""
    smErrorMessage = slRet
    If Len(smPassword) = 0 Then
        smErrorMessage = "Password not set"
    End If
    If Len(smUserName) = 0 Then
        smErrorMessage = smErrorMessage & " username not set"
    End If
    If Len(smErrorMessage) = 0 Then
        slRet = "<authorization><Password xmlns=""" & smSchemaNameSpace & """>" & smPassword & "</Password>"
        slRet = slRet & "<UserName xmlns=""" & smSchemaNameSpace & """>" & smUserName & "</Username></authorization>"
    End If
    mBuildAuthorization = slRet
End Function
Private Function mGetXmlIniValues() As Boolean
    Dim blRet As Boolean
    Dim slIniPath As String
    Dim slRet As String
    Dim slFound As String
    Dim slProxyUrl As String
    Dim slProxyPort As String
    Dim slProxyTestUrl As String
    Dim blUseProxySecure As Boolean
    
    smUrl = ""
    blRet = False
    slFound = ""
    slRet = ""
    slProxyUrl = ""
    slProxyPort = ""
    slProxyTestUrl = ""
    blUseProxySecure = False
    slIniPath = gXmlIniPath(True)
    If LenB(slIniPath) > 0 Then
        gLoadFromIni STATIONXMLMARKETRONNAME, "Host", slIniPath, slRet
        If slRet = "Not Found" Then
            smErrorMessage = "Missing 'host' from 'Marketron-EO' Section of " & slIniPath
        Else
            If InStr(1, slRet, "TEST", vbTextCompare) = 1 Then
                Me.isTest = True
                slFound = UCase(slRet)
            ElseIf InStr(1, slRet, "HTTP", vbTextCompare) = 1 Then
                slFound = slRet
            Else
                slFound = slRet
                gLoadFromIni STATIONXMLMARKETRONNAME, "UseSecure", slIniPath, slRet
                If UCase(slRet) = "TRUE" Then
                    slFound = "https://" & slFound
                Else
                    slFound = "http://" & slFound
                End If
            End If
            slRet = ""
            gLoadFromIni STATIONXMLMARKETRONNAME, "Port", slIniPath, slRet
            If slRet = "Not Found" Then
                slFound = slFound & ":80"
            Else
                If IsNumeric(slRet) Then
                    slFound = slFound & ":" & slRet
                Else
                    smErrorMessage = "Port is not a number:" & slRet
                    slFound = ""
                End If
            End If
            'username and password
            If Len(slFound) > 0 Then
                gLoadFromIni STATIONXMLMARKETRONNAME, "Username", slIniPath, slRet
                If slRet = "Not Found" Then
                    smErrorMessage = "Missing 'Username' from 'EO' Section of " & slIniPath
                    slFound = ""
                Else
                    Me.UserName = slRet
                    gLoadFromIni STATIONXMLMARKETRONNAME, "Password", slIniPath, slRet
                    If slRet = "Not Found" Then
                        smErrorMessage = "Missing 'Password' from 'EO' Section of " & slIniPath
                        slFound = ""
                    Else
                        Me.Password = slRet
                    End If
                End If
            End If 'username and password
            If Len(slFound) > 0 Then
                'test?
                blRet = True
                gLoadFromIni STATIONXMLMARKETRONNAME, "GenerateFileOnly", slIniPath, slRet
                If UCase(slRet) = "TRUE" Then
                    Me.isTest = True
                Else
                    If InStrRev(slFound, "/") = Len(slFound) Then
                        slFound = slFound & WEBSERVICEPAGE
                    Else
                        slFound = slFound & "/" & WEBSERVICEPAGE
                    End If
                End If
'                gLoadFromIni "EO", "WebServiceURL", slIniPath, slRet
'                If slRet = "Not Found" Then
'                    slFound = ""
'                    smErrorMessage = "Missing 'WebServiceURL' from 'EO' Section of " & slIniPath
'                Else
'                    blRet = True
'                    If Me.isTest = False Then
'                        If InStr(1, slRet, "/") = 1 Or InStrRev(slFound, "/") = Len(slFound) Then
'                            slFound = slFound & slRet
'                        Else
'                            slFound = slFound & "/" & slRet
'                        End If
'                    End If
'                End If
                'proxy
                gLoadFromIni STATIONXMLMARKETRONNAME, "ProxyServer", slIniPath, slRet
                If slRet <> "Not Found" Then
                    slProxyUrl = slRet
                    'must have port defined also
                    gLoadFromIni STATIONXMLMARKETRONNAME, "ProxyPort", slIniPath, slRet
                    If slRet <> "Not Found" Then
                        slProxyPort = slRet
                        gLoadFromIni STATIONXMLMARKETRONNAME, "ProxyTestURL", slIniPath, slRet
                        If slRet <> "Not Found" Then
                            slProxyTestUrl = slRet
                        End If
                        gLoadFromIni STATIONXMLMARKETRONNAME, "UseSecureProxy", slIniPath, slRet
                        If UCase(slRet) = "TRUE" Then
                            blUseProxySecure = True
                        End If
                    Else
                        slProxyUrl = ""
                    End If
                    If Len(slProxyUrl) > 0 Then
                        If Not Me.Proxy(slProxyUrl, slProxyPort, blUseProxySecure, slProxyTestUrl) Then
                            smErrorMessage = "Proxy not set up properly"
                        End If
                    End If
                End If 'proxy
            End If 'port
        End If 'url
    End If 'ini path
    smUrl = slFound
    mGetXmlIniValues = blRet
End Function
'Private Sub mSendXD(slReturnField As String)
'    Dim slvalue As String
'
'    With myService
'        If .SendSpecial1() Then
'            If .isTest Then
'                slvalue = .ReturnValue
'            Else
'                slvalue = .Read(slReturnField)
'            End If
'            If Len(slvalue) = 0 Then
'                smErrorMessage = .ErrorMessage
'                smRet = ""
'            Else
'                smRet = slvalue
'            End If
'        Else
'            smErrorMessage = .ErrorMessage
'        End If
'    End With
'End Sub

