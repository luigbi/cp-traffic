VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMarketron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'must fix mtestconnection, namespace, urlsafe
'I use smUrl differently for Marketron, because there are different web service pages ( 2 for exports and 1 for imports)
' use mSetSoap before 'sending'
Dim myService As CWebService
Private smErrorMessage As String
Private bmIsConnected As Boolean
Private smUrl As String
Private smRet As String
'7539 to help testing
Private bmGetAll As Boolean
Private smImportPath As String
Private smGuids() As String
Private imPort As Integer

Public UserName As String
Public Password As String
'7539
Public UseSecure As Boolean
'Dim smCallLetters As Dictionary
'Dim smServicePage
' exportPage not in current xml.ini. Otherwise, export and import are different.
' the idea is to get from ini if possible.  If don't exist, use these constants.
Private Const cEXPORTPAGESTATION As String = "/mx/orders/OrderServices.asmx"
Private Const cEXPORTPAGE As String = "/mx/orders/2.0/ProcessOrderNetwork.asmx"
Private Const cIMPORTPAGE As String = "/mx/orders/2.0/NetworkOrderServices.asmx"
Public ExportPage As String
Public ImportPage As String
'for Dick and getorderidbyorderid
Public fileName As String
Dim myRsStations As ADODB.Recordset
'properties that 'pass-through'
'NameSpace
'LogPath
'isLog
'isTest
Public Property Let Port(ilValue As Integer)
    imPort = ilValue
End Property
Public Property Let ImportPath(slvalue As String)
    'assumes "\" as last character
    smImportPath = slvalue
End Property
Public Property Let ReturnAll(blValue As Boolean)
    bmGetAll = blValue
End Property
Public Property Get ReturnAll() As Boolean
    ReturnAll = bmGetAll
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
    'Marketron needs this added
'    smUrl = "http://" & slUrl
'    If mUrlSafe(slUrl) Then
'        smUrl = slUrl
'        mSetSoap
'    Else
'        smErrorMessage = "The url is not valid.  Export not sent."
'        smUrl = ""
'    End If
End Property
'Public Property Get IsConnected() As Boolean
'    IsConnected = mTestConnection()
'End Property
Private Sub Class_Initialize()
    Clear
    Set myService = New CWebService
    With myService
        .NameSpace = "www.Marketron.com/webservices"
    End With
End Sub
Private Sub Class_Terminate()
    Set myService = Nothing
    If Not myRsStations Is Nothing Then
        If (myRsStations.State And adStateOpen) <> 0 Then
            myRsStations.Close
        End If
        Set myRsStations = Nothing
    End If
    Erase smGuids
End Sub

Public Function CreateLogName(slPathAndName As String)
    CreateLogName = myService.CreateLogName(slPathAndName)
End Function
Public Sub Clear()
        smErrorMessage = ""
        UserName = ""
        Password = ""
        'Set smCallLetters = New Dictionary
        ImportPage = ""
        ExportPage = ""
        'smret = "" ?
        bmGetAll = False
        smImportPath = ""
        ReDim smGuids(0)
        imPort = 80
        UseSecure = False
        fileName = ""
End Sub
Public Sub LogStart()
    'simple message, and tests to see if need new file
    myService.LogStart
End Sub
Public Sub LogEnd()
    myService.LogEnd
End Sub
'Public Sub ClearStations()
'    smCallLetters.RemoveAll
'End Sub
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

Public Function GetOrders(Optional blAndDownload As Boolean = False) As Integer
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slTemp As String
    Dim slTempMore As String
    Dim slReturnAmount As String
    Dim c As Integer
    If bmGetAll Then
        slReturnAmount = "Received"
    Else
        slReturnAmount = "Available"
    End If
    If mIsComplete() Then
        mSetSoap ImportPage
On Error GoTo ERRORBOX
        With myService
            slTemp = .BuildXmlElement("Username", UserName) & .BuildXmlElement("Password", Password)
            .MethodName = "GetNetworkOrderList"
           ' slTempMore = "<Stations xsi:nil=""true"" /><Orders xsi:nil=""true"" /><OrderStartDate xsi:nil=""true"" /><OrderEndDate xsi:nil=""true"" /><OrderStates xsi:nil=""true"" /><AffidavitStates>" & .BuildXmlElement("AffidavitState", "Available") & "</AffidavitStates>"
            slTempMore = "<AffidavitStates>" & .BuildXmlElement("AffidavitState", slReturnAmount) & "</AffidavitStates>"
            slTemp = .BuildXmlElement("AuthenticationToken", slTemp) & slTempMore
            .Body = "<GetNetworkOrderListRequest xmlns = ""http://www.Marketron.com/schemas/"">" & slTemp & "</GetNetworkOrderListRequest>"
        End With
        mSend "Success"
        If Len(smErrorMessage) = 0 Then
            If UCase(smRet) = "TRUE" Then
                blRet = True
            'Dan 03/03/14
            Else
                smErrorMessage = myService.ReturnValue
            End If
        End If
        If blRet And Not myService.isTest Then
            ilRet = myService.CountElements("Order")
            If blAndDownload And ilRet > 0 Then
                mFillGuidArray
                For c = 0 To UBound(smGuids) - 1
                   ' GetOrder smGuids(c)
                    '10215
                    If Not GetOrder(smGuids(c)) Then
                        ilRet = 0
                        Exit For
                    End If
                Next c
            End If
        End If
    End If
Cleanup:
    GetOrders = ilRet
    Exit Function
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Problem in 'GetStations': " & Err.Description
    GoTo Cleanup
    
End Function
Public Function GetOrder(slGuid As String) As Boolean
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slTemp As String
    Dim slTempMore As String
    Dim slNewName As String
    Dim hlImportFile As Integer
    Dim slXmlValue As String
    
    If mIsComplete() Then
        If Len(smImportPath) > 0 Then
            mSetSoap ImportPage
    On Error GoTo ERRORBOX
            With myService
                slTemp = .BuildXmlElement("Username", UserName) & .BuildXmlElement("Password", Password)
                .MethodName = "GetNetworkOrderAffidavit"
                slTempMore = .BuildXmlElement("AffidavitGUID", slGuid)
                slTemp = .BuildXmlElement("AuthenticationToken", slTemp) & slTempMore
                .Body = "<GetNetworkOrderAffidavitRequest xmlns = ""http://www.Marketron.com/schemas/"">" & slTemp & "</GetNetworkOrderAffidavitRequest>"
            End With
            mSend "Result"
            If Len(smErrorMessage) = 0 Then
                If UCase(smRet) = "TRUE" Then
                    blRet = True
                Else
                    smErrorMessage = myService.ReturnValue
                End If
            End If
            'now we parse and write
            If blRet Then
                slXmlValue = myService.Read("OrderID")
                slXmlValue = mUnencodeXmlData(slXmlValue)
                'for dick's getorderbyorderid
                If Len(fileName) = 0 Then
                    slNewName = gFileNameFilter(slXmlValue)
                    slNewName = slNewName & ".txt"
                Else
                    slNewName = fileName
                End If
'                If Not gDeleteFile(smImportPath & slNewName) Then
'                    blRet = False
'                    smErrorMessage = smErrorMessage & " Could not delete file '" & slNewName & "'"
'                    GoTo Cleanup
'                End If
                ilRet = 0
On Error GoTo OpenFileError
                'hlImportFile = FreeFile
                'Open smImportPath & slNewName For Output Lock Write As hlImportFile
                ilRet = gFileOpen(smImportPath & slNewName, "Output Lock Write", hlImportFile)
                If ilRet <> 0 Then
                    smErrorMessage = smErrorMessage & " Could not create file '" & slNewName & "'"
                    Close #hlImportFile
                    blRet = False
                End If
                Print #hlImportFile, myService.ReturnValue
                Close hlImportFile
            End If
        Else
            smErrorMessage = "ImportPath cannot be empty."
        End If
    End If
Cleanup:
    GetOrder = blRet
    Exit Function
OpenFileError:
    ilRet = Err
    Resume Next
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Problem in 'GetOrder': " & Err.Description
    GoTo Cleanup
End Function
Public Function GetOrdersByOrderId(slOrderId As String, blReceivedOnly As Boolean) As Integer
    '
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slTemp As String
    Dim slTempMore As String
    Dim slReturnAmount As String
    Dim c As Integer
    If blReceivedOnly Then
        slReturnAmount = "Received"
    Else
        slReturnAmount = "Available"
    End If
    If mIsComplete() Then
        mSetSoap ImportPage
On Error GoTo ERRORBOX
        With myService
            slTemp = .BuildXmlElement("Username", UserName) & .BuildXmlElement("Password", Password)
            .MethodName = "GetNetworkOrderList"
            slTempMore = "<Orders>" & .BuildXmlElement("OrderID", slOrderId) & "</Orders>"
            slTempMore = slTempMore & "<AffidavitStates>" & .BuildXmlElement("AffidavitState", "Received")
            If Not blReceivedOnly Then
                slTempMore = slTempMore & .BuildXmlElement("AffidavitState", "Available")
            End If
            slTempMore = slTempMore & "</AffidavitStates>"
            slTemp = .BuildXmlElement("AuthenticationToken", slTemp) & slTempMore
            .Body = "<GetNetworkOrderListRequest xmlns = ""http://www.Marketron.com/schemas/"">" & slTemp & "</GetNetworkOrderListRequest>"
        End With
        mSend "Success"
        If Len(smErrorMessage) = 0 Then
            If UCase(smRet) = "TRUE" Then
                blRet = True
            'Dan 03/03/14
            Else
                smErrorMessage = myService.ReturnValue
            End If
        End If
        If blRet And Not myService.isTest Then
            ilRet = myService.CountElements("Order")
            If ilRet > 0 Then
                mFillGuidArray
                For c = 0 To UBound(smGuids) - 1
                    GetOrder smGuids(c)
                Next c
            End If
        End If
    End If
Cleanup:
    GetOrdersByOrderId = ilRet
    Exit Function
ERRORBOX:
    blRet = False
    smErrorMessage = smErrorMessage & " Problem in 'GetOrdersByOrderId': " & Err.Description
    GoTo Cleanup
    
End Function

'Public Function GetStations() As Boolean
'    Dim myXml As MSXML2.DOMDocument
'    Dim blRet As Boolean
'    Dim myNodeList As IXMLDOMNodeList
'    Dim MyNode As IXMLDOMNode
'    Dim slReturn As String
'    Dim slTemp As String
'
'    blRet = False
'    'already ran? then don't do again.
'    If Not myRsStations Is Nothing Then
'        myRsStations.Filter = adFilterNone
'        If myRsStations.RecordCount > 0 Then
'            GetStations = True
'            Exit Function
'        End If
'    End If
'    If mIsComplete() Then
'        mSetSoap cEXPORTPAGESTATION
'On Error GoTo ERRORBOX
'        With myService
'            slTemp = .BuildXmlElement("Username", UserName) & .BuildXmlElement("Password", Password)
'            .MethodName = "ListActiveStations"
'            slTemp = .BuildXmlElement("AuthenticationToken", slTemp)
'            .Body = "<ListActiveStationsRequest xmlns = ""http://www.Marketron.com/schemas/"">" & slTemp & "</ListActiveStationsRequest>"
'        End With
'        mSend "Result"
'        If Len(smErrorMessage) = 0 Then
'            If UCase(smRet) = "TRUE" Then
'                blRet = True
'            ElseIf myService.isTest Then
'                blRet = True
'            End If
'        End If
'        If blRet And Not myService.isTest Then
'            Set myXml = New MSXML2.DOMDocument
'            Set myRsStations = mPrepRecordset()
'            If Len(myService.ReturnValue) > 0 Then
'                If myXml.loadXML(myService.ReturnValue) Then
'                    blRet = True
'                    Set myNodeList = myXml.selectNodes("//ActiveStationData")
'                    For Each MyNode In myNodeList
'                        slReturn = MyNode.selectSingleNode("CallLetters").Text
'                        slReturn = slReturn & "-" & MyNode.selectSingleNode("Band").Text
'                        myRsStations.AddNew Array("StationName"), Array(UCase(slReturn))
'                    Next MyNode
'                Else
'                   smErrorMessage = smErrorMessage & myXml.parseError.reason
'                   blRet = False
'                   GoTo Cleanup
'                End If
'            End If
'        End If
'    End If
'Cleanup:
'    GetStations = blRet
'    Set myXml = Nothing
'    Set myNodeList = Nothing
'    Set MyNode = Nothing
'    Exit Function
'ERRORBOX:
'    blRet = False
'    smErrorMessage = smErrorMessage & " Problem in 'GetStations': " & Err.Description
'    GoTo Cleanup
'End Function
Public Function TestStation(slStation As String) As Boolean
    Dim blRet As Boolean
    
    blRet = False
    slStation = UCase(slStation)
    If Not myRsStations Is Nothing Then
        myRsStations.Filter = "StationName = '" & slStation & " '"
        If Not myRsStations.EOF Then
            blRet = True
        End If
    End If
    If myService.isTest Then
        blRet = True
    End If
    TestStation = blRet
End Function

'Public Function TestStation(slStation As String) As Boolean
'    Dim blRet As Boolean
'    Dim slCall As String
'    Dim slBand As String
'    Dim slTemp As String
'    Dim ilPos As Integer
'
'    blRet = False
'    mSetSoap cEXPORTPAGESTATION
'    If smCallLetters.Exists(slStation) Then
'        blRet = CBool(smCallLetters.Item(slStation))
'    Else
'        If mIsComplete() Then
'            ilPos = InStr(1, slStation, "-")
'            If ilPos > 0 Then
'                slCall = Mid(slStation, 1, ilPos - 1)
'                slBand = Mid(slStation, ilPos + 1)
'                With myService
'                    slTemp = .BuildXmlElement("Username", UserName) & .BuildXmlElement("Password", Password)
'                    .MethodName = "IsStationActive"
'                    slTemp = .BuildXmlElement("AuthenticationToken", slTemp) & .BuildXmlElement("CallLetters", slCall) & .BuildXmlElement("Band", slBand)
'                    .Body = "<IsStationActiveRequest xmlns = ""http://www.Marketron.com/schemas/"">" & slTemp & "</IsStationActiveRequest>"
'                End With
'                mSend "IsStationActive"
'                If Len(smErrorMessage) = 0 Then
'                    If smRet = "true" Then
'                        blRet = True
'                    ElseIf myService.isTest Then
'                        blRet = True
'                    End If
'                    smCallLetters.Add slStation, blRet
'                End If
'            Else
'                smErrorMessage = "Problem in TestStation: couldn't parse " & slStation
'            End If
'        End If
'    End If
'    TestStation = blRet
'End Function
Private Function mIsComplete() As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    If Len(smUrl) = 0 Then
        smErrorMessage = "Soap Url cannot be empty."
        blOk = False
    End If
    If Len(UserName) = 0 Then
        smErrorMessage = "Username cannot be empty."
        blOk = False
    End If
    If Len(Password) = 0 Then
        smErrorMessage = "Password cannot be empty."
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
    myService.WebServiceAddress = slStart & smUrl & ":" & imPort & slPage
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
Private Function mPrepRecordset() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "StationName", adChar, 40
        End With
    myRs.Open
    myRs!StationName.Properties("optimize") = True
    myRs.Sort = "StationName"
    Set mPrepRecordset = myRs
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
Private Sub mFillGuidArray()
    myService.ReadMultiple "AffidavitGUID", smGuids
End Sub
