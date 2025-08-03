VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'if going to send to install, use 6.0.  ie6.0 used 3.5
Public MethodName As String
Public WebServiceAddress As String
Public Body As String
'what goes next to method name?
Public NameSpace As String
'only create log
Public isTest As Boolean
Private smReturn As String
Private smError As String
Private bmSent As Boolean
Private myLogger As CLogger
'7539
Private objHttp As MSXML2.ServerXMLHTTP60
Public SecureProxy As Boolean
Private bmIsProxy As Boolean

Private Sub Class_Initialize()
    Clear True, True
    Set myLogger = New CLogger
End Sub
Public Property Get isProxy() As Boolean
    isProxy = bmIsProxy
End Property
Public Property Get isLog() As Boolean
    isLog = myLogger.isLog
End Property
Public Property Get LogPath() As String
    LogPath = myLogger.LogPath
End Property
Public Property Let LogPath(slvalue As String)
    myLogger.LogPath = slvalue
End Property
Public Property Get ReturnValue() As String
    ReturnValue = smReturn
End Property
Public Property Get ErrorMessage() As Variant
    ErrorMessage = smError
End Property
Public Function CreateLogName(slPathAndName As String)
    CreateLogName = myLogger.CreateLogName(slPathAndName)
End Function
Public Sub Clear(blURL As Boolean, blError As Boolean)
    MethodName = ""
    Body = ""
    smReturn = ""
    isTest = False
    bmSent = False
    If blURL Then
        WebServiceAddress = ""
        Set objHttp = New ServerXMLHTTP60
        SecureProxy = False
        bmIsProxy = False
    End If
    If blError Then
        smError = ""
    End If
End Sub
Public Function Send() As Boolean
    Dim blRet As Boolean
   ' Dim objHttp As MSXML2.XMLHTTP60
    Dim strEnvelope As String
    Dim slRet As String
    Dim slNameSpace As String
On Error GoTo ERRORBOX
    blRet = True
    If Len(WebServiceAddress) = 0 Then
        blRet = False
        smError = "Missing Web Service Address"
    ElseIf Len(MethodName) = 0 Then
        blRet = False
        smError = "Missing Method Name"
    End If
    slNameSpace = NameSpace
    If Len(slNameSpace) = 0 Then
        slNameSpace = "tempuri.org"
    End If
    slNameSpace = "http://" & slNameSpace
    If blRet Then
        If Len(Body) > 0 Then
            strEnvelope = ""
            strEnvelope = strEnvelope & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            strEnvelope = strEnvelope & "<soap:Body><" & MethodName & " xmlns=""" & slNameSpace & """ >" '
            strEnvelope = strEnvelope & Body & "</" & MethodName & "></soap:Body></soap:Envelope>"
        Else
            strEnvelope = "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/xmlSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            strEnvelope = strEnvelope & "<soap:Body><" & MethodName & " xmlns=""" & slNameSpace & """ />" '
            strEnvelope = strEnvelope & "</soap:Body></soap:Envelope>"
        End If
        If isTest Then
            smReturn = "Test mode.  No return messages."
        Else
           ' Set objHttp = New MSXML2.XMLHTTP60
           '7714
            objHttp.Open "POST", WebServiceAddress, False
            objHttp.setRequestHeader "Content-Type", "text/xml"
            objHttp.setRequestHeader "SOAPMethodName", slNameSpace & "/" & MethodName
            objHttp.Send strEnvelope
            smReturn = objHttp.responseText
        End If
        If isLog Then
            myLogger.WriteXml strEnvelope, smReturn
        End If
    End If
    Send = blRet
Cleanup:
    bmSent = blRet
'    Set objHttp = Nothing
    Exit Function
ERRORBOX:
    Send = False
    smError = "Could not connect."
    GoTo Cleanup
End Function
Public Function Read(slElement As String) As String
    '7878
    Dim objReturn As Object ' As MSXML2.DOMDocument30
    Dim objElem As Object 'MSXML2.IXMLDOMElement
    
    smError = ""
    If Not bmSent Then
        Read = ""
        smError = "Cannot read until perform a 'send'"
        Exit Function
    End If
    'Set objReturn = New MSXML2.DOMDocument30
    Set objReturn = CreateObject("MSXML2.DOMDocument")
    objReturn.loadXML smReturn
    Set objElem = objReturn.selectSingleNode("//" & slElement)
    If Not objElem Is Nothing Then
        Read = objElem.Text
    Else
        Read = ""
        Set objElem = objReturn.selectSingleNode("//faultstring")
        If Not objElem Is Nothing Then
            smError = objElem.Text
        ElseIf InStr(1, smReturn, "Bad Request") > 0 Then
            smError = "Bad Request.  Xml improperly formatted."
        Else
            smError = "Could not read element '" & slElement & "' in " & smReturn
        End If
    End If
    Set objReturn = Nothing
    Set objElem = Nothing
End Function
Public Function BuildXmlElement(slElementName As String, slvalue As String) As String
    BuildXmlElement = "<" & slElementName & ">" & slvalue & "</" & slElementName & ">"
End Function
Public Function FormatDate(ByVal slDate As String) As String
    Dim slNewDate As String
    
    slNewDate = Format(slDate, "yyyy-mm-dd")
    FormatDate = slNewDate
End Function
Public Sub LogStart()
    Dim slTest As String
    
    If isTest Then
        slTest = " TEST MODE "
    End If
    myLogger.WriteFacts ":::" & Now() & " Begin Session to " & WebServiceAddress & slTest & ":::" & vbCrLf
End Sub
Public Sub LogEnd()
    myLogger.WriteFacts ":::" & Now() & " End Session :::"
End Sub
Public Function CountElements(slElement As String) As Integer
    '7878
    Dim objReturn As Object ' MSXML2.DOMDocument30
    Dim myNodeList As Object 'IXMLDOMNodeList
    Dim MyNode As Object 'IXMLDOMNode
    Dim ilCount As Integer
    
    ilCount = 0
    smError = ""
    If Not bmSent Then
        CountElements = 0
        smError = "Cannot count until perform a 'send'"
        Exit Function
    End If
On Error GoTo ERRORBOX
    'Set objReturn = New MSXML2.DOMDocument30
    Set objReturn = CreateObject("MSXML2.DOMDocument")
    objReturn.loadXML smReturn
    Set myNodeList = objReturn.selectNodes("//" & slElement)
    For Each MyNode In myNodeList
        ilCount = ilCount + 1
    Next MyNode
Cleanup:
    Set objReturn = Nothing
    Set myNodeList = Nothing
    Set MyNode = Nothing
    CountElements = ilCount
    Exit Function
ERRORBOX:
    smError = Err.Description
    ilCount = 0
    GoTo Cleanup
End Function
Public Sub ReadMultiple(slElement As String, slGuids() As String)
    '7878
    Dim objReturn As Object ' MSXML2.DOMDocument30
    Dim myNodeList As Object ' IXMLDOMNodeList
    Dim MyNode As Object 'IXMLDOMNode
    Dim ilUpper As Integer
    
    smError = ""
    If Not bmSent Then
        smError = "Cannot count until perform a 'send'"
        Exit Sub
    End If
On Error GoTo ERRORBOX
    'Set objReturn = New MSXML2.DOMDocument30
    Set objReturn = CreateObject("MSXML2.DOMDocument")
    objReturn.loadXML smReturn
    Set myNodeList = objReturn.selectNodes("//" & slElement)
    ilUpper = UBound(slGuids)
    For Each MyNode In myNodeList
        slGuids(ilUpper) = MyNode.Text
        ilUpper = ilUpper + 1
        ReDim Preserve slGuids(ilUpper)
    Next MyNode
Cleanup:
    Set objReturn = Nothing
    Set myNodeList = Nothing
    Set MyNode = Nothing
    Exit Sub
ERRORBOX:
    smError = Err.Description
    GoTo Cleanup
End Sub
Public Function SendSpecial1() As Boolean
    '7539 use the modular objHttp
    'special namespace, special soap 'env', special where namespace is
    Dim blRet As Boolean
    'Dim objHttp As MSXML2.XMLHTTP60
    Dim strEnvelope As String
    Dim slRet As String
    Dim slNameSpace As String
    
On Error GoTo ERRORBOX
    blRet = True
    If Len(WebServiceAddress) = 0 Then
        blRet = False
        smError = "Missing Web Service Address"
    ElseIf Len(MethodName) = 0 Then
        blRet = False
        smError = "Missing Method Name"
    End If
    slNameSpace = "Datacast.com"
    If blRet Then
        If Len(Body) > 0 Then
            strEnvelope = ""
            strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:dat=""http://DataCast.com"">"
            strEnvelope = strEnvelope & "<soapenv:Header/><soapenv:Body><" & MethodName & " soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"  '& " xmlns=""" & slNameSpace & ""
            strEnvelope = strEnvelope & Body & "</" & MethodName & "></soap:Body></soap:Envelope>"
        End If
        If isTest Then
            smReturn = "Test mode.  No return messages."
        Else
            'Set objHttp = New MSXML2.XMLHTTP60
            objHttp.Open "post", WebServiceAddress, False
            objHttp.setRequestHeader "Content-Type", "text/xml"
            objHttp.Send strEnvelope
            smReturn = objHttp.responseText
        End If
        If isLog Then
            myLogger.WriteXml strEnvelope, smReturn
        End If
    End If
    SendSpecial1 = blRet
Cleanup:
    bmSent = blRet
    'Set objHttp = Nothing
    Exit Function
ERRORBOX:
    SendSpecial1 = False
    smError = "Could not connect."
    GoTo Cleanup
End Function
Public Sub Proxy(slUrl As String, slPort As String)
    Dim slStart As String
    
On Error GoTo ERRBOX
    If SecureProxy Then
        slStart = "https://"
    Else
        slStart = "http://"
    End If
    objHttp.setProxy 2, slStart & slUrl & ":" & slPort, ""
    bmIsProxy = True
'this would test that this was right
'    objHttp.Open "POST", "http://www.Counterpoint.net"
'    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    objHttp.Send

    Exit Sub
ERRBOX:
    smError = "Could not set proxy"
    bmIsProxy = False
End Sub
Public Function TestConnection(slUrl As String) As Boolean
    
On Error GoTo ERRBOX
    objHttp.Open "POST", slUrl
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.Send
    TestConnection = True
    Exit Function
ERRBOX:
    TestConnection = False
End Function
Private Sub Class_Terminate()
    Set objHttp = Nothing
End Sub