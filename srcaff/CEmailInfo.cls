VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmailInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim smHost As String
Dim smPass As String
Dim smUser As String
Dim imPort As Integer
Dim bmTls As Boolean


Public Property Let Host(ByVal slvalue As String)
    smHost = slvalue
End Property

Public Property Get Host() As String
    Host = smHost
End Property

Public Property Let Password(ByVal slvalue As String)
    smPass = slvalue
End Property

Public Property Get Password() As String
    Password = smPass
End Property

Public Property Let Port(ByVal ilValue As Integer)
    imPort = ilValue
End Property

Public Property Get Port() As Integer
    Port = imPort
End Property

Public Property Let TLS(ByVal blValue As Boolean)
    bmTls = blValue
End Property

Public Property Get TLS() As Boolean
    TLS = bmTls
End Property

Public Property Let User(ByVal slvalue As String)
    smUser = slvalue
End Property

Public Property Get User() As String
    User = smUser
End Property
Public Sub GetHostInfo()
Dim rstSite As ADODB.Recordset

On Error GoTo ErrHand
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rstSite = gSQLSelectCall(SQLQuery)
    If Not rstSite.EOF Then
        'siteEmailFromName misnamed...now holds TLS info!
        If Trim$(rstSite!siteEmailFromName) = "1" Then
            bmTls = True
        Else
            bmTls = False
        End If
        smHost = Trim$(rstSite!siteEmailHost)
        imPort = rstSite!siteEmailPort
        smUser = Trim(rstSite!siteEmailAcctName)
        smPass = Trim(rstSite!siteEmailPassword)
    End If
    Set rstSite = Nothing
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
        'ttp 5217
    gHandleError "AffErrorLog.txt", "CEmailInfo-GetHostInfo"
End Sub

