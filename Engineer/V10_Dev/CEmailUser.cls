VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmailUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Traffic version
'depends on sgDBPath, SetEmailThisUser uses sgSUNName and btrieve
Dim lmCEFIndex As Long
Dim smError As String
Dim smUser As String
Dim smClient As String
Dim smUserEmail As String
Public Property Get UserEmail() As String
    UserEmail = smUserEmail
End Property
Public Property Get UserName() As String
    UserName = smUser
End Property
Public Property Get Client() As String
    Client = smClient
End Property
Private Sub Class_Initialize()
     
    smClient = gFileNameFilter(Trim$(sgClientName))
    smUser = Trim$(sgUserName) '& " - " & smClient
    smUserEmail = GetEmail()

End Sub

Public Property Get ErrorMessage() As String
    ErrorMessage = smError
End Property

Public Property Let Index(ByVal llCEF As Long)
    lmCEFIndex = llCEF
End Property

Public Property Get Index() As Long
    Index = lmCEFIndex
End Property
Public Sub SetEmailThisUser(slAddress As String)
    Dim slSql As String
    'dan todo
    Me.SetEmail slAddress
    If smError = "" Then
        If lmCEFIndex > 0 Then
            slSql = "UPDATE ust SET ustEMailCefCode = " & lmCEFIndex & " WHERE ustname = '" & smUser & "'"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then  'fail to update, remove from cef. no message box
                gMsg = "A general error has occurred in mInsertNewFromAddress(modGenSubs): "
                gLogMsg "Error: " & gMsg & Err.Description & "; Error # " & Err.Number & "; Line #" & Erl, "engrerrors.txt", False
                slSql = "delete FROM cef_comments_events WHERE cefCode = " & lmCEFIndex
                gSQLWaitNoMsgBox slSql, False
                smError = "Could not set email: " & Err.Description
            End If  'failed to update Ust
        End If  'inserted to CEF
    End If
End Sub
Public Sub SetEmail(slAddress As String)

'Dim slQuery As String
'    lmCEFIndex = 0
'    smError = ""
'On Error GoTo ERRBOX
'    slQuery = "INSERT into cef_comments_events (cefCode, cefComment) values( Replace, '" & slAddress & "')"
'    lmCEFIndex = gInsertAndReturnCode(slQuery, "cef_comments_events", "cefCode", "Replace")
  'no email to set besides users
  smError = "Failed to set email."
    Exit Sub
ERRBOX:
    smError = "Failed to set email: " & Err.Description

End Sub
Public Function GetEmail() As String
    Dim rstUie As ADODB.Recordset
    Dim slSql As String
    Dim slName As String
    
    slName = UCase$(smUser)
    smError = ""
    slSql = "SELECT uieEmail FROM uie_User_Info where UPPER(uieSignOnName) = '" & slName & "' and uieCurrent = 'Y'"
    Set rstUie = cnn.Execute(slSql)
    If Not rstUie.EOF Then
        GetEmail = Trim(rstUie!uieEmail)
        If LenB(GetEmail) = 0 Then
            smError = "Couldn't retrieve email address"
        End If
    Else
        smError = "Couldn't retrieve email address"
    End If
    Set rstUie = Nothing
End Function

