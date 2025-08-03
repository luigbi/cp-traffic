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
Dim hmCef As Integer
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
    Dim ilRet As Integer
    
    hmCef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        smError = "could not open cef table: " & ilRet
    End If
    If Len(Trim(tgUrf(0).sRept)) > 0 Then
        smUser = Trim(tgUrf(0).sRept)
    Else
        smUser = Trim(tgUrf(0).sName)
    End If
    smClient = gFileNameFilter(Trim$(tgSpf.sGClient))
    lmCEFIndex = tgUrf(0).lEMailCefCode
    smUserEmail = GetEmail()
End Sub

Private Sub Class_Terminate()
    btrClose hmCef
    btrDestroy hmCef
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
    Dim hlUrf As Integer
    Dim tlUrf As URF
    Dim tlUrfSearchKey As INTKEY0    'URF key record image
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    'block guide
    
    If (Trim$(tgUrf(0).sName) = sgSUName) Then
       ' smError = "Can't write email address for guide."
        Exit Sub
    End If
    Me.SetEmail (slAddress)
    If smError = "" Then
        hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hlUrf, "", sgDBPath & "urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            ilRecLen = Len(tlUrf)
            tlUrfSearchKey.iCode = tgUrf(0).iCode
            ilRet = btrGetEqual(hlUrf, tlUrf, ilRecLen, tlUrfSearchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gUrfDecrypt tlUrf
                tlUrf.lEMailCefCode = lmCEFIndex
                gUrfEncrypt tlUrf
                ilRet = btrUpdate(hlUrf, tlUrf, ilRecLen)
                If ilRet = BTRV_ERR_NONE Then
                   tgUrf(0).lEMailCefCode = lmCEFIndex
                End If  'updated urf
            Else
                smError = "Couldn't update user with email address:" & ilRet
            End If ' found urf
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
        Else
            smError = "Couldn't open Urf to save email address:" & ilRet
        End If ' opened urf
    End If

End Sub
Public Sub SetEmail(slAddress As String)
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlCef As CEF

    lmCEFIndex = 0
    smError = ""
    tlCef.lCode = 0
    tlCef.sComment = slAddress
    ilRecLen = Len(tlCef)
    ilRet = btrInsert(hmCef, tlCef, ilRecLen, 0)
    If ilRet = BTRV_ERR_NONE Then
        lmCEFIndex = tlCef.lCode
    Else
        smError = "Could not set email:" & ilRet
    End If
End Sub
Public Function GetEmail() As String
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlCef As CEF
    
    smError = ""
    If lmCEFIndex > 0 Then
        tlCef.sComment = ""
        ilRecLen = Len(tlCef)    '1009
        ilRet = btrGetEqual(hmCef, tlCef, ilRecLen, lmCEFIndex, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            GetEmail = mStripChr0(tlCef.sComment)
        Else
            smError = "couldn't retrieve email address: " & ilRet
        End If
    Else
        smError = "Index for user not set."
    End If

End Function
Private Function mStripChr0(slInStr As String) As String
'copy of gStripChr0--9/08/11
    Dim slStr As String
    Dim slChr As String
    Dim ilIndex As Integer

    slStr = ""
    If Len(slInStr) > 0 Then
        ilIndex = 1
        slChr = Mid(slInStr, ilIndex, 1)
        Do While slChr <> Chr$(0)
            slStr = slStr & slChr
            ilIndex = ilIndex + 1
            If ilIndex > Len(slInStr) Then
                Exit Do
            End If
            slChr = Mid(slInStr, ilIndex, 1)
        Loop
    End If
    mStripChr0 = Trim$(slStr)
End Function
