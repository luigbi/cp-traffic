VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' For debugging the csiWinSock.DLL
'Private Declare Function csiRunEXE Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slEXEName$, ByVal slFileName$) As Integer
'Private Declare Function csiExecSQL Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slFileName$, ByVal slSQL$) As Integer
'Private Declare Function csiExportHeaders Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slFileName$) As Integer
'Private Declare Function csiExportCommit Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slFileName$) As Integer
'Private Declare Function csiExportWebLogActivity Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slFileName$) As Integer
'Private Declare Function csiImportSiteOptions Lib "E:\CSI\MyTests\ISAPI\CSIWebManager\TestJig\csiWinSock\Debug\csiWinSock.dll" (ByVal slFileName$) As Integer

Private Declare Function csiRunEXE Lib "csiWinSock.dll" (ByVal slEXEName$, ByVal slFileName$) As Integer
Private Declare Function csiExecSQL Lib "csiWinSock.dll" (ByVal slFileName$, ByVal slSql$) As Integer
Private Declare Function csiExportHeaders Lib "csiWinSock.dll" (ByVal slFileName$) As Integer
Private Declare Function csiExportCommit Lib "csiWinSock.dll" (ByVal slFileName$) As Integer
Private Declare Function csiExportWebLogActivity Lib "csiWinSock.dll" (ByVal slFileName$) As Integer
Private Declare Function csiImportSiteOptions Lib "csiWinSock.dll" (ByVal slFileName$) As Integer

'*******************************************************************************
'
'*******************************************************************************
Private Sub Class_Initialize()
End Sub

'*******************************************************************************
'
'*******************************************************************************
Private Sub Class_Terminate()
End Sub

'*******************************************************************************
'
'*******************************************************************************
Public Function RunEXE(sEXEName As String, sFileName As String) As String
    
    Dim ilRet As Long

    ilRet = csiRunEXE(Trim(sEXEName), Trim(sFileName))
    If ilRet <> 1 Then
        RunEXE = "ERROR:"
    Else
        RunEXE = "SUCCESS"
    End If
End Function

'*******************************************************************************
'
'*******************************************************************************
Public Function ExecSQL(sSQL As String)
    Dim llErrorNum As Long
    Dim ilRet As Integer
    Dim slFileName As String

    On Error GoTo ErrHandler

    ExecSQL = ""
    slFileName = gGetUniqueFileName()

    ilRet = csiExecSQL(slFileName, sSQL)
    
    ilRet = gFileExist(slFileName)
    If ilRet = 1 Then
        ExecSQL = "No Response"
    Else
        ExecSQL = FileToString(slFileName)
        Kill (slFileName)
    End If
    Exit Function
    
ErrHandler:
    llErrorNum = Err.Number
    ExecSQL = "ERROR: " & llErrorNum
End Function

'*******************************************************************************
'
'*******************************************************************************
Public Function ExportHeaders(sFileName As String) As String

    Dim ilRet As Long

    ilRet = csiExportHeaders(Trim(sFileName))
    If ilRet <> 1 Then
        ExportHeaders = "ERROR:"
    Else
        ExportHeaders = "SUCCESS"
    End If
End Function

'*******************************************************************************
'
'*******************************************************************************
Public Function ExportCommit() As String
    
    Dim ilRet As Long

    ilRet = csiExportCommit("Nothing.txt")
    If ilRet <> 1 Then
        ExportCommit = "ERROR:"
    Else
        ExportCommit = "SUCCESS"
    End If
End Function

'*******************************************************************************
'
'*******************************************************************************
Public Function ExportWebL(sFileName As String) As String
    
    Dim ilRet As Long

    ilRet = csiExportWebLogActivity(Trim(sFileName))
    If ilRet <> 1 Then
        ExportWebL = "ERROR:"
    Else
        ExportWebL = "SUCCESS"
    End If
End Function

'*******************************************************************************
'
'*******************************************************************************
Public Function ImportSiteOptions(sFileName As String) As String
    
    Dim ilRet As Long

    ilRet = csiImportSiteOptions(Trim(sFileName))
    If ilRet <> 1 Then
        ImportSiteOptions = "ERROR:"
    Else
        ImportSiteOptions = "SUCCESS"
    End If
End Function

'*******************************************************************************
'
'*******************************************************************************
Private Function FileToString(slFileName As String) As String
    
    Dim ilFile As Integer
    Dim llErrorNum As Long

    On Error GoTo ErrHandler
    
    ilFile = FreeFile
    Open slFileName For Input As #ilFile
    FileToString = StrConv(InputB(LOF(ilFile), ilFile), vbUnicode)
    Close #ilFile
    Exit Function
ErrHandler:
    llErrorNum = Err.Number
    FileToString = "ERROR: " & llErrorNum
End Function



