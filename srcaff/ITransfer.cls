VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private myFiles As Recordset
'Public LogFile As CLogger
'Public TelNetControl As TelnetTTYClient
Public Property Get StartPaused() As Boolean
End Property
Public Property Get StatusMessage() As String
End Property
Public Property Get NextFile() As String
End Property
Public Property Get Name() As String
End Property
Public Property Let Name(slValue As String)
End Property
Public Property Get DateStored() As String
End Property
Public Property Let DateStored(slValue As String)
End Property
Public Property Get ErrorMessage() As String
End Property
Public Property Get DaysToSave() As Integer
End Property
Public Property Get Address() As String
End Property
'Public Property Get ExportFromPath() As String
'End Property
Public Property Get SaveFolder() As String
End Property
'Public Property Get Mode() As TransferMode
'End Property
Public Property Get IsLoaded() As Boolean
End Property
Public Property Get FilesWaiting(blIsNew As Boolean) As Integer
End Property
Public Property Get FilesProcessed(blIsNew As Boolean) As Integer
End Property
Public Property Get IsNotPasswordProtected() As Boolean
End Property
Public Property Let IsNotPasswordProtected(blValue As Boolean)
End Property
Public Property Get ExtensionToFind() As String
End Property
Public Property Let ExtensionToFind(slExt As String)
End Property
Public Function LoadIni(slIniPath As String) As Boolean
End Function
Public Function Move(blMarkAsUnread As Boolean) As Boolean
End Function
Public Function Process() As Boolean
End Function
Public Sub FillGrid(myFileGrid As MSHFlexGrid)
End Sub
Public Function Connect() As Boolean
End Function
Public Function DisConnect() As Boolean
End Function
