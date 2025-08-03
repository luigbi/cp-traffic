VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCsiSystemConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private smGlobalReportPath As String
Private smGlobalReportSavePath As String
Private smGlobalDatabasePath As String
Private smGlobalDSNName As String
Private smGlobalLogoPath As String
Private smSystem As SystemType
Private Const DSNNAME = "Name"
Private smRootDrive As String
Enum SystemType
    Affiliate
    Traffic
End Enum
Private Sub Class_Initialize()
'    bmSearchDsn = True
'    smSystem = Traffic
'    smGlobalReportPath = sgRptPath
'    smGlobalReportSavePath = sgRptSavePath
'    smGlobalDatabasePath = sgDBPath
'    'smGlobalDSNName = ""
'    smGlobalDSNName = sgDatabaseName
   ' affiliate
    smSystem = Affiliate
    smGlobalReportPath = sgReportDirectory
    smGlobalReportSavePath = sgExportDirectory
    smGlobalDSNName = sgDatabaseName
    'so affiliate can make Native calls...but in testing couldn't change odbc reports to native!
    smGlobalDatabasePath = sgDBPath
    smGlobalLogoPath = sgLogoPath 	'8-19-14 & "rptlogo.bmp"
    smRootDrive = sgRootDrive
End Sub
Public Property Get RootDrive() As String
    RootDrive = smRootDrive
End Property
Public Property Get System() As SystemType
    System = smSystem
End Property
Public Property Get GlobalLogoPath() As String
    GlobalLogoPath = smGlobalLogoPath
End Property
Public Property Get GlobalReportPath() As String
    GlobalReportPath = smGlobalReportPath
End Property
Public Property Get GlobalReportSavePath() As String
    GlobalReportSavePath = smGlobalReportSavePath
End Property
Public Property Get GlobalDatabasePath() As String
    GlobalDatabasePath = smGlobalDatabasePath
End Property
Public Property Get GlobalDSNName() As String
    GlobalDSNName = smGlobalDSNName
End Property
