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
Private bmSearchDsn As Boolean
Private Const DSNNAME = "Name"
Private smRootDrive As String
Enum SystemType
    Affiliate = 0
    Traffic = 1
End Enum
Private Sub Class_Initialize()
    bmSearchDsn = True
    smSystem = Traffic
    smGlobalReportPath = sgRptPath
    '2-12-15 Use  path defined in traffic.ini from SaveReports or LogGenExport
    smGlobalReportSavePath = sgSaveFilePathToUse        '2-12-15 sgRptSavePath
    smGlobalDatabasePath = sgDBPath
    smGlobalDSNName = "" 'sgDatabaseName
    smGlobalLogoPath = sgLogoPath       '8-14-14 & "rptlogo.bmp"
    smRootDrive = sgRootDrive
   ' affiliate
'    smSystem = Affiliate
'    smGlobalReportPath = sgReportDirectory
'    smGlobalReportSavePath = sgExportDirectory
'    smGlobalDSNName = sgDatabaseName
'    smGlobalDatabasePath = sgDBPath 'so affiliate can make Native calls
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
    'Allow traffic to get dsnName from affiliat.ini, or traffic.ini (Name=).  Currently  not used(10/26/11), but I put in just in case we needed to make a
    'a report be run as odbc...like in crystal 2008.
    ' the better method would be to create a sgGlobalDSNName and have it grab the dsn name from ini file.  We would then reference that variable here and
    'could remove what I am doing below.  Until then, each time a new report class is created and this call is made, we will go into the ini file.
    'the bmSearchDsn is so that if the report class tried once to get the dsn name and failed, it won't try again.
    If bmSearchDsn = True Then
        If Len(smGlobalDSNName) = 0 Then
            mGetDsnName
        End If
    End If
    GlobalDSNName = smGlobalDSNName
End Property
Private Sub mGetDsnName()
    'affiliate.ini? then grab dsnname from there. Otherwise, search in traffic.ini
    Dim slLocation As String
    Dim ilRet As Integer
    Dim slReturn As String * 130
    Dim slFileName As String

    If igDirectCall = -1 Then
        slFileName = sgIniPath
    Else
        slFileName = CurDir$ & "\"
    End If
    '8886
    'If Dir(slFileName & "affiliat.ini") > " " Then
    If gFileExist(slFileName & "affiliat.ini") = FILEEXISTS Then
        slFileName = slFileName & "affiliat.ini"
        'slLocation = "Database"
        If igTestSystem Then
            slLocation = "TestDatabase"
        Else
            slLocation = "Database"
        End If
    Else
        slFileName = slFileName & "Traffic.ini"
        If igTestSystem Then
            slLocation = "TestLocations"
        Else
            slLocation = "Locations"
        End If
    End If

    On Error GoTo gObtainIniValuesErr
    ilRet = GetPrivateProfileString(slLocation, DSNNAME, "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        bmSearchDsn = False
        smGlobalDSNName = ""
    Else
        smGlobalDSNName = Left$(slReturn, ilRet)
    End If
    Exit Sub
gObtainIniValuesErr:
    bmSearchDsn = False
    smGlobalDSNName = ""
End Sub

