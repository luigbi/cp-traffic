VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRemapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NODATE As String = "1970-01-01"
Private lmRemapAmount As Long
Private imMargin As Integer
Private dmRemappedFirstExportDate As Date
Private dmExportingDate As Date
Private bmIsExportDateAllowed As Boolean
'Private bmIsImport As Boolean
Private bmIsRemapping As Boolean
Public isImport As Boolean
'9851 properties
Private bmIsMonitorTestRun As Boolean
Private bmIsTestMessage As Boolean
Private lmMonitorThreshold As Long
Const REMAPMONITORINI As String = "RemapMonitorTester.ini"
Const REMAPMONITORINIHEADER As String = "RemapMonitorTest"
Const REMAPMONITORINIBlOCK As String = "BlockMonitoring"
Const REMAPMONITORINIASTMAX As String = "AstMaxForTest"
Const REMAPMONITORINIBLOCKEMAIL As String = "TestMessageInsteadOfEmail"
Const REMAPMONITORTHRESHOLDAST As Long = 2000000000
Const REMAPMONITORMONTHS As Integer = 6
Public Property Get ExportingDate() As Date
    ExportingDate = dmExportingDate
End Property
Public Property Let ExportingDate(ByVal vNewValue As Date)
    If bmIsRemapping Then
        If dmExportingDate <> vNewValue Then
            dmExportingDate = vNewValue
            mIsRemapForExportDate
        End If
    End If
End Property
'Public Property Get AstMax() As Long
'    AstMax = lmCurrentAstMax
'End Property
Public Sub StartRemapping()
    Dim blContinueWithRemapping As Boolean
    Dim slSql As String
    Dim rst As ADODB.Recordset
    Dim slData() As String
    Dim slValues() As String
    Dim llCount As Long
    Dim c As Integer
    
    blContinueWithRemapping = False
    isImport = True
    lmRemapAmount = 0
    imMargin = 0
    dmRemappedFirstExportDate = CDate(NODATE)
    slSql = "SELECT CASE WHEN OBJECT_ID('dbo.WebVendorRemapping', 'U') IS NOT NULL THEN 1 ELSE 0 END"
    llCount = gExecWebSQLForVendor(slData, slSql, True)
    If llCount > 0 Then
        For c = 1 To llCount - 1
            slValues = Split(slData(c), ",")
            If slValues(0) = 1 Then
                blContinueWithRemapping = True
            End If
        Next c
    End If
    If blContinueWithRemapping Then
        slSql = "SELECT RemapAmount,Margin,RemappedFirstExportDate FROM WebvendorRemapping"
        llCount = gExecWebSQLForVendor(slData, slSql, True)
        If llCount > 0 Then
            For c = 1 To llCount - 1
                slValues = Split(slData(c), ",")
                If UBound(slValues) >= 2 Then
                    lmRemapAmount = gGetDataNoQuotes(slValues(0))
                    imMargin = gGetDataNoQuotes(slValues(1))
                    dmRemappedFirstExportDate = gGetDataNoQuotes(slValues(2))
                End If
            Next c
        End If
        mSetIsRemapFromProperties
    End If
End Sub
Public Function Remap(llAstCode As Long) As Long
    Dim llRet As Long
    llRet = llAstCode
    If bmIsRemapping Then
        If isImport Then
            If llAstCode > lmRemapAmount Then
                llRet = llAstCode - lmRemapAmount
            End If
        Else
            If bmIsExportDateAllowed And llAstCode < lmRemapAmount Then
                llRet = llAstCode + lmRemapAmount
            End If
        End If
    End If
    Remap = llRet
End Function
Public Sub StartThresholdTesting()
    '9851
    mReadThresholdTestIni
End Sub
Public Sub MonthlyMonitoring()
    '9851
    Dim slSql As String
    Dim clRst As ADODB.Recordset
    Dim dlCurrentDate As Date
    Dim slStr As String
    Dim blReplacePrevious As Boolean
    Dim llPreviousAst As Long
    Dim dlPreviousDate As Date
    Dim llCurrentAstMax As Long

    blReplacePrevious = False
    llPreviousAst = 0
    dlPreviousDate = DateValue(NODATE)
    mReadThresholdTestIni
    If bmIsMonitorTestRun Then
        slSql = "select siteAstMaxLastValue as ast,siteAstMaxDate as astdate from site"
        Set clRst = gSQLSelectCall(slSql)
        If Not clRst.EOF Then
            If Not IsNull(clRst!AST) Then
                llPreviousAst = clRst!AST
            End If
            If Not IsNull(clRst!astdate) Then
                slStr = Format$(clRst!astdate, sgShowDateForm)
                dlPreviousDate = DateValue(slStr)
            End If
        End If
        dlCurrentDate = Now()
        If Year(dlPreviousDate) < Year(dlCurrentDate) Then
            blReplacePrevious = True
        ElseIf Month(dlPreviousDate) < Month(dlCurrentDate) Then
            blReplacePrevious = True
        End If
        If blReplacePrevious Then
            slSql = "Select Max(astcode)as amount from ast"
            Set clRst = gSQLSelectCall(slSql)
            '10005
            If Not IsNull(clRst!amount) Then
                llCurrentAstMax = clRst!amount
            Else
                llCurrentAstMax = 0
            End If
            If mNearingThreshold(llCurrentAstMax, llPreviousAst, lmMonitorThreshold) Then
                mAlertAboutThreshold llCurrentAstMax, lmMonitorThreshold
            Else
                gLogMsg "Monthly Remap tested.  Max astcode=" & llCurrentAstMax, "AffErrorLog.Txt", False
            End If
            slStr = Format$(dlCurrentDate, sgSQLDateForm)
            slSql = "UPDATE site set siteAstMaxLastValue = " & llCurrentAstMax & ",siteAstMaxDate = '" & slStr & "'"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
            End If
        End If
    End If
    Set clRst = Nothing
End Sub
Private Function mNearingThreshold(llCurrentAstMaxCode As Long, llPreviousAstCode As Long, llThreshold As Long) As Boolean
    Dim blRet As Boolean
    Dim llAstDifference As Long
    
    blRet = False
    If llPreviousAstCode > 0 Then
        'how much we increased in a month
        llAstDifference = llCurrentAstMaxCode - llPreviousAstCode
        'how much we'll increase over the next 'x' months
        llAstDifference = llAstDifference * REMAPMONITORMONTHS
        If llCurrentAstMaxCode + llAstDifference >= llThreshold Then
            blRet = True
        End If
    End If
    mNearingThreshold = blRet
End Function
Private Sub mAlertAboutThreshold(llCurrentAstMaxCode As Long, llThreshold As Long)
    Dim slMessage As String
    slMessage = "Warning! " & sgClientName & " is estimated to reach the threshold of " & llThreshold & " within 6 months.  Current ast code is " & llCurrentAstMaxCode
    If igTestSystem Or bmIsTestMessage Then
        gMsgBox slMessage, vbCritical, "Ast threshold"
    Else
        gSendServiceEmail "ast threshold at " & sgClientName, slMessage
    End If
    gLogMsg slMessage, "AffErrorLog.Txt", False
End Sub
Private Sub mReadThresholdTestIni()
    Dim slIniPath As String
    Dim slRet As String
    
    bmIsMonitorTestRun = True
    bmIsTestMessage = False
    lmMonitorThreshold = REMAPMONITORTHRESHOLDAST
On Error GoTo ERRORBOX
    slIniPath = sgExeDirectory & REMAPMONITORINI
    gLoadFromIni REMAPMONITORINIHEADER, REMAPMONITORINIASTMAX, slIniPath, slRet
    If slRet <> "Not Found" Then
        If IsNumeric(slRet) Then
            lmMonitorThreshold = CLng(slRet)
        End If
    End If
    gLoadFromIni REMAPMONITORINIHEADER, REMAPMONITORINIBLOCKEMAIL, slIniPath, slRet
    If slRet <> "Not Found" Then
        If UCase(slRet) = "TRUE" Then
            bmIsTestMessage = True
        End If
    End If
    gLoadFromIni REMAPMONITORINIHEADER, REMAPMONITORINIBlOCK, slIniPath, slRet
    If slRet <> "Not Found" Then
        If UCase(slRet) = "TRUE" Then
            bmIsMonitorTestRun = False
        End If
    End If
    Exit Sub
ERRORBOX:
    gLogMsg "cRemapper-StartThresholdTesting had an issue reading test ini file", "AffErrorLog.Txt", False
End Sub
Private Function mSetIsRemapFromProperties() As Boolean
    Dim dmNoDate As Date
    Dim dlRemapped As Date
    
    dmNoDate = CDate(NODATE)
    bmIsRemapping = False
    If lmRemapAmount > 0 And imMargin > 0 And dmRemappedFirstExportDate > dmNoDate Then
        dlRemapped = DateAdd("d", imMargin, dmRemappedFirstExportDate)
        If dlRemapped >= Now() Then
            bmIsRemapping = True
        End If
    End If
    mSetIsRemapFromProperties = bmIsRemapping
End Function
Private Function mIsRemapForExportDate() As Boolean
    bmIsExportDateAllowed = False
    If dmExportingDate < dmRemappedFirstExportDate Then
        bmIsExportDateAllowed = True
    End If
    mIsRemapForExportDate = bmIsExportDateAllowed
End Function
