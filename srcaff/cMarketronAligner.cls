VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMarketronAligner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'myImportRs:  all the 'not found' import spots
'myExportRs:  all the export spots.  the replacementKey is set to -1 if it's been found.  So that leaves the spots with replacementKey of 0 available for myImportRs spots to try to match.
'When myImportRs does find a match, it changes its replacementKey to the value of myExportRs' key.
'When done all myImportRs that can be done, any spot with a replacementKey of 0 failed to find a match; otherwise, write to database with the replacementKey as the astcode!
'the final step is to set the 'found' to true for each import spot that can be found to have a replacementKey.
Private bmAllowNonCompliance As Boolean
Private myExportRs As ADODB.Recordset
Private myImportRs As ADODB.Recordset
Private lmAlignCount As Long
Private smMondayDate As String
Private Const FORMATTIME As String = "hh:nn:ss"
Private Const MATCHTRUE As Integer = -1
Private Const MATCHFALSE As Integer = 0
Private Const BASEDATE As String = "1/2/1971"
'Private Type SpotComparer
'    dDate As Date
'    dTime As Date
'    dEndTime As Date
'    lKey As Long
'    sISCI As String
'    sReason As String
'    lReplacementKey As Long
'End Type
Public Property Let AllowNonCompliance(blValue As Boolean)
    bmAllowNonCompliance = blValue
End Property
Public Property Get AllowNonCompliance() As Boolean
    AllowNonCompliance = bmAllowNonCompliance
End Property
Public Property Get AlignCount() As Long
    AlignCount = lmAlignCount
End Property

Private Sub Class_Initialize()
    Clear
End Sub
Private Sub Class_Terminate()
    Set myExportRs = Nothing
    Set myImportRs = Nothing
End Sub

Public Function AddImport(slDate As String, slStartTime As String, llKey As Long, slISCI As String, ilAdv As Integer, ilSpotLen As Integer) As Boolean
    AddImport = mAdd(True, slDate, BASEDATE, slStartTime, "00:00:00", llKey, slISCI, ilAdv, ilSpotLen)
End Function
Public Function AddExport(slStartDate As String, slEndDate As String, slStartTime As String, slEndTime As String, llKey As Long, slISCI As String, ilAdf As Integer, ilSpotLen As Integer) As Boolean
    AddExport = mAdd(False, slStartDate, slEndDate, slStartTime, slEndTime, llKey, slISCI, ilAdf, ilSpotLen)
End Function
Public Sub MarkExportAsFound(llAstCode As Long)
    myExportRs.Filter = "Key = " & llAstCode
    If Not myExportRs.EOF Then
        myExportRs.Fields("ReplacementKey") = MATCHTRUE
    End If
End Sub
Public Function Align() As Boolean
    Dim blRet As Boolean
    Dim slExportFilter As String
    Dim slBaseFilter
    Dim slAndBaseFilter As String
    Dim ilImportAdvCode As Integer
    
    blRet = True
    lmAlignCount = 0
    slBaseFilter = "ReplacementKey = " & MATCHFALSE
On Error GoTo errbox
    myExportRs.Filter = slBaseFilter
    If Not myExportRs.EOF Then
        slAndBaseFilter = " AND " & slBaseFilter
        myImportRs.Filter = adFilterNone
        myImportRs.MoveFirst
        Do While Not myImportRs.EOF
            'test 1: compliant?   same isci, on date and in time range
            '9102 test 1: compliant? match adv,isci,spot length,date and time!
            slExportFilter = "ISCI = '" & myImportRs("Isci") & "' And AdvCode = " & myImportRs("AdvCode") & " AND SpotLength = " & myImportRs("SpotLength") & " AND StartDate <= #" & myImportRs("StartDate") & "# AND EndDate >= #" & myImportRs("StartDate") & "# And StartTime <= #" & myImportRs("StartTime") & "# AND EndTime >= #" & myImportRs("StartTime") & "#"
            myExportRs.Filter = slExportFilter & slAndBaseFilter
            If Not myExportRs.EOF Then
                myImportRs("ReplacementKey") = myExportRs("Key")
                myExportRs("ReplacementKey") = MATCHTRUE
                lmAlignCount = lmAlignCount + 1
            End If
            myImportRs.MoveNext
        Loop
        myImportRs.Filter = slBaseFilter
        Do While Not myImportRs.EOF
            'test 2: mostly compliant?   same isci, on date but not in time range
            '9102 test 2: mostly compliant.  Forget isci
            slExportFilter = "AdvCode = " & myImportRs("AdvCode") & " AND SpotLength = " & myImportRs("SpotLength") & " AND StartDate <= #" & myImportRs("StartDate") & "# AND EndDate >= #" & myImportRs("StartDate") & "# And StartTime <= #" & myImportRs("StartTime") & "# AND EndTime >= #" & myImportRs("StartTime") & "#"
            myExportRs.Filter = slExportFilter & slAndBaseFilter
            If Not myExportRs.EOF Then
                myImportRs("ReplacementKey") = myExportRs("Key")
                myExportRs("ReplacementKey") = MATCHTRUE
                lmAlignCount = lmAlignCount + 1
            End If
            myImportRs.MoveNext
        Loop
        If bmAllowNonCompliance Then
            myImportRs.Filter = slBaseFilter
            Do While Not myImportRs.EOF
                'test 3: non-compliant--same isci?  forget time and date.
                '9102 test 3: non compliant -- forget date and time.  Just Adv and spot length
                'slExportFilter = "ISCI = '" & myImportRs("Isci") & "'"
                slExportFilter = "AdvCode = " & myImportRs("AdvCode") & " AND SpotLength = " & myImportRs("SpotLength")
                myExportRs.Filter = slExportFilter & slAndBaseFilter
                If Not myExportRs.EOF Then
                    myImportRs("ReplacementKey") = myExportRs("Key")
                    myExportRs("ReplacementKey") = MATCHTRUE
                    lmAlignCount = lmAlignCount + 1
                End If
                myImportRs.MoveNext
            Loop
'            myImportRs.Filter = slBaseFilter
'            Do While Not myImportRs.EOF
'                'test 4: non-compliant--date and time...forget isci, but be concerned with advertiser!
'                ilImportAdvCode = mReturnAdvertiser(myImportRs("Isci"))
'                If ilImportAdvCode > 0 Then
'                    slExportFilter = "AdvCode =" & ilImportAdvCode & " AND StartDate <= #" & myImportRs("StartDate") & "# AND EndDate >= #" & myImportRs("StartDate") & "# And StartTime <= #" & myImportRs("StartTime") & "# AND EndTime >= #" & myImportRs("StartTime") & "#"
'                    myExportRs.Filter = slExportFilter & slAndBaseFilter
'                    If Not myExportRs.EOF Then
'                        myImportRs("ReplacementKey") = myExportRs("Key")
'                        myExportRs("ReplacementKey") = MATCHTRUE
'                        lmAlignCount = lmAlignCount + 1
'                    End If
'                End If
'                myImportRs.MoveNext
'            Loop
        End If
    End If
    Align = blRet
    Exit Function
errbox:
    Align = False
End Function
Public Function ReturnReplacementAst(llAstCode As Long) As Long
    Dim llRet As Long
    
    llRet = 0
    myImportRs.Filter = "Key = " & llAstCode
    If Not myImportRs.EOF Then
        If myImportRs.Fields("ReplacementKey") <> 0 Then
            llRet = myImportRs.Fields("ReplacementKey")
        End If
    End If
    ReturnReplacementAst = llRet
End Function
Public Sub Clear()
    bmAllowNonCompliance = True
    lmAlignCount = 0
    If Not myExportRs Is Nothing Then
        If (myExportRs.State And adStateOpen) <> 0 Then
            myExportRs.Close
        End If
    End If
    Set myExportRs = mPrepRecordset()
    If Not myImportRs Is Nothing Then
        If (myImportRs.State And adStateOpen) <> 0 Then
            myImportRs.Close
        End If
    End If
    Set myImportRs = mPrepRecordset()
End Sub
Private Function mAdd(blImport As Boolean, slStartDate As String, slEndDate As String, slStartTime As String, slEndTime As String, llKey As Long, slISCI As String, ilAdf As Integer, ilLen As Integer) As Boolean
    Dim dlStartTime As Date
    Dim dlEndTime As Date
    Dim dlStartDate As Date
    Dim dlEndDate As Date
    
On Error GoTo errbox
    dlStartTime = Format(slStartTime, FORMATTIME)
    dlEndTime = Format(slEndTime, FORMATTIME)
    dlStartDate = CDate(slStartDate)
    dlEndDate = CDate(slEndDate)
    If blImport Then
        myImportRs.AddNew Array("StartDate", "EndDate", "StartTime", "EndTime", "Key", "ISCI", "ReplacementKey", "AdvCode", "SpotLength"), Array(dlStartDate, dlEndDate, dlStartTime, dlEndTime, llKey, slISCI, MATCHFALSE, ilAdf, ilLen)
    Else
        myExportRs.AddNew Array("StartDate", "EndDate", "StartTime", "EndTime", "Key", "ISCI", "ReplacementKey", "AdvCode", "SpotLength"), Array(dlStartDate, dlEndDate, dlStartTime, dlEndTime, llKey, slISCI, MATCHFALSE, ilAdf, ilLen)
    End If
    mAdd = True
    Exit Function
errbox:
    mAdd = False
End Function
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "StartDate", adDate
            .Append "EndDate", adDate
            .Append "StartTime", adDBTime
            .Append "EndTime", adDBTime
            .Append "Key", adInteger
            .Append "ISCI", adChar, 20
            .Append "ReplacementKey", adInteger
            .Append "AdvCode", adInteger
            .Append "SpotLength", adInteger
        End With
    myRs.Open
    myRs("Key").Properties("optimize") = True
    myRs.Sort = "Key"
    Set mPrepRecordset = myRs
End Function
'Private Function mReturnAdvertiser(slISCI As String) As Integer
'    'select cifadfcode from "CIF_Copy_Inventory" inner join "CPF_Copy_Prodct_ISCI" on cifcpfcode = cpfcode where cpfISCI = 'ACD-1001-30'
'    Dim ilRet As Integer
'
'    ilRet = 0
'On Error GoTo ERRORBOX
'    SQLQuery = "select cifadfcode from CIF_Copy_Inventory inner join CPF_Copy_Prodct_ISCI on cifcpfcode = cpfcode where cpfISCI = '" & Trim$(slISCI) & "'"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        ilRet = rst!cifAdfCode
'    End If
'Cleanup:
'    mReturnAdvertiser = ilRet
'    If Not rst Is Nothing Then
'        If (rst.State And adStateOpen) <> 0 Then
'            rst.Close
'        End If
'    End If
'    Exit Function
'ERRORBOX:
'    ilRet = 0
'    GoTo Cleanup
'End Function


