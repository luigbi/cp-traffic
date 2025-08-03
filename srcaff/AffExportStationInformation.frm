VERSION 5.00
Begin VB.Form frmExportStationInformation 
   Caption         =   "Export Station Information"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOption 
      Caption         =   "OLA"
      Height          =   390
      Index           =   3
      Left            =   1410
      TabIndex        =   5
      Top             =   1695
      Width           =   1725
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Wegener"
      Height          =   390
      Index           =   2
      Left            =   1410
      TabIndex        =   4
      Top             =   1215
      Width           =   1725
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "X-Digital"
      Height          =   390
      Index           =   1
      Left            =   1410
      TabIndex        =   3
      Top             =   720
      Width           =   1725
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Agreements"
      Height          =   390
      Index           =   0
      Left            =   1410
      TabIndex        =   2
      Top             =   240
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2265
      TabIndex        =   1
      Top             =   2340
      Width           =   1665
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   420
      TabIndex        =   0
      Top             =   2340
      Width           =   1665
   End
   Begin VB.Label lblResult 
      ForeColor       =   &H00000040&
      Height          =   480
      Left            =   300
      TabIndex        =   8
      Top             =   2865
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Export as"
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label lblStation 
      Caption         =   " 'StationInformation.csv'"
      Height          =   510
      Left            =   150
      TabIndex        =   6
      Top             =   3720
      Width           =   4530
   End
End
Attribute VB_Name = "frmExportStationInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkOption_Click(Index As Integer)
Dim c As Integer
Dim blSomethingChecked As Boolean
    lblResult.Caption = ""
    For c = 0 To chkOption.Count - 1
        If chkOption(c).Value = 1 Then
            blSomethingChecked = True
            Exit For
        End If
    Next c
    If blSomethingChecked Then
        cmdExport.Enabled = True
    Else
        cmdExport.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
Dim olRs As ADODB.Recordset
Dim slErrorMessage As String
Dim slArrayOfOptions() As String
Dim slListOfOptions As String
Dim slStartDate As String
Dim blLogSuccess As Boolean
Dim slLogSuccess As String
    slStartDate = mPrepDisplay
    slListOfOptions = mCheckOptions(slArrayOfOptions) ' array modified in function
    slErrorMessage = mQueryDatabase(slArrayOfOptions, olRs)
    If slErrorMessage = "No errors" Then
        slErrorMessage = mWriteCsv(olRs)
    End If
    blLogSuccess = mSendLog(slErrorMessage, slListOfOptions, slStartDate)
    If Not blLogSuccess Then
        slLogSuccess = " Failed to create 'StationInformation.txt' log"
    End If
    If slErrorMessage = "No errors" Then
        lblResult.Caption = "StationInformation created. " & slLogSuccess
    Else
        lblResult.Caption = "Errors writing 'StationInformation.csv' " & slErrorMessage & slLogSuccess
    End If
finish:
Screen.MousePointer = vbDefault
Set olRs = Nothing
Erase slArrayOfOptions
cmdCancel.Caption = "Done"
Exit Sub
ERRORBOX:
MsgBox "Errors writing 'StationInformation.csv' ", vbOKOnly + vbInformation
GoTo finish
End Sub
Private Function mPrepDisplay() As String
    Screen.MousePointer = vbHourglass
    lblResult.Caption = ""
    DoEvents
    mPrepDisplay = Format$(gNow(), "m/d/yyyy hh:mm")

End Function
Private Function mWriteCsv(ByRef olRs As Recordset) As String
Dim slErrorMessage As String
Dim olFileSys As FileSystemObject
Dim olCsv As TextStream
Dim slPath As String
Dim slRowToWrite As String
Dim slComma As String
Dim slAppendLine As String
Dim olField As Field
Dim slFormattedString As String
slComma = ","
Set olFileSys = New FileSystemObject
slPath = sgExportDirectory & "StationInformation.csv"
On Error GoTo ERRORBOX
If olFileSys.FolderExists(sgExportDirectory) Then
    Set olCsv = olFileSys.OpenTextFile(slPath, ForWriting, True)
    olCsv.WriteLine "Counterpoint Station Information Export"
    olCsv.WriteLine "Call Letters,City of License,Market,Market Code,Market Rank, Format,Format Code,State Name,ST,State Code,Zone,Zone Code,Owner,Agreements,X-Digital,Wegener,OLA"
    If olRs.EOF And olRs.BOF Then
        mWriteCsv = "There are no records to write to 'StationInformation.csv'"
        GoTo finish
    End If
    olRs.MoveFirst
    Do While Not olRs.EOF
        slRowToWrite = ""
        For Each olField In olRs.Fields
            slFormattedString = mWriteField(olField)
            'not sure if need to test for first error string.
            If slFormattedString = "Error reading records in mWriteCsv" Or slFormattedString = "Error in function mWriteField" Then
                mWriteCsv = slFormattedString
                GoTo finish
            End If
                slRowToWrite = slRowToWrite & slFormattedString & slComma
        Next olField
        olCsv.WriteLine slRowToWrite
        olRs.MoveNext
    Loop
    olRs.Close
    olCsv.Close
    slErrorMessage = "No errors"
Else
    slErrorMessage = "Couldn't find Export Folder."
End If
mWriteCsv = slErrorMessage
finish:
Set olField = Nothing
Set olFileSys = Nothing
Set olCsv = Nothing
Exit Function
ERRORBOX:
mWriteCsv = "Error reading records in mWriteCsv"
GoTo finish
End Function
Private Function mWriteField(olField As Field) As String
Dim slReturnName As String
Dim ilWriteToLineCode As Integer
On Error GoTo ERRORBOX
If IsNull(olField.Value) Then
    mWriteField = ""
    Select Case UCase(olField.Name)
        Case "SHTTCITYLIC", "MKTNAME", "FMTNAME", "ARTTLASTNAME"
            mWriteField = gAddQuotes(mWriteField)
    End Select
    Exit Function
Else
    Select Case UCase(olField.Name)
        Case "SHTTCITYLIC", "MKTNAME", "FMTNAME", "ARTTLASTNAME"
            slReturnName = Trim(olField.Value)
            mWriteField = gAddQuotes(slReturnName)
        Case "SHTTUSEDFORATT", "SHTTUSEDFORXDIGITAL", "SHTTUSEDFORWEGENER", "SHTTUSEDFOROLA"
            If Trim(olField.Value) = "Y" Then
                mWriteField = "x"
            Else
                mWriteField = ""
            End If
        Case Else
            mWriteField = Trim(olField.Value)
    End Select
End If
Exit Function
ERRORBOX:
mWriteField = "Error in function mWriteField"
End Function
Private Function mQueryDatabase(slArrayOfOptions() As String, ByRef olRs As Recordset) As String
Dim slErrorMessage As String
Dim SQLQuery As String
    SQLQuery = mSelectClause
    SQLQuery = SQLQuery & mFromClause
    SQLQuery = SQLQuery & "WHERE shttType = 0"
    SQLQuery = SQLQuery & mAppendOptionsToSql(slArrayOfOptions)
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    On Error GoTo ERRORBOX
    Set olRs = gSQLSelectCall(SQLQuery)
    slErrorMessage = "No errors"
    mQueryDatabase = slErrorMessage
    Exit Function
ERRORBOX:
    slErrorMessage = "Problem with query in mQueryDatabase. "
    mQueryDatabase = slErrorMessage
End Function
Private Function mSelectClause() As String
Dim slSelect As String
slSelect = "select shttCallLetters, shttCityLic, mktname,mktGroupName,mktRank,fmtName,fmtGroupName, " _
& "sntName, sntPostalName,sntGroupName,tztName,tztGroupName,arttLastName, shttUsedForAtt, shttUsedForXDigital, shttUsedforWegener, shttUsedForOLA"
slSelect = slSelect & " "
mSelectClause = slSelect
End Function
Private Function mFromClause() As String
Dim slFrom As String
slFrom = "from shtt left outer join mkt on shttmktcode = mktcode left outer join fmt_station_format on shttfmtcode = fmtcode left outer join SNT on shttState = sntPostalName left outer join tzt on shtttztcode = tztcode left outer join artt on shttownerarttcode = arttcode"
slFrom = slFrom & " "
mFromClause = slFrom
End Function

Private Function mAppendOptionsToSql(slArrayOfOptions() As String) As String
Dim slAppendLine As String
Dim c As Integer
    slAppendLine = " and ("
    For c = 0 To UBound(slArrayOfOptions)
        Select Case UCase(slArrayOfOptions(c))
            Case "AGREEMENTS"
                slAppendLine = slAppendLine & "shttUsedForAtt <> 'N' "
            Case "X-DIGITAL"
                slAppendLine = slAppendLine & "shttUsedForXDigital = 'Y' "
            Case "WEGENER"
                slAppendLine = slAppendLine & "shttUsedForWegener = 'Y' "
            Case "OLA"
                slAppendLine = slAppendLine & "shttUsedForOla = 'Y' "
        End Select
        slAppendLine = slAppendLine & "OR "
    Next c
    slAppendLine = Left$(slAppendLine, Len(slAppendLine) - 3) & ")"
    mAppendOptionsToSql = slAppendLine
End Function
Private Function mCheckOptions(ByRef slArrayOfOptions() As String) As String
Dim c As Integer
Dim slListOfOptions As String
    For c = 0 To chkOption.Count - 1
        If chkOption(c).Value = 1 Then
            slListOfOptions = slListOfOptions & chkOption.Item(c).Caption & ","
        End If
    Next c
    slListOfOptions = Left(slListOfOptions, Len(slListOfOptions) - 1)
    slArrayOfOptions = Split(slListOfOptions, ",")
    mCheckOptions = slListOfOptions
End Function
Private Function mSendLog(slErrorMessage, slListOfOptions As String, slStartDate As String) As Boolean
Dim olFileSys As FileSystemObject
Dim olLog As TextStream
Dim slToFile As String
Dim slNowTime As String

Set olFileSys = New FileSystemObject
On Error GoTo ERRORBOX:
    slToFile = sgMsgDirectory & "StationInformation.Txt"
    slNowTime = Format$(gNow(), "hh:mm")
    If olFileSys.FolderExists(sgExportDirectory) Then
        Set olLog = olFileSys.OpenTextFile(slToFile, ForAppending, True)
        olLog.WriteLine "Exported Station Information. Options Chosen: " & slListOfOptions & " Started: " & slStartDate & " Errors: " & slErrorMessage _
        & " Ended: " & slNowTime
        olLog.Close
        mSendLog = True
    Else
        mSendLog = False
    End If
finish:
Set olLog = Nothing
Set olFileSys = Nothing
Exit Function
ERRORBOX:
GoTo finish
End Function

Private Sub Form_Initialize()
    gCenterForm frmExportStationInformation
End Sub

Private Sub Form_Load()
On Error Resume Next
lblStation = sgExportDirectory & "StationInformation.csv"
End Sub

