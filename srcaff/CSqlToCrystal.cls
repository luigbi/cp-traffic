VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSqlToCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ChangeNumbers As Boolean
Public ConvertFrom As String
Public ConvertTo As String
Public UseGlobalTableList As Boolean
Private tmTableNames() As DDFFILENAMES
Private emFromType As TypeOfFrom
Enum TypeOfFrom
    fromdot
    fromcomma
    fromjoin
End Enum

Public Function Convert(From As String) As String
Dim slConvertPartial As String
Dim blRet As Boolean
    If mPrepForCrystal(From) Then
        If Not UseGlobalTableList Then
            blRet = mGetTables()
        Else
            mArrayFromGlobalArray
        End If
        slConvertPartial = mStripToWhereClause
        ConvertTo = mAddTablesToFields(tmTableNames, slConvertPartial)
        Convert = ConvertTo
    End If
End Function
Private Function mGetTables() As Boolean
    Dim slLongNames() As String
    Dim ilWhereStart As Integer
    Dim slFromStatement As String
    Dim ilNumberInArray As Integer
    Dim ilFromStart As Integer
    ilFromStart = InStr(1, ConvertFrom, "From", vbTextCompare)
    If ilFromStart > 0 Then        'from exists
        ilWhereStart = InStr(1, ConvertFrom, "Where", vbTextCompare)
        If ilWhereStart > 0 Then    'where exists
            slFromStatement = Mid(ConvertFrom, ilFromStart + 4, ilWhereStart - ilFromStart - 5)
        Else
            slFromStatement = ConvertFrom
        End If
        mArrayFromStatement slFromStatement, slLongNames
     End If
     ilNumberInArray = UBound(slLongNames)
    If ilNumberInArray = 0 Then ' is tgddffilenames 1 too large?
        'no from statement or something wrong: grab tgddffilenames as source of table names
        mArrayFromGlobalArray
        Exit Function
    End If
    ReDim tmTableNames(0 To ilNumberInArray)
    mCleanTableNames slLongNames, ilNumberInArray
    mSetTableShortNames
End Function
Private Sub mCleanTableNames(slLongNames() As String, ilNumberInArray As Integer)
Dim c As Integer
Dim ilNameStart As Integer
Dim ilNameEnd As Integer
Dim slTempString As String
Dim ilCountUniqueNames As Integer
    ilCountUniqueNames = 0
    Select Case emFromType
        Case fromdot    ' fragment looks like shttCode = vef_Vehicles
       ' always the last word of array, except for last element: this is unused.
            c = 0
            tmTableNames(ilCountUniqueNames).sLongName = Trim(slLongNames(c))
            ilCountUniqueNames = ilCountUniqueNames + 1
             For c = 1 To ilNumberInArray - 1
                ilNameStart = InStrRev(slLongNames(c), " ", , vbTextCompare)
                If ilNameStart > 0 Then
                    slTempString = Trim(Mid(slLongNames(c), ilNameStart))
                    If Not mGetTablesTestIfExists(slTempString) Then
                        tmTableNames(ilCountUniqueNames).sLongName = slTempString
                        ilCountUniqueNames = ilCountUniqueNames + 1
                   End If
                End If
            Next c
        Case fromcomma  ' fragment looks like VEF_Vehicles
        'usually just the word, except where 'join' was also there, then could be long phrase after
            For c = 0 To ilNumberInArray
                slTempString = Trim(slLongNames(c))
                ilNameEnd = InStr(1, Trim(slLongNames(c)), " ", vbTextCompare)
                If ilNameEnd > 0 Then 'more than one word
                    slTempString = Trim(Mid(slLongNames(c), 1, ilNameEnd))
                End If
                If Not mGetTablesTestIfExists(slTempString) Then
                    tmTableNames(ilCountUniqueNames).sLongName = slTempString
                    ilCountUniqueNames = ilCountUniqueNames + 1
                End If
           Next c
        Case fromjoin  'fragment looks like VEF_Vehicles......
        ' first word. first element has space  before word..or (
            c = 0
            slLongNames(c) = Replace(slLongNames(c), "(", " ")
            tmTableNames(ilCountUniqueNames).sLongName = Trim(slLongNames(c))
            ilCountUniqueNames = ilCountUniqueNames + 1
            For c = 1 To ilNumberInArray
                ilNameEnd = InStr(1, Trim(slLongNames(c)), " ", vbTextCompare)
               ' slTempString = Trim(Mid(slLongNames(c), 1, ilNameEnd))
                slTempString = Trim(Mid(Trim(slLongNames(c)), 1, ilNameEnd))
               If Not mGetTablesTestIfExists(slTempString) Then
                   tmTableNames(ilCountUniqueNames).sLongName = slTempString
                   ilCountUniqueNames = ilCountUniqueNames + 1
                End If
           Next c
    End Select
    ReDim Preserve tmTableNames(0 To ilCountUniqueNames)
End Sub
Private Sub mArrayFromGlobalArray()
Dim c As Integer
    ReDim tmTableNames(LBound(tgDDFFileNames) To UBound(tgDDFFileNames))
    For c = LBound(tgDDFFileNames) To UBound(tgDDFFileNames) Step 1
        tmTableNames(c) = tgDDFFileNames(c)
    Next c
End Sub
Private Sub mArrayFromStatement(slFromStatement As String, ByRef slLongNames() As String)
Dim c As Integer
    If InStr(1, slFromStatement, ".", vbTextCompare) > 0 Then   ' FROM att.attShfCode INNER JOIN VEF_Vehicles VEF_Vehicles ON att.attVefCode = VEF_Vehicles.vefCode
        emFromType = fromdot
        slLongNames = Split(slFromStatement, ".", , vbTextCompare)
    ElseIf InStr(1, slFromStatement, ",", vbTextCompare) > 0 Then 'FROM  ast, lst, shtt, ADF_Advertisers,
        emFromType = fromcomma
        If InStr(1, slFromStatement, "join", vbTextCompare) > 0 Then 'FROM SVR_7Day_Report, shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode
            slFromStatement = mStripJoins(slFromStatement)
        End If
        slLongNames = Split(slFromStatement, ",", , vbTextCompare)
    Else    'from GRF_Generic_Report INNER JOIN VEF_Vehicles on grfsofCode = vefCode
        emFromType = fromjoin
        slFromStatement = mStripJoins(slFromStatement)
        slLongNames = Split(slFromStatement, ",", , vbTextCompare)
    End If
End Sub
Private Function mStripJoins(slFromStatement As String) As String
    slFromStatement = Replace(slFromStatement, "Inner", "", , , vbTextCompare)
    slFromStatement = Replace(slFromStatement, "Left Outer", "", , , vbTextCompare)
    slFromStatement = Replace(slFromStatement, "Right Outer", "", , , vbTextCompare)
    slFromStatement = Replace(slFromStatement, "Outer", "", , , vbTextCompare)
    slFromStatement = Replace(slFromStatement, "join", ",", , , vbTextCompare)
    mStripJoins = slFromStatement
End Function
Private Sub mSetTableShortNames()
Dim c As Integer
Dim slTempName As String
Dim ilUnderscore As Integer
For c = 0 To UBound(tmTableNames) - 1
    slTempName = Trim(tmTableNames(c).sLongName)
    If Len(slTempName) > 4 Then
        ilUnderscore = InStr(1, slTempName, "_", vbTextCompare)
        If ilUnderscore > 0 Then
            tmTableNames(c).sShortName = Mid(slTempName, 1, ilUnderscore - 1)
        End If
    Else
        tmTableNames(c).sShortName = Trim(tmTableNames(c).sLongName)
    End If
Next c
End Sub
Private Function mAddTablesToFields(tlMyTables() As DDFFILENAMES, slFrom As String) As String
    ' for each short name in mytables, compare to START of a word.  Replace on a match
    Dim slTemp As String
    Dim slAnotherTemp As String
    Dim slNameToSearch As String
    Dim c As Integer
    Dim ilInString As Integer
    Dim ilNewStart As Integer
    Dim slPreviousLetter As String
    Dim ilLengthAddedToString As Integer
    slTemp = slFrom
    For c = 0 To UBound(tlMyTables) - 1
        slNameToSearch = Trim(tlMyTables(c).sShortName)
        If Not StrComp(slNameToSearch, "", vbTextCompare) = 0 Then
            ilLengthAddedToString = Len(slNameToSearch) + Len(Trim(tlMyTables(c).sLongName))
            ilInString = InStr(1, slTemp, slNameToSearch, vbTextCompare)
            Do While ilInString > 0
                ilNewStart = ilInString
                If ilInString > 1 Then
                    slPreviousLetter = Mid(slTemp, ilNewStart - 1, 1)
                Else '
                    slPreviousLetter = " "
                End If ' ilinstring not first char in string
                    If StrComp(slPreviousLetter, " ", vbTextCompare) = 0 Or StrComp(slPreviousLetter, "=", vbTextCompare) = 0 Or StrComp(slPreviousLetter, "(", vbTextCompare) = 0 Then
                        slAnotherTemp = Replace(slTemp, slNameToSearch, Trim(tlMyTables(c).sLongName) & "." & slNameToSearch, ilNewStart, 1, vbTextCompare)
                        mAddTablesAddEnclosuresToField slAnotherTemp
                        slTemp = Mid(slTemp, 1, ilNewStart - 1) & slAnotherTemp
                        ilInString = InStr(ilNewStart + ilLengthAddedToString, slTemp, slNameToSearch, vbTextCompare)
                    Else    'didn't add to string. continue search after last find
                        ilInString = InStr(ilNewStart + 1, slTemp, slNameToSearch, vbTextCompare)
                    End If
           Loop
        End If ' array item not empty
     Next c
    mFixDates slTemp
    mFixTime slTemp
    If ChangeNumbers Then
        mFixNumbers slTemp
    End If
    mAddTablesToFields = slTemp
End Function
Private Sub mAddTablesAddEnclosuresToField(ByRef slstring As String)
Dim ilNextBlank As Integer
Dim ilNextEndChar As Integer
Dim ilNextEnclosureSpot As Integer
    'first blank or first parenthesis end of field, whichever is lowest
    ilNextBlank = InStr(1, slstring, " ", vbTextCompare)
    ilNextEndChar = mFindEarliestEndChar(slstring) 'InStr(1, slstring, ")", vbTextCompare)
    If ilNextBlank = ilNextEndChar Then 'end of string
        ilNextEnclosureSpot = Len(slstring) + 1
    ElseIf ilNextBlank = 0 Then
        ilNextEnclosureSpot = ilNextEndChar
    ElseIf ilNextEndChar = 0 Then
        ilNextEnclosureSpot = ilNextBlank
    ElseIf ilNextBlank < ilNextEndChar Then
        ilNextEnclosureSpot = ilNextBlank
    Else
        ilNextEnclosureSpot = ilNextEndChar
    End If
    If ilNextEnclosureSpot > 0 Then
        slstring = Mid(slstring, 1, ilNextEnclosureSpot - 1) & "}" & Mid(slstring, ilNextEnclosureSpot)
    End If
    slstring = "{" & slstring
End Sub
Private Function mFindEarliestEndChar(slstring As String) As Integer
' these chars mean field is at an end: ) + -
Dim ilEarliestChar As Integer
Dim ilTestChar As Integer
Dim c As Integer
Dim slArrayofEndChars(0 To 3) As String
slArrayofEndChars(0) = "-"
slArrayofEndChars(1) = "+"
slArrayofEndChars(2) = ")"
'dan added 7/20/2011
slArrayofEndChars(3) = ","
ilEarliestChar = InStr(1, slstring, slArrayofEndChars(0), vbTextCompare)
For c = 1 To UBound(slArrayofEndChars)
    ilTestChar = InStr(1, slstring, slArrayofEndChars(c), vbTextCompare)
    If ilTestChar > 0 Then
        If ilEarliestChar = 0 Or ilTestChar < ilEarliestChar Then
            ilEarliestChar = ilTestChar
        End If
    End If
Next c
mFindEarliestEndChar = ilEarliestChar
End Function
Private Sub mFixDates(ByRef slstring As String)
Dim myRegExp As RegExp
Dim myMatches As MatchCollection
Dim myMatch As Match
Dim NewString As String
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
   ' myRegExp.Pattern = "'[1-2][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9]'"
    myRegExp.Pattern = "'\d\d\d\d-\d\d-\d\d'"
    Set myMatches = myRegExp.Execute(slstring)
    For Each myMatch In myMatches
       NewString = Replace(myMatch.Value, "-", ",")
       NewString = Mid(NewString, 2, Len(NewString) - 2)
       NewString = "Date(" & NewString & ")"
       slstring = Replace(slstring, myMatch.Value, NewString)
    Next myMatch
Set myRegExp = Nothing
Set myMatches = Nothing
Set myMatch = Nothing
End Sub
Private Sub mFixTime(ByRef slstring As String)
Dim myRegExp As RegExp
Dim myMatches As MatchCollection
Dim myMatch As Match
Dim NewString As String
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "'[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'"
    Set myMatches = myRegExp.Execute(slstring)
    For Each myMatch In myMatches
       NewString = Replace(myMatch.Value, ":", ",")
       NewString = Mid(NewString, 2, Len(NewString) - 2)
       NewString = "Time(" & NewString & ")"
       slstring = Replace(slstring, myMatch.Value, NewString)
    Next myMatch
Set myRegExp = Nothing
Set myMatches = Nothing
Set myMatch = Nothing
End Sub

Private Sub mFixNumbers(ByRef slstring As String)
Dim myRegExp As RegExp
Dim myMatches As MatchCollection
Dim myMatch As Match
Dim NewString As String
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "'\d*'"
    Set myMatches = myRegExp.Execute(slstring)
    For Each myMatch In myMatches
        NewString = Replace(myMatch.Value, "'", " ")
        If Len(Trim(NewString)) > 0 Then '     ' ' sneaking in here
            slstring = Replace(slstring, myMatch.Value, NewString)
        End If
    Next myMatch
Set myRegExp = Nothing
Set myMatches = Nothing
Set myMatch = Nothing

End Sub
Private Function mGetTablesTestIfExists(slName As String) As Boolean
    Dim c As Integer
    For c = 0 To UBound(tmTableNames)
        If StrComp(slName, Trim(tmTableNames(c).sLongName), vbTextCompare) = 0 Then
            mGetTablesTestIfExists = True
            Exit Function
        End If
    Next c
    mGetTablesTestIfExists = False
End Function
Private Function mPrepForCrystal(slFrom As String) As Boolean
    mPrepForCrystal = True
    ConvertFrom = Trim(slFrom)
    ConvertTo = ""
    If Len(ConvertFrom) = 0 Then
        mPrepForCrystal = False
    End If
End Function
Private Function mStripToWhereClause() As String
    Dim slStripped As String
    Dim ilWhereStart As Integer
    Dim ilWhereEnd As Integer
    Dim ilOrderStart As Integer
    ilWhereStart = InStr(1, ConvertFrom, "Where", vbTextCompare)
    If ilWhereStart > 0 Then
        ilWhereEnd = ilWhereStart + 6
       slStripped = Mid(ConvertFrom, ilWhereEnd)
    End If
    ilOrderStart = InStr(1, slStripped, "Order", vbTextCompare)
    If ilOrderStart > 0 Then
        slStripped = Mid(slStripped, 1, ilOrderStart - 1)
    End If
    mStripToWhereClause = slStripped
End Function

