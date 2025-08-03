Attribute VB_Name = "modStation"

'******************************************************
'*  modStation - various global declarations for Station Input
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text


Type HISTORYINFO
    lCode As Long
    sCallLetters As String * 40
    sEndDate As String * 12
    sDelete As String * 1
End Type

Type EMAILINFO
    lCode As Long
    iSeqNo As Integer
    sEmail As String * 240
    sDelete As String * 1
End Type


Type OWNEDSTATIONS
    sCallLetters As String * 10
    sMktName As String * 40
    sOwnerName As String * 40
    lOwnerCode As Long
    iShttCode As Integer
    iMktCode As Integer
    iMgtCode As Integer
    iSelected As Integer
    iPosition As Integer
    iGroupID As Integer
    lAttCode As Long
End Type

Public igIndex As Integer
Public sgOrigCallLetters As String
Public sgNewCallLetters As String
Public sgLastAirDate As String
Public igHistoryStatus As Integer '0=Save without history; 1=Save with history; 2= Cancel
Public sgOrigTimeZone As String
Public sgNewTimeZone As String
Public sgTimeZoneChangeDate As String
Public igTimeZoneStatus As Integer '0=Remap not required, 1=Remap required, 2= Cancel
Public sgTitle As String
Public bgAffRepCanceled As Boolean
Public bgFrmTitleCanceled As Boolean

Public Function gUpdateRegions(slCategory As String, llFromCode As Long, llToCode As Long, slLogFileName As String) As Integer
    Dim rst_Raf As ADODB.Recordset
    Dim rst_Sef As ADODB.Recordset

    gUpdateRegions = True
    On Error GoTo ErrHand:
    'Handle old form (category within raf).  Old form still used with Network
    SQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = '" & slCategory & "') and (rafType = 'C' OR rafType = 'N'))"
    Set rst_Raf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Raf.EOF
        If llToCode <> -1 Then
            'Check if the To Code exist, if so remove the from code only
            SQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & llToCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            Set rst_Sef = gSQLSelectCall(SQLQuery)
            If rst_Sef.EOF Then
                SQLQuery = "Update sef_Split_Entity Set sefIntCode = " & llToCode & " Where sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                    gUpdateRegions = False
                    On Error Resume Next
                    rst_Raf.Close
                    rst_Sef.Close
                    Exit Function
                End If
            Else
                SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                    gUpdateRegions = False
                    On Error Resume Next
                    rst_Raf.Close
                    rst_Sef.Close
                    Exit Function
                End If
            End If
        Else
            SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                gUpdateRegions = False
                On Error Resume Next
                rst_Raf.Close
                rst_Sef.Close
                Exit Function
            End If
        End If
        rst_Raf.MoveNext
    Loop
    'Handle new form
    SQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = '" & " " & "') and (rafType = 'C' OR rafType = 'N'))"
    Set rst_Raf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Raf.EOF
        If llToCode <> -1 Then
            SQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & llToCode & " AND " & "sefRafCode = " & rst_Raf!rafCode & " AND sefCategory = '" & slCategory & "'"
            Set rst_Sef = gSQLSelectCall(SQLQuery)
            If rst_Sef.EOF Then
                SQLQuery = "Update sef_Split_Entity Set sefIntCode = " & llToCode & " Where sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode & " AND sefCategory = '" & slCategory & "'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                    gUpdateRegions = False
                    On Error Resume Next
                    rst_Raf.Close
                    rst_Sef.Close
                    Exit Function
                End If
            Else
                SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode & " AND sefCategory = '" & slCategory & "'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                    gUpdateRegions = False
                    On Error Resume Next
                    rst_Raf.Close
                    rst_Sef.Close
                    Exit Function
                End If
            End If
        Else
            SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & llFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode & " AND sefCategory = '" & slCategory & "'"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "modStation-gUpdateRegions"
                gUpdateRegions = False
                On Error Resume Next
                rst_Raf.Close
                rst_Sef.Close
                Exit Function
            End If
        End If
        rst_Raf.MoveNext
    Loop

    On Error Resume Next
    rst_Raf.Close
    On Error Resume Next
    rst_Sef.Close
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modStations-gUpdateRegions"
    gUpdateRegions = False
    Exit Function
End Function
