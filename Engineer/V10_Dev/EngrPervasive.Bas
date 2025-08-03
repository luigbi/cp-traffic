Attribute VB_Name = "EngrPervasive"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Pervasive.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for FILE.DDF
Option Explicit

Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
End Type

Public Const DDFFILEPK As String = "IB30B64BB10"
'ttp 5218
Global Const BTRV_ERR_REC_LOCKED = 84
Global Const BTRV_ERR_FILE_LOCKED = 85
Global Const BTRV_ERR_INCOM_LOCK = 93


Dim tmFileDDF As DDFFILE
Dim hmFile As Integer
Dim imFileRecLen As Integer
Public Function gInsertAndReturnCode(slSQLQuery As String, slTable As String, slFieldName As String, slValueToReplace As String) As Long
'   Dan M 9/17/09 Perform insert and return new autoincremented code.
'   I: slSqlQuery an insert command (INSERT INTO CEF_Comments_Events (cefCode,cefComments) VALUES (replace,'This is a test') )
'   I: slTable (CEF_Comments_events)
'   I: slFieldName  (cefCode)
'   I: slValueToReplace (replace) the word that will be replaced with the incremented code value
'   O: autoincremented code number--0 means error
    Dim slMaxQuery As String
    Dim llCode As Long
    Dim slNewQuery As String
    Dim ilRet As Integer
    On Error GoTo ErrHand
    slMaxQuery = "SELECT MAX(" & slFieldName & ") from " & slTable
    Do
        Set rst = cnn.Execute(slMaxQuery)
        'Dan M 9/14/10 take care of Null
        If IsNull(rst(0).Value) Then
            llCode = 1
        Else
            If Not rst.EOF Then
                llCode = rst(0).Value + 1
            Else
                llCode = 1
            End If
        End If
        ilRet = 0
        slNewQuery = Replace(slSQLQuery, slValueToReplace, llCode, , , vbTextCompare)
        If gSQLWaitNoMsgBox(slNewQuery, False) <> 0 Then
            GoTo ErrHand1    'changes ilRet if duplicate value
        End If
    Loop While ilRet <> 0
    gInsertAndReturnCode = llCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            If gErrSQL.NativeError = -4994 Then
                ilRet = 1
                'Return
                Resume Next
            End If
            gMsg = "A SQL error has occurred in EngrPervasive-gInsertAndReturnCode: "
            gLogMsg "Error: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "engrerrors.txt", False
        ElseIf gErrSQL.Number <> 0 Then
            gMsg = "A SQL error has occured in EngrPervasive-gInsertAndReturnCode: "
            gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "engrerrors.txt", False
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in  EngrPervasive-gInsertAndReturnCode: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error # " & Err.Number & "; Line #" & Erl, "engrerrors.txt", False
    End If
    gInsertAndReturnCode = 0
    Exit Function
ErrHand1:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            If gErrSQL.NativeError = -4994 Then
                ilRet = 1
                Return
            End If
            gMsg = "A SQL error has occurred in EngrPervasive-gInsertAndReturnCode: "
            gLogMsg "Error: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "engrerrors.txt", False
        ElseIf gErrSQL.Number <> 0 Then
            gMsg = "A SQL error has occured in EngrPervasive-gInsertAndReturnCode: "
            gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "engrerrors.txt", False
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in  EngrPervasive-gInsertAndReturnCode: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error # " & Err.Number & "; Line #" & Erl, "engrerrors.txt", False
    End If
    gInsertAndReturnCode = 0
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gSQLWait                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Execute Insert or Update or     *
'*                     Delete SQL operation            *
'*                                                     *
'*******************************************************
Public Function gSQLWaitNoMsgBox(sSQLQuery As String, iDoTrans As Integer) As Long
    'Dan m 9/19/09 changed llRet and function return to Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery
        If llRet = 0 Then
            If iDoTrans Then
                cnn.CommitTrans
            End If
        ElseIf (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            fStart = Timer
            Do While Timer <= fStart
                llRet = llRet
            Loop
            iCount = iCount + 1
            If iCount > igWaitCount Then
                'gMsgBox "A SQL error has occurred: " & "Error # " & llRet, vbCritical
                Exit Do
            End If
        End If
    Loop While (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT)
    gSQLWaitNoMsgBox = llRet
    If llRet <> 0 Then
        On Error GoTo mOpenFileErr:
        hlMsg = FreeFile
        Open sgMsgDirectory & "engrerrors.txt" For Append As hlMsg
        Print #hlMsg, sSQLQuery
        Print #hlMsg, "Error # " & llRet
        Close #hlMsg
    End If
    On Error GoTo 0
    Exit Function
    
ErrHand:
    For Each gErrSQL In cnn.Errors
        llRet = gErrSQL.NativeError
        If llRet < 0 Then
            llRet = llRet + 4999
        End If
        'If (llRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        If (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            If iDoTrans Then
                cnn.RollbackTrans
            End If
            cnn.Errors.Clear
            Resume Next
        End If
        'If llRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        '    gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        'End If
    Next gErrSQL
    If llRet = 0 Then
        llRet = Err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    'cnn.Errors.Clear
    Resume Next
mOpenFileErr:
    Resume Next
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gCheckDDFDates                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check DDF dates with DDFOddst.csi*
'*                     and DDFPack.csi                 *
'*                                                     *
'*******************************************************
Public Function gCheckDDFDates() As Integer
    Dim hlFrom As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim slDateTime As String
    Dim slDDFFile As String
    Dim slDDFDateTime As String
    Dim ilPos As Integer
    Dim ilEof As Integer
    Dim slDate1 As String
    Dim slDate2 As String
    Dim slTime1 As String
    Dim llTime1S As Long
    Dim llTime1E As Long
    Dim slTime2 As String
    Dim llTime2 As Long
    Dim llLen As Long
    Dim ilLoop As Integer
    Dim slFolder As String
    Dim ilTVIFound As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffset As Integer
    Dim llRecPos As Long
    Dim slTVIDateTime As String
    
    ilRet = 0
    On Error GoTo gCheckDDFDatesErr:
    llLen = FileLen(sgExeDirectory & "csi_io32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_io32.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    ilRet = 0
    slDDFFile = sgDBPath & "Field.DDF"
    slDDFDateTime = gFileDateTime(slDDFFile)
    If ilRet <> 0 Then
        gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please call Counterpoint", vbCritical, "DDF Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
    slDate1 = Left$(slDDFDateTime, ilPos - 1)
    slFolder = sgDBPath
    ilPos = InStrRev(slFolder, "\", Len(sgDBPath) - 1, vbTextCompare)
    ilRet = 0
    slDDFFile = Left$(slFolder, ilPos) & "NewDDF\Field.DDF"
    slDDFDateTime = gFileDateTime(slDDFFile)
    If ilRet <> 0 Then
        gMsgBox "Unable to find Field.DDF in " & Left$(slFolder, ilPos) & "NewDDF" & ", please call Counterpoint", vbCritical, "DDF Missing"
        gCheckDDFDates = False
        Exit Function
    End If

    ilTVIFound = False
    slTVIDateTime = ""
    hmFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hmFile, "", sgDBPath & "File.DDF", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imFileRecLen = Len(tmFileDDF) 'btrRecordLength(hlAdf)  'Get and save record length
    ilExtLen = Len(tmFileDDF)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hmFile, tmFileDDF, imFileRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet = BTRV_ERR_NONE Then
            Call btrExtSetBounds(hmFile, llNoRec, -1, "UC", "DDFFILEPK", DDFFILEPK) 'Set extract limits (all records)
            ilOffset = 0
            ilRet = btrExtAddField(hmFile, ilOffset, imFileRecLen)  'Extract iCode field
            If ilRet = BTRV_ERR_NONE Then
                'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                        ilExtLen = Len(tmFileDDF)  'Extract operation record size
                        'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            If StrComp(Left$(tmFileDDF.sName, 3), "TVI", vbTextCompare) = 0 Then
                                slTVIDateTime = Trim$(tmFileDDF.sName)
                                ilTVIFound = True
                                Exit Do
                            End If
                            ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            End If
        End If
    End If
    btrDestroy hmFile
    If ilTVIFound Then
        ilTVIFound = False
        hmFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hmFile, "", Left$(slFolder, ilPos) & "NewDDF\File.DDF", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imFileRecLen = Len(tmFileDDF) 'btrRecordLength(hlAdf)  'Get and save record length
        ilExtLen = Len(tmFileDDF)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hmFile   'Clear any previous extend operation
        ilRet = btrGetFirst(hmFile, tmFileDDF, imFileRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet = BTRV_ERR_NONE Then
                Call btrExtSetBounds(hmFile, llNoRec, -1, "UC", "DDFFILEPK", DDFFILEPK) 'Set extract limits (all records)
                ilOffset = 0
                ilRet = btrExtAddField(hmFile, ilOffset, imFileRecLen)  'Extract iCode field
                If ilRet = BTRV_ERR_NONE Then
                    'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
                    ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                            ilExtLen = Len(tmFileDDF)  'Extract operation record size
                            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Loop
                            Do While ilRet = BTRV_ERR_NONE
                                If StrComp(Left$(tmFileDDF.sName, 3), "TVI", vbTextCompare) = 0 Then
                                    'Compare Dates and time
                                    If StrComp(slTVIDateTime, Trim$(tmFileDDF.sName), vbTextCompare) <> 0 Then
                                        gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbCritical, "DDF Problem"
                                        gCheckDDFDates = False
                                        Exit Function
                                    End If
                                    ilTVIFound = True
                                    Exit Do
                                End If
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                                Loop
                            Loop
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not ilTVIFound Then
        ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
        slDate2 = Left$(slDDFDateTime, ilPos - 1)
        If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
            gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbCritical, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    
    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
    'is about 900000.  The other is about 90000.
    If llLen > 200000 Then
        ilRet = 0
        llLen = FileLen(sgExeDirectory & "csi_os32.dll")
        If ilRet <> 0 Then
            gMsgBox "Unable to find csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Missing"
            gCheckDDFDates = False
            Exit Function
        End If
        If llLen < 20000 Then
            gMsgBox "Incompatible version of csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Incompatible"
            gCheckDDFDates = False
            Exit Function
        End If
        gCheckDDFDates = True
        Exit Function
    End If
    ilRet = 0
    llLen = FileLen(sgExeDirectory & "csi_os32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    If llLen > 20000 Then
        gMsgBox "Incompatible version of csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    ilRet = 0
    llLen = FileLen(sgExeDirectory & "cbtrv432.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find cbtrv432.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    If llLen < 400000 Then
        gMsgBox "Incompatible version of cbtrv.dll in " & sgExeDirectory & ", please call Counterpoint", vbCritical, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    If Not ilTVIFound Then
        ilRet = 0
        slDDFFile = sgDBPath & "Field.DDF"
        slDDFDateTime = gFileDateTime(slDDFFile)
        If ilRet <> 0 Then
            gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please place File in folder and run DDFOffst.exe", vbCritical, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
        ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
        slDate1 = Left$(slDDFDateTime, ilPos - 1)
        slTime1 = Mid$(slDDFDateTime, ilPos + 1)
        llTime1S = gTimeToLong(slTime1, False) - 10800    '3 hours
        llTime1E = gTimeToLong(slTime1, False) + 10800    '3 hours
    Else
        slDate1 = slTVIDateTime
    End If
    'Test Offset table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        hlFrom = FreeFile
        Open sgDBPath & "DDFOffst.csi" For Input Access Read Shared As hlFrom
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFOffst.csi" & " Error " & Str$(ilRet), vbCritical, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    Do
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                ilPos = InStr(1, slLine, "'DDF Date", vbTextCompare)
                If ilPos = 1 Then
                    slDateTime = Trim$(Mid$(slLine, 11))
                    Exit Do
                End If
            End If
        End If
    Loop Until ilEof
    If slDateTime = "" Then
        Close hlFrom
        gMsgBox "Unable to find DDF Date line in DDFOffst.csi, please run DDFOffst.exe", vbCritical, "DDFOffst.Csi"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        slDate2 = Left$(slDateTime, ilPos - 1)
        slTime2 = Mid$(slDateTime, ilPos + 1)
        llTime2 = gTimeToLong(slTime2, False)
        If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
            gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDFOffst.Csi"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If
    'Test Pack table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        hlFrom = FreeFile
        Open sgDBPath & "DDFPack.csi" For Input Access Read Shared As hlFrom
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFPack.csi" & " Error " & Str$(ilRet), vbCritical, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    Do
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                ilPos = InStr(1, slLine, "'DDF Date", vbTextCompare)
                If ilPos = 1 Then
                    slDateTime = Trim$(Mid$(slLine, 11))
                    Exit Do
                End If
            End If
        End If
    Loop Until ilEof
    If slDateTime = "" Then
        Close hlFrom
        gMsgBox "Unable to find DDF Date line in DDFPack.csi, please run DDFOffst.exe", vbCritical, "DDF Offset"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        slDate2 = Left$(slDateTime, ilPos - 1)
        slTime2 = Mid$(slDateTime, ilPos + 1)
        llTime2 = gTimeToLong(slTime2, False)
        If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
            gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDF Pack"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDF Pack"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbCritical, "DDF Pack"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If

    gCheckDDFDates = True
    Exit Function
gCheckDDFDatesErr:
    ilRet = Err.Number
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gExtNoRec                       *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute max # records for       *
'*                     btrieve extend operation        *
'*                                                     *
'*            Formula: # Rec = 60000/(6+RecSize)       *
'*                     6= description size added to    *
'*                        each record extracted and    *
'*                        into the return buffer       *
'*                                                     *
'*******************************************************
Function gExtNoRec(ilRecSize As Integer) As Long
    gExtNoRec = 8000 \ (6 + ilRecSize)  'Change 60000 to 8000
End Function


Public Sub gHandleError(slLogName As String, slMethodName As String)
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, slLogName, False
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical, ""
        ElseIf gErrSQL.Number <> 0 Then
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, slLogName, False
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, vbCritical, ""
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in " & slMethodName & ": "
        gLogMsg "ERROR: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, slLogName, False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical, ""
    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gSQLWait                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Execute Insert or Update or     *
'*                     Delete SQL operation            *
'*                                                     *
'*                     This routine is used by Code    *
'*                     Protect                         *
'*******************************************************
Public Function gSQLWaitNoMsgBoxEX(sSQLQuery As String, iDoTrans As Integer, slModNameLineNo As String) As Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery
        If llRet = 0 Then
            If iDoTrans Then
                cnn.CommitTrans
            End If
        ElseIf (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            fStart = Timer
            Do While Timer <= fStart
                llRet = llRet
            Loop
            iCount = iCount + 1
            If iCount > igWaitCount Then
                'gMsgBox "A SQL error has occurred: " & "Error # " & llRet, vbCritical
                Exit Do
            End If
        End If
    Loop While (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT)
    gSQLWaitNoMsgBoxEX = llRet
    If llRet <> 0 Then
        'On Error GoTo mOpenFileErr:
        'hlMsg = FreeFile
        'Open sgMsgDirectory & "engrerrors.txt" For Append As hlMsg
        'Print #hlMsg, sSQLQuery
        'Print #hlMsg, slModNameLineNo & " Error # " & llRet
        'Close #hlMsg
        gLogMsg slModNameLineNo & " Error # " & llRet, "engrerrors.txt", False
    End If
    On Error GoTo 0
    Exit Function
    
ErrHand:
    For Each gErrSQL In cnn.Errors
        llRet = gErrSQL.NativeError
        If llRet < 0 Then
            llRet = llRet + 4999
        End If
        'If (llRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        If (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            If iDoTrans Then
                cnn.RollbackTrans
            End If
            cnn.Errors.Clear
            Resume Next
        End If
        'If llRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        '    gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        'End If
    Next gErrSQL
    If llRet = 0 Then
        llRet = Err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    'cnn.Errors.Clear
    Resume Next
mOpenFileErr:
    Resume Next
End Function
