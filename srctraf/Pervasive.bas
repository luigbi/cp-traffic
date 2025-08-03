Attribute VB_Name = "Pervasive"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Pervasive.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

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

Global igShowMsgBox As Integer

'D.S. 07/20/15
'ADO variables
Public cnn As ADODB.Connection
Public rst As ADODB.Recordset
Public rst2 As ADODB.Recordset
Public gErrSQL As ADODB.Error

'D.S. 07/20/15
'SQL variables
Public SQLQuery As String
Public sgDatabaseName As String
Public sgSQLDateForm As String
Public sgSQLTimeForm As String
'8199
Public igWaitCount As Integer
Public bgIgnoreDuplicateError As Boolean
Public gMsg As String

Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
End Type

Public Const DDFFILEPK As String = "IB30B64BB10"

Dim tmFileDDF As DDFFILE
Dim hmFile As Integer
Dim imFileRecLen As Integer
Public Function mOpenPervasiveAPI() As Integer
    
    Dim hgDB As Integer
    
    sgMDBPath = ""
    sgSDBPath = ""
    sgTDBPath = sgDBPath
    igRetrievalDB = 0
    
    'hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    'hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, "", igRetrievalDB, "") 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    Do While csiHandleValue(0, 3) = 0
        '7/6/11
        Sleep 1000
    Loop

    If hgDB <> 0 Then
        gMsgBox "CBtrvMngrInit Failed"
        mOpenPervasiveAPI = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    mOpenPervasiveAPI = True
    
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
Public Function gCheckDDFDates(Optional blShowMsg As Boolean = True) As Integer
    'D.S. 5/31/13 Added optional blShowMsg parameter above

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

    sgDDFDateInfo = ""
    ilRet = 0
    'On Error GoTo gCheckDDFDatesErr:
    'llLen = FileLen(sgExePath & "csi_io32.dll")
    ilRet = gFileExist(sgExePath & "csi_io32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_io32.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    llLen = FileLen(sgExePath & "csi_io32.dll")
    ilRet = 0
    slDDFFile = sgDBPath & "Field.DDF"
    slDDFDateTime = gFileDateTime(slDDFFile)
    If ilRet <> 0 Then
        gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please call Counterpoint", vbExclamation, "DDF Missing"
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
        gMsgBox "Unable to find Field.DDF in " & Left$(slFolder, ilPos) & "NewDDF" & ", please call Counterpoint", vbExclamation, "DDF Missing"
        gCheckDDFDates = False
        Exit Function
    End If

    ilTVIFound = False
    slTVIDateTime = ""
    hmFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hmFile, "", sgDBPath & "File.DDF", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'D.S. 08/15/12 Added full path information and well as the file name
        'gMsgBox "Unable to open file.DDF, Error # = " & ilRet, vbExclamation, "DDF Missing"
        'D.S. 5/31/13 Don't show the message box unnecessarily
        If blShowMsg <> False Then
            gMsgBox "Unable to open: [" & sgDBPath & "File.DDF" & "] Error # = " & ilRet, vbExclamation, "DDF Missing"
        End If
        gCheckDDFDates = False
        Exit Function
    End If

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
        sgDDFDateInfo = Mid$(slTVIDateTime, 5, 2) & "/" & Mid$(slTVIDateTime, 7, 2) & "/" & Mid$(slTVIDateTime, 9, 2) & " at " & Mid$(slTVIDateTime, 11, 2) & ":" & Mid$(slTVIDateTime, 13, 2) & " " & Trim(Mid$(slTVIDateTime, 15))
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
                                        gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbExclamation, "DDF Problem"
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
        btrDestroy hmFile
    End If
    If Not ilTVIFound Then
        If Trim$(slDDFDateTime) <> "" Then
            ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
            If ilPos > 1 Then
                slDate2 = Left$(slDDFDateTime, ilPos - 1)
                If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
                    gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbExclamation, "DDF Problem"
                    gCheckDDFDates = False
                    Exit Function
                End If
            Else
                gMsgBox "Restart Traffic, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    End If

    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
    'is about 900000.  The other is about 90000.
    If llLen > 200000 Then
        '2/17/15: if using the new csi_io32 routine, than the csi_os32 routine is not required
'        ilRet = 0
'        llLen = FileLen(sgExePath & "csi_os32.dll")
'        If ilRet <> 0 Then
'            gMsgBox "Unable to find csi_os32.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
'            gCheckDDFDates = False
'            Exit Function
'        End If
'        If llLen < 20000 Then
'            gMsgBox "Incompatible version of csi_os32.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
'            gCheckDDFDates = False
'            Exit Function
'        End If
        gCheckDDFDates = True
        Exit Function
    End If
    ilRet = 0
    'llLen = FileLen(sgExePath & "csi_os32.dll")
    ilRet = gFileExist(sgExePath & "csi_os32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_os32.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    llLen = FileLen(sgExePath & "csi_os32.dll")
    If llLen > 20000 Then
        gMsgBox "Incompatible version of csi_os32.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    ilRet = 0
    'llLen = FileLen(sgExePath & "cbtrv432.dll")
    ilRet = gFileExist(sgExePath & "cbtrv432.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find cbtrv432.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    llLen = FileLen(sgExePath & "cbtrv432.dll")
    If llLen < 400000 Then
        gMsgBox "Incompatible version of cbtrv.dll in " & sgExePath & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    If Not ilTVIFound Then
        ilRet = 0
        slDDFFile = sgDBPath & "Field.DDF"
        slDDFDateTime = gFileDateTime(slDDFFile)
        If ilRet <> 0 Then
            gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please place File in folder and run DDFOffst.exe", vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
        ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate1 = Left$(slDDFDateTime, ilPos - 1)
            slTime1 = Mid$(slDDFDateTime, ilPos + 1)
            llTime1S = gTimeToLong(slTime1, False) - 10800    '3 hours
            llTime1E = gTimeToLong(slTime1, False) + 10800    '3 hours
        Else
            gMsgBox "Restart Traffic, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate1 = slTVIDateTime
    End If
    'Test Offset table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        'hlFrom = FreeFile
        'Open sgDBPath & "DDFOffst.csi" For Input Access Read Shared As hlFrom
        ilRet = gFileOpen(sgDBPath & "DDFOffst.csi", "Input Access Read Shared", hlFrom)
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFOffst.csi" & " Error " & Str$(ilRet), vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    err.Clear
    Do
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        If EOF(hlFrom) Then
            Exit Do
        End If
        Line Input #hlFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
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
        gMsgBox "Unable to find DDF Date line in DDFOffst.csi, please run DDFOffst.exe", vbExclamation, "DDFOffst.Csi"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate2 = Left$(slDateTime, ilPos - 1)
            slTime2 = Mid$(slDateTime, ilPos + 1)
            llTime2 = gTimeToLong(slTime2, False)
            If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
                gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Restart Traffic, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If
    'Test Pack table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        'hlFrom = FreeFile
        'Open sgDBPath & "DDFPack.csi" For Input Access Read Shared As hlFrom
        ilRet = gFileOpen(sgDBPath & "DDFPack.csi", "Input Access Read Shared", hlFrom)
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFPack.csi" & " Error " & Str$(ilRet), vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    err.Clear
    Do
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        If EOF(hlFrom) Then
            Exit Do
        End If
        Line Input #hlFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
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
        gMsgBox "Unable to find DDF Date line in DDFPack.csi, please run DDFOffst.exe", vbExclamation, "DDF Offset"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate2 = Left$(slDateTime, ilPos - 1)
            slTime2 = Mid$(slDateTime, ilPos + 1)
            llTime2 = gTimeToLong(slTime2, False)
            If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
                gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Restart Traffic, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If

    gCheckDDFDates = True
    Exit Function
'gCheckDDFDatesErr:
'    ilRet = Err.Number
'    Resume Next
End Function
Public Function gSQLWaitNoMsgBox(sSQLQuery As String, iDoTrans As Integer) As Long
    '8199
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand
    'Dan removed 8199
'    '12/4/12: Check if activity should be logged
'    mLogActivityFileName sSQLQuery
'    '12/4/12: end of change
    
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
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            On Error GoTo mOpenFileErr:
            gLogMsg sSQLQuery, "TrafficErrors.Txt", False
            gLogMsg "Error # " & llRet, "TrafficErrors.Txt", False
            'dan removed
'            hlMsg = FreeFile
'            Open sgMsgDirectory & "TrafficErrors.Txt" For Append As hlMsg
'            Print #hlMsg, sSQLQuery
'            Print #hlMsg, "Error # " & llRet
'            Close #hlMsg
        End If
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
        llRet = err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    'cnn.Errors.Clear
    Resume Next
mOpenFileErr:
    Resume Next
End Function
Public Function gSQLWaitNoMsgBoxEX(sSQLQuery As String, iDoTrans As Integer, slModNameLineNo As String) As Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand
    'Dan removed 8199
'    '12/4/12: Check if activity should be logged
'    mLogActivityFileName sSQLQuery
'    '12/4/12: end of change
    
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
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            'On Error GoTo mOpenFileErr:
            'hlMsg = FreeFile
            'Open sgMsgDirectory & "AffErrorLog.Txt" For Append As hlMsg
            'Print #hlMsg, sSQLQuery
            'Print #hlMsg, slModNameLineNo & " Error # " & llRet
            'Close #hlMsg
            gLogMsg slModNameLineNo & " Error # " & llRet, "TrafficErrors.Txt", False
        End If
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
        llRet = err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    'cnn.Errors.Clear
    Resume Next
mOpenFileErr:
    Resume Next
End Function

Public Sub gHandleError(slLogName As String, slMethodName As String)
'8199
'General routine to be used in error handler of mehtods with sql calls:
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gHandleError LOGFILE, "Export IDC-mCleanIef"
'    mCleanIef = False
'   always write to TrafficErrors.txt.  Unfortunately, gmsgbox does this if igBkgdProg <> 0 ( affiliate: igShowMsgBox = 0)
'   write to alternate if slLogName is included and not TrafficErrors.Txt
    Dim blIsAlternateLog As Boolean
    
    'we have an alternate log. always write it out.
    If UCase(slLogName) = "TrafficErrors.txt" Or Len(slLogName) = 0 Then
        blIsAlternateLog = False
    Else
        blIsAlternateLog = True
    End If
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            sgTmfStatus = "E"
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, slLogName, False
            End If
            'Dan
            If igBkgdProg = 0 Then
            'If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "TrafficErrors.Txt", False
            End If
        ElseIf gErrSQL.Number <> 0 Then
            sgTmfStatus = "E"
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, slLogName, False
            End If
            If igBkgdProg = 0 Then
            'If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "TrafficErrors.Txt", False
            End If
        End If
    Next gErrSQL
    If (err.Number <> 0) And (gMsg = "") Then
        sgTmfStatus = "E"
        gMsg = "A general error has occured in " & slMethodName & ": "
        gMsgBox gMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        If blIsAlternateLog Then
            gLogMsg "ERROR: " & gMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, slLogName, False
        End If
        If igBkgdProg = 0 Then
        'If igShowMsgBox <> 0 Then
             gLogMsg "ERROR: " & gMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, "TrafficErrors.Txt", False
        End If

    End If

End Sub
Public Function gSQLAndReturn(sSQLQuery As String, iDoTrans As Integer, llAffectedRecords As Long) As Long
    '8199
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand

    '12/4/12: Check if activity should be logged
    'mLogActivityFileName sSQLQuery
    '12/4/12: end of change
    llAffectedRecords = 0
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery, llAffectedRecords
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
    gSQLAndReturn = llRet
    If llRet <> 0 Then
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            On Error GoTo mOpenFileErr:
            'Dan replaced from affiliate:
            gLogMsg SQLQuery, "trafficErrors.txt", False
            gLogMsg "Error # " & llRet, "trafficErrors.txt", False
'            hlMsg = FreeFile
'            Open sgMsgDirectory & "TrafficErrors.txt" For Append As hlMsg
'            Print #hlMsg, sSQLQuery
'            Print #hlMsg, "Error # " & llRet
'            Close #hlMsg
        End If
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
        llRet = err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    'cnn.Errors.Clear
    Resume Next
mOpenFileErr:
    Resume Next
End Function
Public Function gLoadOptionTrafficThenAffiliate(slInSection As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    Dim slFileName As String
    Dim Section As String

    Section = slInSection
    If igDirectCall = -1 Then
        slFileName = sgIniPath & "Traffic.Ini"
    Else
        slFileName = CurDir$ & "\Traffic.Ini"
    End If
    gLoadOptionTrafficThenAffiliate = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadOptionTrafficThenAffiliate = True
        Else
            If igDirectCall = -1 Then
                slFileName = sgIniPath & "Affiliat.Ini"
            Else
                slFileName = CurDir$ & "\Affiliat.Ini"
            End If
            'affiliate has a different section
            If Key = "Name" Then
                If igTestSystem Then
                    Section = "TestDatabase"
                Else
                    Section = "Database"
                End If
            End If
            BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slFileName)
            If BytesCopied > 0 Then
                If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                    sValue = Left(sBuffer, BytesCopied)
                    gLoadOptionTrafficThenAffiliate = True
                End If
            End If
        End If
    Else
        'not sure if this ever gets called
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "Affiliat.Ini"
        Else
            slFileName = CurDir$ & "\Affiliat.Ini"
        End If
        If Key = "Name" Then
            If igTestSystem Then
                Section = "TestDatabase"
            Else
                Section = "Database"
            End If
        End If
        BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slFileName)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                sValue = Left(sBuffer, BytesCopied)
                gLoadOptionTrafficThenAffiliate = True
            End If
        End If
    End If
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function

Public Function gSQLSelectCall(slSQLQuery As String, Optional slMsg As String = "") As ADODB.Recordset
    On Error GoTo ErrHand
    Set gSQLSelectCall = cnn.Execute(slSQLQuery)
    Exit Function
ErrHand:
    If slMsg <> "" Then
        gHandleError "TrafficErrors.txt", slMsg & " " & slSQLQuery
    Else
        gHandleError "TrafficErrors.txt", slSQLQuery
    End If
End Function
Public Sub gSQLCallIgnoreError(slSQLQuery As String)
    Dim rst As ADODB.Recordset
    On Error Resume Next
    Set rst = cnn.Execute(slSQLQuery)
    Exit Sub
End Sub

Public Function gInsertAndReturnCode(slSQLQuery As String, slTable As String, slFieldName As String, slValueToReplace As String, Optional blTestRepeatingMax As Boolean = False) As Long
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
    Dim llPrevCode As Long
    On Error GoTo ErrHand
    
    'LogActivityFileName slSQLQuery
    llPrevCode = -1
    slMaxQuery = "SELECT MAX(" & slFieldName & ") from " & slTable
    Do
        Set rst = gSQLSelectCall(slMaxQuery)
        If IsNull(rst(0).Value) Then
            llCode = 1
        Else
            If Not rst.EOF Then
                llCode = rst(0).Value + 1
            Else
                llCode = 1
            End If
        End If
        If (llCode = llPrevCode) And (blTestRepeatingMax) Then
            gInsertAndReturnCode = -1
            Exit Function
        End If
        llPrevCode = llCode
        ilRet = 0
        slNewQuery = Replace(slSQLQuery, slValueToReplace, llCode, , , vbTextCompare)
        bgIgnoreDuplicateError = True
        If gSQLWaitNoMsgBox(slNewQuery, False) <> 0 Then
            bgIgnoreDuplicateError = False
            If Not gHandleError4994("AffErrorLog.txt", "modPervasive-gInsertAndReturnCode") Then
                gInsertAndReturnCode = -1
                Exit Function
            End If
            ilRet = 1
        End If
        bgIgnoreDuplicateError = False
    Loop While ilRet <> 0
    gInsertAndReturnCode = llCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "modPervasive-gInsertAndReturnCode"
    gInsertAndReturnCode = 0
    Exit Function
ErrHand1:
    Screen.MousePointer = vbDefault
    If gHandleError4994("", "modPervasive-gInsertAndReturnCode") Then
        ilRet = 1
        Return
    End If
    gInsertAndReturnCode = 0
End Function

Public Function gHandleError4994(slLogName As String, slMethodName As String) As Boolean
    'DUPLICATE KEYS?  Return true  ttp 5217
    Dim blIsAlternateLog As Boolean
    
    If UCase(slLogName) = "AFFERRORLOG.TXT" Or Len(slLogName) = 0 Then
        blIsAlternateLog = False
    Else
        blIsAlternateLog = True
    End If
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            If gErrSQL.NativeError = -4994 Or gErrSQL.NativeError = 5 Then
                gHandleError4994 = True
                Exit Function
            End If
            gMsg = "A SQL error has occurred in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            If blIsAlternateLog Then
                gLogMsg "Error: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        ElseIf gErrSQL.Number <> 0 Then
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        End If
    Next gErrSQL
    If (err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in " & slMethodName & ": "
        gMsgBox gMsg & err.Description & "; Error #" & err.Number, vbCritical
        If blIsAlternateLog Then
            gLogMsg "ERROR: " & gMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, slLogName, False
        End If
        If igShowMsgBox <> 0 Then
             gLogMsg "ERROR: " & gMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, "affErrorLog.txt", False
        End If
    End If
    gHandleError4994 = False
End Function
Public Function gSqlSafeAndTrim(slString As String) As String
        gSqlSafeAndTrim = Trim(Replace(slString, "'", "''"))
End Function
