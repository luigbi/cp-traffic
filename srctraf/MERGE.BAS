Attribute VB_Name = "MERGESUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Merge.bas on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Merge.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the File subs and functions
Option Explicit
Option Compare Text
Type MERGEINFO
    iType As Integer '0=Integer; 1=Long
    iOldCode As Integer
    iNewCode As Integer
    lOldCode As Long
    lNewCode As Long
End Type
Type CNTRINFO
    lCode As Long
    lCntrNo As Long
End Type
'*******************************************************
'*                                                     *
'*      Procedure Name:gSlfCodeExistInChf              *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified code     *
'*                     number exists in specified file *
'*                                                     *
'*******************************************************
Function gSlfCodeExistInChf(Frm As Form, ilCodeNo As Integer) As Integer
'
'   ilRet = gSlfCodeExistInChf (MainForm, iCodeNo)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       iCodeNo (I)- Code number to be checked
'       ilRet (O)- True = Found
'                       False = not found
'
    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlTypeBuff As POPICODE   'Type field record
    Dim tlRecBuff As ARRAYREC
    Dim ilCodeOffset As Integer
    ilCodeOffset = gFieldOffset("Chf", "ChfSlfCode1")
    If ilCodeOffset < 0 Then
        gSlfCodeExistInChf = True
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrOpen):" & " Chf.Btr", Frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    If ilRecLen > Len(tlRecBuff) Then
        ilRecLen = Len(tlRecBuff)
    End If
    llNoRec = btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFile, tlRecBuff, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        On Error GoTo gSlfCodeExistInChfErr
        gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrReset):" & " Chf.Btr", Frm
        On Error GoTo 0
        btrDestroy hlFile
        gSlfCodeExistInChf = False
        Exit Function
    ElseIf ilRet <> BTRV_ERR_DATABUF_LENGTH Then    'Variable length records- return a 22 error- ignore
        On Error GoTo gSlfCodeExistInChfErr
        gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrGetFirst):" & " Chf.Btr", Frm
        On Error GoTo 0
    End If
'    ilCode = Asc(tlRecBuff.sChar(ilCodeOffset)) + 256 * Asc(tlRecBuff.sChar(ilCodeOffset + 1))'Rec definition starts with 2 bytes prior to the array
'    If ilCode = ilCodeNo Then
'        ilRet = btrClose(hlFile)
'        On Error GoTo gIICodeRefExistErr
'        gBtrvErrorMsg ilRet, "gIICodeRefExist (btrReset):" & " Chf.Btr", Frm
'        On Error GoTo 0
'        btrDestroy hlFile
'        gIICodeRefExist = True
'        Exit Function
'    End If
    llNoRec = gExtNoRec(ilRecLen)
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "CHF", "") 'Set extract limits (all records)
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 2, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 4, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 6, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 8, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 10, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 12, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 14, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 16, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset + 18, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTypeBuff, 2)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddLogicConst):" & " Chf.Btr", Frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract iCode field
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrExtAddField):" & " Chf.Btr", Frm
    On Error GoTo 0
    If ddExtRecordCount(hlFile) > 0 Then
        gSlfCodeExistInChf = True
    Else
        gSlfCodeExistInChf = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gSlfCodeExistInChfErr
    gBtrvErrorMsg ilRet, "gSlfCodeExistInChf (btrReset):" & " Chf.Btr", Frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gSlfCodeExistInChfErr:
    gDbg_HandleError "MergeSubs: gSlfCodeExistInChf"
'    gSlfCodeExistInChf = True
End Function
