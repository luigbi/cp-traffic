Attribute VB_Name = "SPECRPTSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Specrpt.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSpec.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Report Specification File Description
Option Explicit
Option Compare Text
'********************************************************
'
'Report Name file definition
'
'*********************************************************
Type RNF
    iCode As Integer        'AutoInc
    sType As String * 1     'C=Category; R=Report
    sName As String * 60    'Name of Category or Report
    sRptExe As String * 12  'Exe Selection module for Report
    sRptSample As String * 12   'Picture of Report (report directory is used)
    sMoneyShown As String * 1   'Y=Money is shown and can be turn off; N=Money Doesn't show or it can't be eliminated from the report
    sPassValue As String * 10   'Pass value to EXE
    sState As String * 1        'A=Active; D=Dormnant
    iJobListNo As Integer       'Job or List number- if defined then used to move
                                'the report to the top of the tree
    'iStrLen As Integer  'String length (required by LVar)
    'sDescription As String * 1002   'Last two bytes after the comment must be 0
    sDescription As String * 1004   'Last bytes after the comment must be 0
End Type
'RNF key record layout- use INTKEY0
'Type RNFKEY0
'    iCode As Integer
'End Type
Type RNFLIST
    sKey As String * 80
    tRnf As RNF
    iCode As Integer
End Type
'********************************************************
'
'Report Tree file definition
'
'*********************************************************
Type RTF
    iCode As Integer        'AutoInc
    iRnfCode As Integer     'RNF Code
    iLevel As Integer       'Level
    iNextRtfCode As Integer 'Link List
    sAutoReturn As String * 1   'Y=Return to Originating EXE; N=Return to Report Selection EXE
    lColor As Long          'Color of the List item
    sRnfType As String * 1  'Rnf.sType (C=Catorgies; R=Report)
    sRnfState As String * 1 'Rnf.sState (A=Active; D=Dormant)
End Type
'RTF key record layout- use INTKEY0
'Type RTFKEY0
'    iCode As Integer
'End Type
Type RTFLIST
    sKey As String * 4
    tRtf As RTF
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    lRecPos As Long
End Type
'********************************************************
'
'Set Name file definition
'
'*********************************************************
Type SNF
    iCode As Integer        'AutoInc
    sName As String * 60    'Name of Category or Report
    sState As String * 1    'A=Active; D=Dormant
    'iStrLen As Integer  'String length (required by LVar)
    'sDescription As String * 1002   'Last two bytes after the comment must be 0
    sDescription As String * 1004   'Last bytes after the comment must be 0
End Type
'SNF key record layout- use INTKEY0
'Type SNFKEY0
'    iCode As Integer
'End Type
Type SNFCODE
    sKey As String * 60
    tSnf As SNF
    lRecPos As Long
End Type
'********************************************************
'
'Set Name Reports file definition
'
'*********************************************************
Type SRF
    lCode As Long           'Unique value (AutoInc)
    iSnfCode As Integer     'Snf Code (Set Name)
    iRnfCode As Integer     'Rnf Code (Report name)
    sViewMoney As String * 1   'Y=Money is shown and can be turn off; N=Money Doesn't show or it can't be eliminated from the report
End Type
'SNF key record layout- use LONGKEY0
'Type SRFKEY0
'    lCode As Long
'End Type
'SRF key record layout- use INTKEY0
'Type SRFKEY1
'    iSnfCode As Integer
'End Type
Type RPTLST
    sName As String * 100
    iLevel As Integer
    tRnf As RNF
End Type
Type RPTNAMEMAP
    sName As String * 60
    iRptCallType As Integer
End Type
Public sgRptSetName As String
Public igRptSetReturn As Integer      '0=Ok; 1=Cancel
Public tgRnfList() As RNFLIST
Public tgRtfList() As RTFLIST
Public tgDelRtfList() As RTFLIST
Dim tmRnf As RNF
Public tgSnfCode() As SNFCODE
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRNF                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Sub gObtainRNF(hlRnf As Integer)
    Dim ilRecLen As Integer
    Dim ilSortCode As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilExtLen As Integer
    'slStamp = gFileDateTime(sgDBPath & "Rnf.Btr")
    'If sgRnfStamp <> "" Then
    '    If StrComp(slStamp, sgRnfStamp, 1) = 0 Then
    '        'If UBound(tgCompMnf) > 1 Then
    '            Exit Sub
    '        'End If
    '    End If
    '    slGetStamp = ""
    'Else
    '    slGetStamp = gGetCSIStamp("RNF")
    'End If
    'If slGetStamp <> "" Then
    '    sgRnfStamp = slGetStamp
    '    ilRet = csiGetAlloc("RNF", ilStartIndex, ilEndIndex)
    '    ReDim tgRNFLIST(ilStartIndex To ilEndIndex) As RNFLIST
    '    For ilLoop = LBound(tgRNFLIST) To UBound(tgRNFLIST) Step 1
    '        ilRet = csiGetRec("RNF", ilLoop, VarPtr( tgRNFLIST(ilLoop)), LenB( tgRNFLIST(ilLoop)))
    '    Next ilLoop
    'Else
        ilRecLen = Len(tmRnf) 'btrRecordLength(hlRnf)  'Get and save record length
        ilSortCode = 0
        ReDim tgRnfList(0 To 0) As RNFLIST   'VB list box clear (list box used to retain code number so record can be found)
        'Can't use extended operations and retrieve lvar field
        'ilExtLen = Len(tmRnf)  'Extract operation record size
        'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRnf) 'Obtain number of records
        'btrExtClear hlRnf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlRnf, tmRnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
        'Call btrExtSetBounds(hlRnf, llNoRec, -1, "UC") 'Set extract limits (all records)
        'ilOffset = 0
        'ilRet = btrExtAddField(hlRnf, ilOffset, ilExtLen)  'Extract First Name field
        'If ilRet <> BTRV_ERR_NONE Then
        '    Exit Sub
        'End If
        'ilExtLen = Len(tmRnf)  'Extract operation record size
        'ilRet = btrExtGetNext(hlRnf, tmRnf, ilExtLen, llRecPos)
        'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        '    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
        '        Exit Sub
        '    End If
        '    ilExtLen = Len(tmRnf)  'Extract operation record size
        '    Do While ilRet = BTRV_ERR_REJECT_COUNT
        '        ilExtLen = Len(tmRnf)  'Extract operation record size
        '        ilRet = btrExtGetNext(hlRnf, tmRnf, ilExtLen, llRecPos)
        '    Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = tmRnf.sName
                If tmRnf.sType = "C" Then
                    slName = "C|" & slName & "|" & tmRnf.sState & "\" & Trim$(str$(tmRnf.iCode))
                Else
                    slName = "R|" & slName & "|" & tmRnf.sState & "\" & Trim$(str$(tmRnf.iCode))
                End If
                tgRnfList(ilSortCode).sKey = slName
                tgRnfList(ilSortCode).tRnf = tmRnf
                If ilSortCode >= UBound(tgRnfList) Then
                    ReDim Preserve tgRnfList(0 To UBound(tgRnfList) + 100) As RNFLIST
                End If
                ilSortCode = ilSortCode + 1
                ilExtLen = Len(tmRnf)  'Extract operation record size
       '         ilRet = btrExtGetNext(hlRnf, tmRnf, ilExtLen, llRecPos)
       '         Do While ilRet = BTRV_ERR_REJECT_COUNT
       '             ilExtLen = Len(tmRnf)  'Extract operation record size
       '             ilRet = btrExtGetNext(hlRnf, tmRnf, ilExtLen, llRecPos)
       '         Loop
                ilRecLen = Len(tmRnf) 'btrRecordLength(hlRnf)  'Get and save record length
                ilRet = btrGetNext(hlRnf, tmRnf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            ReDim Preserve tgRnfList(0 To ilSortCode) As RNFLIST
            If UBound(tgRnfList) - 1 > 0 Then
                ArraySortTyp fnAV(tgRnfList(), 0), UBound(tgRnfList), 0, LenB(tgRnfList(0)), 0, LenB(tgRnfList(0).sKey), 0
            End If
        'End If
        'ilRet = csiSetStamp("RNF", sgRnfStamp)
        'ilRet = csiSetAlloc("RNF", LBound(tgRNFLIST), UBound(tgRNFLIST))
        'For ilLoop = LBound(tgRNFLIST) To UBound(tgRNFLIST) Step 1
        '    ilRet = csiSetRec("RNF", ilLoop, VarPtr( tgRNFLIST(ilLoop)), LenB( tgRNFLIST(ilLoop)))
        'Next ilLoop
    'End If

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRTF                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Sub gObtainRTF(hlRtf As Integer, ilSaveDormant As Integer)
'
'   gObtainRTF hlRtf, ilSaveDormant
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilRecLen As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffset As Integer
    Dim slType As String
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilFind As Integer
    Dim ilFound As Integer
    Dim ilCount As Integer
    Dim ilDel As Integer
    Dim slStr As String
    Dim slState As String
    Dim ilRtf As Integer
    ReDim tgRtfList(0 To 0) As RTFLIST   'VB list box clear (list box used to retain code number so record can be found)
    If ilSaveDormant Then
        ReDim tgDelRtfList(0 To 0) As RTFLIST   'VB list box clear (list box used to retain code number so record can be found)
    End If
    ilUpper = 0
    ilRecLen = Len(tgRtfList(ilUpper).tRtf) 'btrRecordLength(hlRtf)  'Get and save record length
    ilExtLen = Len(tgRtfList(ilUpper).tRtf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRtf) 'Obtain number of records
    btrExtClear hlRtf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRtf, tgRtfList(ilUpper).tRtf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlRtf, llNoRec, -1, "UC", "RTF", "") 'Set extract limits (all records)
    ilOffset = 0
    ilRet = btrExtAddField(hlRtf, ilOffset, ilExtLen)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hlRtf)    'Extract record
    ilUpper = UBound(tgRtfList)
    ilRet = btrExtGetNext(hlRtf, tgRtfList(ilUpper).tRtf, ilExtLen, tgRtfList(ilUpper).lRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
        ilUpper = UBound(tgRtfList)
        ilExtLen = Len(tgRtfList(ilUpper))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlRtf, tgCompMnf(ilUpper), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRtf, tgRtfList(ilUpper).tRtf, ilExtLen, tgRtfList(ilUpper).lRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            tgRtfList(ilUpper).iStatus = 1
            ilUpper = ilUpper + 1
            ReDim Preserve tgRtfList(0 To ilUpper) As RTFLIST
            ilRet = btrExtGetNext(hlRtf, tgRtfList(ilUpper).tRtf, ilExtLen, tgRtfList(ilUpper).lRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRtf, tgRtfList(ilUpper), ilExtLen, tgRtfList(ilUpper).lRecPos)
            Loop
        Loop
        'Find end- then work back to start
        ilFind = -1
        Do
            ilFound = False
            For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
                If tgRtfList(ilLoop).tRtf.iNextRtfCode = ilFind Then
                    ilFind = tgRtfList(ilLoop).tRtf.iCode
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        Loop While ilFound
        ilCount = 0
        Do While ilFind <> -1
            For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
                If tgRtfList(ilLoop).tRtf.iCode = ilFind Then
                    ilCount = ilCount + 1
                    slStr = Trim$(str$(ilCount))
                    Do While Len(slStr) < 4
                        slStr = "0" & slStr
                    Loop
                    tgRtfList(ilLoop).sKey = slStr
                    ilFind = tgRtfList(ilLoop).tRtf.iNextRtfCode
                    Exit For
                End If
            Next ilLoop
        Loop
        If UBound(tgRtfList) - 1 > 0 Then
            ArraySortTyp fnAV(tgRtfList(), 0), UBound(tgRtfList), 0, LenB(tgRtfList(0)), 0, LenB(tgRtfList(0).sKey), 0
        End If
        'Eliminate Dormant Reports and categories
        ilDel = False
        ilLoop = LBound(tgRtfList)
        Do
            slState = tgRtfList(ilLoop).tRtf.sRnfState  '"D"
            slType = tgRtfList(ilLoop).tRtf.sRnfType    '""
            'For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            '    'slNameCode = tgRnfList(ilRnf).sKey    'lbcMster.List(ilLoop)
            '    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '    'If Val(slCode) = tgRtfList(ilLoop).tRtf.iRnfCode Then
            '    '    ilRet = gParseItem(slName, 1, "|", slType)
            '    '    ilRet = gParseItem(slNameCode, 3, "|", slState)
            '    '    Exit For
            '    'End If
            '    If tgRnfList(ilRnf).tRnf.iCode = tgRtfList(ilLoop).tRtf.iRnfCode Then
            '        slType = tgRnfList(ilRnf).tRnf.sType
            '        slState = tgRnfList(ilRnf).tRnf.sState
            '        Exit For
            '    End If
            'Next ilRnf
            If (ilDel) Or (slState = "D") Then
                If ilSaveDormant Then
                    tgDelRtfList(UBound(tgDelRtfList)) = tgRtfList(ilLoop)
                    ReDim Preserve tgDelRtfList(0 To UBound(tgDelRtfList) + 1) As RTFLIST
                End If
                For ilRtf = ilLoop To UBound(tgRtfList) - 1 Step 1
                    tgRtfList(ilRtf) = tgRtfList(ilRtf + 1)
                Next ilRtf
                ReDim Preserve tgRtfList(0 To UBound(tgRtfList) - 1) As RTFLIST
            Else
                ilLoop = ilLoop + 1
            End If
            If slType = "C" Then
                If slState = "D" Then
                    ilDel = True
                Else
                    ilDel = False
                End If
            End If
        Loop While ilLoop < UBound(tgRtfList)
    End If
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSNF                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain set name records        *
'*                                                     *
'*******************************************************
Sub gObtainSNF(hlSnf As Integer, ilIncludeDormant As Integer)
    Dim ilRecLen As Integer
    Dim ilRet As Integer
    Dim ilUpper As Integer
    ReDim tgSnfCode(0 To 0) As SNFCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilUpper = 0
    'ilExtLen = Len(tgSnfCode(0).tSnf)  'Extract operation record size
    ilRecLen = Len(tgSnfCode(0).tSnf) 'btrRecordLength(hlRnf)  'Get and save record length
    'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSnf) 'Obtain number of records
    'btrExtClear hlSnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlSnf, tgSnfCode(ilUpper).tSnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
    'Call btrExtSetBounds(hlSnf, llNoRec, -1, "UC") 'Set extract limits (all records)
    'ilOffset = 0
    'ilRet = btrExtAddField(hlSnf, ilOffset, ilExtLen)  'Extract First Name field
    'If ilRet <> BTRV_ERR_NONE Then
    '    Exit Sub
    'End If
    'ilExtLen = Len(tgSnfCode(ilUpper).tSnf)  'Extract operation record size
    'ilRet = btrExtGetNext(hlSnf, tgSnfCode(ilUpper).tSnf, ilExtLen, tgSnfCode(ilUpper).lRecPos)
    'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
    '    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
    '        Exit Sub
    '    End If
    '    'ilRet = btrExtGetFirst(hlRtf, tgCompMnf(ilUpper), ilExtLen, llRecPos)
    '    Do While ilRet = BTRV_ERR_REJECT_COUNT
    '        ilExtLen = Len(tgSnfCode(ilUpper).tSnf)  'Extract operation record size
    '        ilRet = btrExtGetNext(hlSnf, tgSnfCode(ilUpper).tSnf, ilExtLen, tgSnfCode(ilUpper).lRecPos)
    '    Loop
        Do While ilRet = BTRV_ERR_NONE
            If (ilIncludeDormant) Or (tgSnfCode(ilUpper).tSnf.sState = "A") Then
                tgSnfCode(ilUpper).sKey = tgSnfCode(ilUpper).tSnf.sName
                ilUpper = ilUpper + 1
                ReDim Preserve tgSnfCode(0 To ilUpper) As SNFCODE
       '         ilExtLen = Len(tgSnfCode(ilUpper).tSnf)  'Extract operation record size
       '         ilRet = btrExtGetNext(hlSnf, tgSnfCode(ilUpper).tSnf, ilExtLen, tgSnfCode(ilUpper).lRecPos)
       '         Do While ilRet = BTRV_ERR_REJECT_COUNT
       '             ilExtLen = Len(tgSnfCode(ilUpper).tSnf)  'Extract operation record size
       '             ilRet = btrExtGetNext(hlSnf, tgSnfCode(ilUpper).tSnf, ilExtLen, tgSnfCode(ilUpper).lRecPos)
       '         Loop
            End If
            ilRecLen = Len(tgSnfCode(0).tSnf) 'btrRecordLength(hlRnf)  'Get and save record length
            ilRet = btrGetNext(hlSnf, tgSnfCode(ilUpper).tSnf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
        Loop
        If UBound(tgSnfCode) - 1 > 0 Then
            ArraySortTyp fnAV(tgSnfCode(), 0), UBound(tgSnfCode), 0, LenB(tgSnfCode(0)), 0, LenB(tgSnfCode(0).sKey), 0
        End If
   ' End If

End Sub
