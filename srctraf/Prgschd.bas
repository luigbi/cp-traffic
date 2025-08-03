Attribute VB_Name = "PRGSCHD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Prgschd.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PrgSchd.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule subs and functions
Option Explicit
Option Compare Text
'Vehicle
Dim tmVef As VEF
Dim hmVef As Integer        'VEF handle
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0
Dim lmVlfCode() As Long
Type LIBADJINFO
    iLtfCode As Integer
    iMaxVersion As Integer
    iMaxExist As Integer
End Type
Dim tmLibAdjInfo() As LIBADJINFO
Dim lmAdjLvfCode() As Long



'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjLVFVersion                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Adjust library version so the   *
'*                     ones defined as TFN have the    *
'*                     latest version number           *
'*                     Since we want the latest version*
'*                     to show in Program screen       *
'*                                                     *
'*******************************************************
Function gAdjLVFVersion(hlLcf As Integer, hlLvf As Integer, ilType As Integer, sLCP As String, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer) As Integer
'
'   gAdjLVFVersion hlLcf, hlLvf, ilType, ilVefCode, ilLogDate0, ilLogDate1
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       hlLnf (I)- LVF Handle
'       ilType (I)- 0=Regular Programming; 1->NN = Sports Programming (Game Number)
'       slCP (I)- "C" = Current; "P" = Pending
'       ilVefCode (I)- Vehicle code
'       ilLogDate0 (I)- Log date to be be checked
'       ilLogDate1
'
                    'Remove current TFN so new one can be created
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlLvf As LVF                'LVF record image
    Dim tlLvfVersion As LVF                'LVF record image
    Dim tlLvfSrchKey0 As LONGKEY0     'LVF key record image
    Dim ilRet As Integer
    Dim ilLcfIndex As Integer
    Dim tlLvfSrchKey1 As LVFKEY1     'LVF key record image
    Dim ilLvfRecLen As Integer         'LVF record length
    Dim slDate As String
    Dim ilTFNDate0 As Integer
    Dim ilTFNDate1 As Integer
    Dim ilMaxExist As Integer
    Dim ilLoop As Integer

    'Check TFN LCF for specified date
    If (ilLogDate0 <= 7) And (ilLogDate1 = 0) Then
        ilTFNDate0 = ilLogDate0
        ilTFNDate1 = ilLogDate1
    Else
        gUnpackDate ilLogDate0, ilLogDate1, slDate
        ilTFNDate0 = gWeekDayStr(slDate) + 1
        ilTFNDate1 = 0
    End If
    ilLcfRecLen = Len(tlLcf)
    ilLvfRecLen = Len(tlLvf)
    tlLcfSrchKey.iType = ilType
    tlLcfSrchKey.sStatus = sLCP
    tlLcfSrchKey.iVefCode = ilVefCode
    tlLcfSrchKey.iLogDate(0) = ilTFNDate0
    tlLcfSrchKey.iLogDate(1) = ilTFNDate1
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) And (tlLcf.iLogDate(0) = ilTFNDate0) And (tlLcf.iLogDate(1) = ilTFNDate1)
        'Adjust library version numbers
        For ilLcfIndex = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
            If tlLcf.lLvfCode(ilLcfIndex) > 0 Then
                tlLvfSrchKey0.lCode = tlLcf.lLvfCode(ilLcfIndex)
                ilRet = btrGetEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = ilRet
                    sgErrLoc = "gAdjLVFVersion-Get Equal Lvf(1)"
                    gAdjLVFVersion = False
                    Exit Function
                End If
                ilMaxExist = False
                For ilLoop = 0 To UBound(tmLibAdjInfo) - 1 Step 1
                    If tmLibAdjInfo(ilLoop).iLtfCode = tlLvf.iLtfCode Then
                        ilMaxExist = tmLibAdjInfo(ilLoop).iMaxExist
                        Exit For
                    End If
                Next ilLoop
                If Not ilMaxExist Then
                    For ilLoop = 0 To UBound(lmAdjLvfCode) - 1 Step 1
                        If lmAdjLvfCode(ilLoop) = tlLvf.lCode Then
                            ilMaxExist = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If Not ilMaxExist Then
                    tlLvfSrchKey1.iLtfCode = tlLvf.iLtfCode
                    tlLvfSrchKey1.iVersion = 32000
                    ilRet = btrGetGreaterOrEqual(hlLvf, tlLvfVersion, ilLvfRecLen, tlLvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    If (ilRet = BTRV_ERR_NONE) And (tlLvfVersion.iLtfCode = tlLvf.iLtfCode) And (tlLvfVersion.iVersion > tlLvf.iVersion) Then
                        Do
                            tlLvfSrchKey0.lCode = tlLcf.lLvfCode(ilLcfIndex)
                            ilRet = btrGetEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get current record
                            If ilRet <> BTRV_ERR_NONE Then
                                igBtrError = ilRet
                                sgErrLoc = "gAdjLVFVersion-Get Equal Lvf(2)"
                                gAdjLVFVersion = False
                                Exit Function
                            End If
                            tlLvf.iVersion = tlLvfVersion.iVersion + 1
                            ilRet = btrUpdate(hlLvf, tlLvf, ilLvfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = ilRet
                            sgErrLoc = "gAdjLVFVersion-Update Lvf(3)"
                            gAdjLVFVersion = False
                            Exit Function
                        End If
                        lmAdjLvfCode(UBound(lmAdjLvfCode)) = tlLvf.lCode
                        ReDim Preserve lmAdjLvfCode(0 To UBound(lmAdjLvfCode) + 1) As Long
                    End If
                End If
            Else
                Exit For
            End If
        Next ilLcfIndex
        ilRet = btrGetNext(hlLcf, tlLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    Loop
    gAdjLVFVersion = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildLCF                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Lcf for specified         *
'*                     vehicles                        *
'*                                                     *
'*******************************************************
Function gBuildLCF(tlVefSch() As VEFSCH, lbcNotSchd As control) As Integer
    Dim hlLcf As Integer            'Log calendar library file handle
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim hlLvf As Integer            'Log version library file handle
    Dim hlSsf As Integer            'Spot summary file handle
    Dim hlSdf As Integer            'Spot detail file handle
    Dim hlSmf As Integer            'Spot MG file handle
    Dim ilRet As Integer
    Dim llCEarliestDate As Long   'Earliest date of current libraries
    Dim llPEarliestDate As Long   'Pending Earliest date libraries
    Dim llDEarliestDate As Long   'Pending Delete Earliest date libraries
    Dim llPDEarliestDate As Long   'Pending/Delete Earliest date libraries
    Dim llCLatestDate As Long   'Latest date of current libraries
    Dim llPLatestDate As Long   'Pending latest date libraries
    Dim llDLatestDate As Long   'Pending Delete latest date libraries
    Dim llPDLatestDate As Long   'Pending/Delete latest date libraries
    Dim llSundayDate As Long
    Dim llMondayDate As Long
    Dim ilCTFNExist As Integer  'True=Current TFN exist
    Dim ilPTFNExist As Integer  'True = Pending TFN exist
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilPending As Integer
    Dim ilTFN As Integer
    Dim slMsg As String
    Dim ilVef As Integer
    Dim slType As String
    ReDim ilVersionUpdated(0 To 6) As Integer 'True=Version updated
    ReDim ilEvtType(0 To 14) As Integer
    ReDim tlLLC(0 To 0) As LLC  'Merged library names
    ReDim tlPLLC(0 To 0) As LLC
    If Not gOpenSchFiles() Then
        gBuildLCF = False
        Exit Function
    End If
    hlLcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Lcf(1), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        btrDestroy hlLcf
        gCloseSchFiles
        Exit Function
    End If
    ilLcfRecLen = Len(tlLcf)
    hlLvf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Lvf(2), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLvf)
        btrDestroy hlLcf
        btrDestroy hlLvf
        gCloseSchFiles
        Exit Function
    End If
    hlSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Ssf(1), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLvf)
        ilRet = btrClose(hlSsf)
        btrDestroy hlLcf
        btrDestroy hlLvf
        btrDestroy hlSsf
        gCloseSchFiles
        Exit Function
    End If
    hlSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Sdf(1), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLvf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        btrDestroy hlLcf
        btrDestroy hlLvf
        btrDestroy hlSsf
        btrDestroy hlSdf
        gCloseSchFiles
        Exit Function
    End If
    hlSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Smf(1), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLvf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        btrDestroy hlLcf
        btrDestroy hlLvf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        gCloseSchFiles
        Exit Function
    End If
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildLcf- Open Vef(1), I/O Error" & str(ilRet)
        gBuildLCF = False
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLvf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hmVef)
        btrDestroy hlLcf
        btrDestroy hlLvf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hmVef
        gCloseSchFiles
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(0) = True
    For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
        'Exclude Sports as done in GameSchd
        ilVef = gBinarySearchVef(tlVefSch(ilLoop).iVefCode)
        If ilVef <> -1 Then
            slType = tgMVef(ilVef).sType
        Else
            slType = "G"
        End If
        If (tlVefSch(ilLoop).sOnAirSchStatus = "S") And (slType <> "G") Then
            ReDim lmAdjLvfCode(0 To 0) As Long
            mSetAdjMax hlLcf, hlLvf, "C", tlVefSch(ilLoop).iVefCode, ilEvtType()
            For ilIndex = 0 To 6 Step 1
                ilVersionUpdated(ilIndex) = False
            Next ilIndex
            llCEarliestDate = gGetEarliestLCFDate(hlLcf, "C", tlVefSch(ilLoop).iVefCode)
            llPEarliestDate = gGetEarliestLCFDate(hlLcf, "P", tlVefSch(ilLoop).iVefCode)
            llDEarliestDate = gGetEarliestLCFDate(hlLcf, "D", tlVefSch(ilLoop).iVefCode)
            If llPEarliestDate > 0 Then
                If (llDEarliestDate > 0) And (llDEarliestDate < llPEarliestDate) Then
                    llPDEarliestDate = llDEarliestDate
                Else
                    llPDEarliestDate = llPEarliestDate
                End If
            Else
                llPDEarliestDate = llDEarliestDate
            End If
            If (llPDEarliestDate > 0) Then 'Check if any pending- if not byapss (Vlf change only)
                ilRet = btrBeginTrans(hlLcf, 1000)
                If ilRet <> BTRV_ERR_NONE Then
                    slMsg = "I/O Error:" & str$(ilRet) & " Begin Transactions"
                    '6/5/16: Replaced GoSub
                    'GoSub BuildLCFErrorMsg
                    mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                    gBuildLCF = False
                    Exit Function
                End If
                llCLatestDate = gGetLatestLCFDate(hlLcf, "C", tlVefSch(ilLoop).iVefCode)
                llPLatestDate = gGetLatestLCFDate(hlLcf, "P", tlVefSch(ilLoop).iVefCode)
                llDLatestDate = gGetLatestLCFDate(hlLcf, "D", tlVefSch(ilLoop).iVefCode)
                If llPLatestDate > 0 Then
                    If (llDLatestDate > 0) And (llDLatestDate > llPLatestDate) Then
                        llPDLatestDate = llDLatestDate
                    Else
                        llPDLatestDate = llPLatestDate
                    End If
                Else
                    llPDLatestDate = llDLatestDate
                End If
                '
                'ABC Air vehicle Code to ADD to 2.0304- start point
                '
                'Extend Current to Start of Pending (ignore delete)
                'Fill in grip between current and start of pending before TFN changed
                '  Current   |---------|
                '  Pending                  |----|
                '  Fill-In              |--| with TFN from current prior to altering TFN
                '
                If llCLatestDate <> -1 Then
                    If llPEarliestDate <> -1 Then
                        'Later, the days will be extend to the Latest Pending Date
                        slDate = Format$(llPEarliestDate, "m/d/yy")
                        slDate = gObtainPrevSunday(slDate)
                        llSundayDate = gDateValue(slDate)
                        For llDate = llCLatestDate + 1 To llSundayDate Step 1
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilLogDate0, ilLogDate1
                            ilRet = gExtendTFN(hlLcf, hlSsf, hlSdf, hlSmf, "C", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1, True)
                            If Not ilRet Then
                                slMsg = "Extend TFN Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                        Next llDate
                    End If
                End If
                '
                'ABC Air vehicle Code to add to 2.0304- end point
                '
                ilCTFNExist = mTFNExist(hlLcf, 0, "C", tlVefSch(ilLoop).iVefCode)
                ilPTFNExist = mTFNExist(hlLcf, 0, "P", tlVefSch(ilLoop).iVefCode)
                If Not ilPTFNExist Then
                    ilPTFNExist = mTFNExist(hlLcf, 0, "D", tlVefSch(ilLoop).iVefCode)
                End If
                If ilPTFNExist Then
                    'Merge TFN pending into Current, then replace current with merged current and pending, then correct library version number
                    '(TFN libraries are to contain the latest version number)
                    For ilTFN = 1 To 7 Step 1
                        ReDim tlLLC(0 To 0) As LLC  'Merged library names
                        Select Case ilTFN
                            Case 1
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                            Case 2
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                            Case 3
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                            Case 4
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                            Case 5
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                            Case 6
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                            Case 7
                                ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                        End Select
                        If Not ilRet Then
                            slMsg = "Build Event Day Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
                        'Remove current TFN so new one can be created
                        ilRet = gRemoveLCFDate(hlLcf, "C", tlVefSch(ilLoop).iVefCode, ilTFN, 0)
                        If Not ilRet Then
                            slMsg = "Remove LCF Date Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
                        'Make LCF from tlLLC
                        ilRet = mMakeLCFFromLLC(hlLcf, 0, tlVefSch(ilLoop).iVefCode, ilTFN, 0, tlLLC())
                        If Not ilRet Then
                            slMsg = "Make LCF from LLC Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
                        'Adjust library version numbers (use current instead of pending to reduce doing same TFN twice)
                        ilRet = gAdjLVFVersion(hlLcf, hlLvf, 0, "C", tlVefSch(ilLoop).iVefCode, ilTFN, 0)
                        If Not ilRet Then
                            slMsg = "Adjust Library Version Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
                        ilVersionUpdated(ilTFN - 1) = True
                        'Remove LCF for pending TFN
                        ilRet = gRemoveLCFDate(hlLcf, "D", tlVefSch(ilLoop).iVefCode, ilTFN, 0)
                        If Not ilRet Then
                            slMsg = "Remove LCF Date Error"
                        '6/5/16: Replaced GoSub
                        'GoSub BuildLCFErrorMsg
                        mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                        gBuildLCF = False
                            Exit Function
                        End If
                        ilRet = gRemoveLCFDate(hlLcf, "P", tlVefSch(ilLoop).iVefCode, ilTFN, 0)
                        If Not ilRet Then
                            slMsg = "Remove LCF Date Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
                    Next ilTFN
                End If
                'If pending end date greater then current, extend current to match pending

                'Fill in grip between current and start of pending before TFN changed
                '  Current   |-------------|
                '  Pending                  |----|
                '  Fill-In                  |----| with TFN from TFN made up of Current and Pending
                '
                '12/29/05- If current did not have TFN, this code was extending it from its end to end of new library
                'If llCLatestDate <> -1 Then
                If (llCLatestDate <> -1) And (ilCTFNExist) Then
                    If llPDLatestDate <> -1 Then
                        'Any dates created previously will be bypassed in gExtendTFN
                        slDate = Format$(llPDLatestDate, "m/d/yy")
                        slDate = gObtainNextSunday(slDate)
                        llSundayDate = gDateValue(slDate)
                        For llDate = llCLatestDate + 1 To llSundayDate Step 1
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilLogDate0, ilLogDate1
                            ilRet = gExtendTFN(hlLcf, hlSsf, hlSdf, hlSmf, "C", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1, True)
                            If Not ilRet Then
                                slMsg = "Extend TFN Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                        Next llDate
                    End If
                Else
                    If llPDLatestDate <> -1 Then
                        slDate = Format$(llPDEarliestDate, "m/d/yy")
                        slDate = gObtainPrevMonday(slDate)
                        llMondayDate = gDateValue(slDate)
                        slDate = Format$(llPDLatestDate, "m/d/yy")
                        slDate = gObtainNextSunday(slDate)
                        llSundayDate = gDateValue(slDate)
                        For llDate = llMondayDate To llSundayDate Step 1
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilLogDate0, ilLogDate1
                            ilRet = gExtendTFN(hlLcf, hlSsf, hlSdf, hlSmf, "C", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1, True)
                            If Not ilRet Then
                                slMsg = "Extend TFN Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                        Next llDate
                    End If
                End If
                If llPDLatestDate <> -1 Then
                    slDate = Format$(llPDLatestDate, "m/d/yy")
                    slDate = gObtainNextSunday(slDate)
                    llSundayDate = gDateValue(slDate)
                    'ReDim lgReschSdfCode(1 To 1) As Long
                    ReDim lgReschSdfCode(0 To 0) As Long
                    For llDate = llPDEarliestDate To llSundayDate Step 1
                        ReDim tlLLC(0 To 0) As LLC  'Merged library names
    '                    ReDim tlCLLC(0 To 0) As LLC
                        ReDim tlPLLC(0 To 0) As LLC
                        For ilIndex = LBound(ilEvtType) To UBound(ilEvtType) Step 1
                            ilEvtType(ilIndex) = False
                        Next ilIndex
                        ilEvtType(0) = True
                        slDate = Format$(llDate, "m/d/yy")
                        gPackDate slDate, ilLogDate0, ilLogDate1
                        'Build only library names
                        ilRet = gBuildEventDay(0, "B", tlVefSch(ilLoop).iVefCode, slDate, "12M", "12M", ilEvtType(), tlLLC())
                        If Not ilRet Then
                            slMsg = "Build Event Day Error"
                            '6/5/16: Replaced GoSub
                            'GoSub BuildLCFErrorMsg
                            mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                            gBuildLCF = False
                            Exit Function
                        End If
    '                    ilRet = gBuildEventDay("O", "C", tlVefSch(ilLoop).iVefCode, slDate, "12M", "12M", ilEvtType(), tlCLLC())
                        'Build libraries and avails
    '                    For ilIndex = 2 To 9 Step 1 'All avails
    '                        ilEvtType(ilIndex) = True
    '                    Next ilIndex
                        'Determine if day has any pending to be scheduled
                        ilPending = False
                        'ilRet = gBuildEventDay("O", "P", tlVefSch(ilLoop).iVefCode, slDate, "12M", "12M", ilEvtType(), tlPLLC())
                        'Test if any Pending
                        tlLcfSrchKey.iType = 0
                        tlLcfSrchKey.sStatus = "P"
                        tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
                        tlLcfSrchKey.iLogDate(0) = ilLogDate0
                        tlLcfSrchKey.iLogDate(1) = ilLogDate1
                        tlLcfSrchKey.iSeqNo = 1
                        ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                        If ilRet = BTRV_ERR_NONE Then
                            ilPending = True
                        End If
                        If Not ilPending Then
                            'Test if any deleted
                            tlLcfSrchKey.iType = 0
                            tlLcfSrchKey.sStatus = "D"
                            tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
                            tlLcfSrchKey.iLogDate(0) = ilLogDate0
                            tlLcfSrchKey.iLogDate(1) = ilLogDate1
                            tlLcfSrchKey.iSeqNo = 1
                            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilRet = BTRV_ERR_NONE Then
                                ilPending = True
                            End If
                        End If
                        If ilPending Then
                            'Remove current Date so new one can be created
                            ilRet = gRemoveLCFDate(hlLcf, "C", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1)
                            If Not ilRet Then
                                slMsg = "Remove LCF Date Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                            'Make LCF from tlLLC
                            ilRet = mMakeLCFFromLLC(hlLcf, 0, tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1, tlLLC())
                            If Not ilRet Then
                                slMsg = "Make LCF from LLC Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                            If Not ilVersionUpdated(gWeekDayStr(slDate)) Then
                                'Adjust library version numbers
                                ilRet = gAdjLVFVersion(hlLcf, hlLvf, 0, "C", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1)
                                If Not ilRet Then
                                    slMsg = "Adjust Library Version Error"
                                    '6/5/16: Replaced GoSub
                                    'GoSub BuildLCFErrorMsg
                                    mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                    gBuildLCF = False
                                    Exit Function
                                End If
                                ilVersionUpdated(gWeekDayStr(slDate)) = True
                            End If
                            'Remove LCF for pending
                            ilRet = gRemoveLCFDate(hlLcf, "D", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1)
                            If Not ilRet Then
                                slMsg = "Remove LCF Date Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                            ilRet = gRemoveLCFDate(hlLcf, "P", tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1)
                            If Not ilRet Then
                                slMsg = "Remove LCF Date Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                            'Remove spots that are not within an avail or overbooked
                            '   Compare tlPLLC (pending lib and avails) to tlCLLC (old current)
                            '   The pending day has been moved into lcf and made current.
                            '
                            '           Time->
                            '   tlCLLC  I--------II------------II-------II-------------II-----------I
                            '   tlPLLC                      I----------------I
                            '           I Retain II Remove II    Test spots  II Remove II   Retain  I
                            '              spots     spots                       spots       spots
                            '
                            '           Shown are library times and what to do with spots within the
                            '           library limits
                            'Make avail summary file- Merge avails for day with spots for day
                            '
    '                        ReDim tlLLC(0 To 0) As LLC  'Merged library names
    '                        For ilIndex = LBound(ilEvtType) To UBound(ilEvtType) Step 1
    '                            ilEvtType(ilIndex) = False
    '                        Next ilIndex
    ''                        ilEvtType(0) = True
    '                        ilEvtType(1) = True 'Program
    '                        For ilIndex = 2 To 9 Step 1 'All avails
    '                            ilEvtType(ilIndex) = True
    '                        Next ilIndex
    '                        ilRet = gBuildEventDay("O", "C", tlVefSch(ilLoop).iVefCode, slDate, "12M", "12M", ilEvtType(), tlLLC())
                            ilRet = gMakeSSF(False, hlSsf, hlSdf, hlSmf, 0, tlVefSch(ilLoop).iVefCode, ilLogDate0, ilLogDate1, 0)
                            If Not ilRet Then
                                slMsg = "Make SSF Error"
                                '6/5/16: Replaced GoSub
                                'GoSub BuildLCFErrorMsg
                                mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                                gBuildLCF = False
                                Exit Function
                            End If
                        End If
                    Next llDate
                    ilRet = btrEndTrans(hlLcf)
                    Erase sgSSFErrorMsg
                    ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
                    If Not ilRet Then
                        slMsg = "Reschedule Spot Error"
                        '6/5/16: Replaced GoSub
                        'GoSub BuildLCFErrorMsg
                        mBuildLCFErrorMsg hlLcf, hlLvf, hlSsf, hlSdf, hlSmf, tlVefSch(), ilLoop, slMsg, lbcNotSchd
                        gBuildLCF = False
                        Exit Function
                    End If
                Else
                    ilRet = btrEndTrans(hlLcf)
                End If
            End If
        End If
    Next ilLoop
    Erase lmAdjLvfCode
    Erase tmLibAdjInfo
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    ilRet = btrClose(hlSsf)
    btrDestroy hlSsf
    ilRet = btrClose(hlSdf)
    btrDestroy hlSdf
    ilRet = btrClose(hlSmf)
    btrDestroy hlSmf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    gCloseSchFiles
    gBuildLCF = True
    Exit Function
'BuildLCFErrorMsg:
'    ilRet = btrAbortTrans(hlLcf)
'    tmVefSrchKey.iCode = tlVefSch(ilLoop).iVefCode
'    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If ilRet <> BTRV_ERR_NONE Then
'        tmVef.sName = "Vehicle Name Missing"
'    End If
'    If igBtrError > 0 Then
'        lbcNotSchd.AddItem "Error #" & str(igBtrError) & "/" & sgErrLoc & " for " & Trim$(tmVef.sName) & ": " & slMsg
'    Else
'        lbcNotSchd.AddItem Trim$(tmVef.sName) & ": " & slMsg
'    End If
'    btrDestroy hlLcf
'    btrDestroy hlLvf
'    btrDestroy hlSsf
'    btrDestroy hlSdf
'    btrDestroy hlSmf
'    btrDestroy hmVef
'    Erase sgSSFErrorMsg
'    gBuildLCF = False
'    gCloseSchFiles
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildVCF                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Vcf for specified         *
'*                     vehicles                        *
'*                                                     *
'*******************************************************
Function gBuildVCF(tlVefSch() As VEFSCH, lbcNotSchd As control) As Integer
'
'   ilRet = gBuildVCF( tlVefSch)
'
'   Where:
'       tlVefSch (I) - Type structure containing schedule info for each vehicle
'                      sOnAirSchStatus- On Air Schedule status- "S"=Schedule this vehicle (links and/or library); ""=No scheduling required
'                      sAltSchStatus- Alternate Schedule status- "S"=Schedule this vehicle (links and/or library); ""=No scheduling required
'                      iDay- used by selling and airing only (0=M-F; 6=Sa; 7=Su)
'                      iVefCode- Vehicle code
'                      sType- Vehicle type:"C"-Conventional; "S"-Selling; "A"-Airing; "D"-Delivery
'                      iGroup- Links selling and airing (same group number)
'       ilRet (O)- True if VCF successfully made
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilLink As Integer
    Dim ilMatchFd As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim hlVcf As Integer        'Vehicle file handle
    Dim tlVcf As VCF
    Dim tlVcfSrchKey0 As VCFKEY0  'Vcf key record image
    Dim ilVcfRecLen As Integer     'VEF record length
    Dim llVcfRecPos As Long
    Dim hlVlf As Integer        'Vehicle link file handle
    Dim tlVlf As VLF
    Dim tlAirVlf As VLF
    Dim tlVlfSrchKey0 As VLFKEY0  'Vlf key record image
    Dim tlVlfSrchKey1 As VLFKEY1  'Vlf key record image
    Dim tlVlfSrchKey2 As LONGKEY0
    Dim ilVlfRecLen As Integer     'VLF record length
    Dim llVlfCode As Long
    Dim ilNoXLinks As Integer
    Dim ilVlf As Integer
    Dim ilDay As Integer
    Dim ilTermDate0 As Integer
    Dim ilTermDate1 As Integer
    Dim llAirTime As Long
    Dim llAirTimeP60 As Long
    Dim llAirTimeN60 As Long
    Dim llStartDate As Long     'Terminate date plus one
    Dim llDate As Long
    Dim slMsg As String
    Dim ilVff As Integer
    hlVcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVcf, "", sgDBPath & "Vcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildVcf- Open Vcf(1), I/O Error" & str(ilRet)
        gBuildVCF = False
        ilRet = btrClose(hlVcf)
        btrDestroy hlVcf
        Exit Function
    End If
    ilVcfRecLen = Len(tlVcf)  'btrRecordLength(hlVcf)  'Get and save record length
    hlVlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVlf, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        lbcNotSchd.AddItem "gBuildVcf- Open Vlf(2), I/O Error" & str(ilRet)
        gBuildVCF = False
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlVlf)
        btrDestroy hlVcf
        btrDestroy hlVlf
        Exit Function
    End If
    ilRet = btrBeginTrans(hlVcf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        slMsg = "I/O Error:" & str$(ilRet) & " Begin Transactions"
        '6/6/16: Replaced GoSub
        'GoSub BuildVCFErrorMsg
        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
        gBuildVCF = False
        Exit Function
    End If
    'Terminate all VCF (termiante date set within gBuildVehSchInfo, if 0 ignore day)
    For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
        If (tlVefSch(ilLoop).sOnAirSchStatus = "S") And (tlVefSch(ilLoop).sType = "S") Then
            For ilDay = 0 To 2 Step 1
                tlVcfSrchKey0.iSellCode = tlVefSch(ilLoop).iVefCode
                Select Case ilDay
                    Case 0
                        tlVefSch(ilLoop).iDay = 0
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate0(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate0(1)
                    Case 1
                        tlVefSch(ilLoop).iDay = 6
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate6(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate6(1)
                    Case 2
                        tlVefSch(ilLoop).iDay = 7
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate7(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate7(1)
                End Select
                If (ilTermDate0 <> 0) Or (ilTermDate1 <> 0) Then
                    tlVcfSrchKey0.iSellDay = tlVefSch(ilLoop).iDay
                    tlVcfSrchKey0.iEffDate(0) = 0
                    tlVcfSrchKey0.iEffDate(1) = 0
                    tlVcfSrchKey0.iSellTime(0) = 0
                    tlVcfSrchKey0.iSellTime(1) = 0
                    tlVcfSrchKey0.iSellPosNo = 0
                    ilRet = btrGetGreaterOrEqual(hlVcf, tlVcf, ilVcfRecLen, tlVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                    Do While (ilRet = BTRV_ERR_NONE) And (tlVcf.iSellCode = tlVefSch(ilLoop).iVefCode) And (tlVcf.iSellDay = tlVefSch(ilLoop).iDay)
                        If (tlVcf.iTermDate(0) = 0) And (tlVcf.iTermDate(1) = 0) Then
                            gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                            llStartDate = llStartDate + 1
                            gUnpackDateLong tlVcf.iEffDate(0), tlVcf.iEffDate(1), llDate
                            If llStartDate = llDate Then
                                'Terminate before start- remove
                                'The GetNext will still work even when record is deleted
                                ilRet = btrGetPosition(hlVcf, llVcfRecPos)
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Get Position VCF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                                Do
                                    'tmSRec = tlVcf
                                    'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
                                    'tlVcf = tmSRec
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
                                    '    GoSub BuildVCFErrorMsg
                                    '    Exit Function
                                    'End If
                                    ilRet = btrDelete(hlVcf)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        If ilCRet <> BTRV_ERR_NONE Then
                                            slMsg = "I/O Error:" & str$(ilCRet) & " Get Direct VCF"
                                            '6/6/16: Replaced GoSub
                                            'GoSub BuildVCFErrorMsg
                                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                            gBuildVCF = False
                                            Exit Function
                                        End If
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Delete VCF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                            Else
                                If (ilTermDate1 < tlVcf.iEffDate(1)) Or ((ilTermDate1 = tlVcf.iEffDate(1)) And (ilTermDate0 < tlVcf.iEffDate(0))) Then
    '                                'Terminate before start- remove
    '                                'The GetNext will still work even when record is deleted
    '                                ilRet = btrGetPosition(hlVcf, llVcfRecPos)
    '                                If ilRet <> BTRV_ERR_NONE Then
    '                                    slMsg = "I/O Error:" & Str$(ilRet) & " Get Position VCF"
    '                                    GoSub BuildVCFErrorMsg
    '                                    Exit Function
    '                                End If
    '                                Do
    '                                    'tmSRec = tlVcf
    '                                    'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
    '                                    'tlVcf = tmSRec
    '                                    'If ilRet <> BTRV_ERR_NONE Then
    '                                    '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
    '                                    '    GoSub BuildVCFErrorMsg
    '                                    '    Exit Function
    '                                    'End If
    '                                    ilRet = btrDelete(hlVcf)
    '                                    If ilRet = BTRV_ERR_CONFLICT Then
    '                                        ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    '                                        If ilCRet <> BTRV_ERR_NONE Then
    '                                            slMsg = "I/O Error:" & Str$(ilCRet) & " Get Direct VCF"
    '                                            GoSub BuildVCFErrorMsg
    '                                            Exit Function
    '                                        End If
    '                                    End If
    '                                Loop While ilRet = BTRV_ERR_CONFLICT
    '                                If ilRet <> BTRV_ERR_NONE Then
    '                                    slMsg = "I/O Error:" & Str$(ilRet) & " Delete VCF"
    '                                    GoSub BuildVCFErrorMsg
    '                                    Exit Function
    '                                End If
                                Else
                                    ilRet = btrGetPosition(hlVcf, llVcfRecPos)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Get Position VCF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                    Do
                                        'tmSRec = tlVcf
                                        'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
                                        'tlVcf = tmSRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
                                        '    GoSub BuildVCFErrorMsg
                                        '    Exit Function
                                        'End If
                                        tlVcf.iTermDate(0) = ilTermDate0
                                        tlVcf.iTermDate(1) = ilTermDate1
                                        ilRet = btrUpdate(hlVcf, tlVcf, ilVcfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            If ilCRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilCRet) & " Get Direct VCF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Update VCF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else    'remove any previously terminated before start
                            If (tlVcf.iTermDate(1) < tlVcf.iEffDate(1)) Or ((tlVcf.iTermDate(1) = tlVcf.iEffDate(1)) And (tlVcf.iTermDate(0) < tlVcf.iEffDate(0))) Then
                                'Terminate before start- remove
                                'The GetNext will still work even when record is deleted
                                ilRet = btrGetPosition(hlVcf, llVcfRecPos)
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Get Position VCF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                                Do
                                    'tmSRec = tlVcf
                                    'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
                                    'tlVcf = tmSRec
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
                                    '    GoSub BuildVCFErrorMsg
                                    '    Exit Function
                                    'End If
                                    ilRet = btrDelete(hlVcf)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        If ilCRet <> BTRV_ERR_NONE Then
                                            slMsg = "I/O Error:" & str$(ilCRet) & " Get Direct VCF"
                                            '6/6/16: Replaced GoSub
                                            'GoSub BuildVCFErrorMsg
                                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                            gBuildVCF = False
                                            Exit Function
                                        End If
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Delete VCF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                            Else
                                'Add code 5/25
                                gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                                llStartDate = llStartDate + 1
                                gUnpackDateLong tlVcf.iEffDate(0), tlVcf.iEffDate(1), llDate
                                If llStartDate = llDate Then
                                    'Delete as it overlaps
                                    ilRet = btrGetPosition(hlVcf, llVcfRecPos)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Get Position VCF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                    Do
                                        'tmSRec = tlVcf
                                        'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
                                        'tlVcf = tmSRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
                                        '    GoSub BuildVCFErrorMsg
                                        '    Exit Function
                                        'End If
                                        ilRet = btrDelete(hlVcf)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            If ilCRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilCRet) & " Get Direct VCF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Delete VCF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                Else
                                    If (ilTermDate1 > tlVcf.iEffDate(1)) Or ((ilTermDate1 = tlVcf.iEffDate(1)) And (ilTermDate0 > tlVcf.iEffDate(0))) Then
                                        If (ilTermDate1 < tlVcf.iTermDate(1)) Or ((ilTermDate1 = tlVcf.iTermDate(1)) And (ilTermDate0 <= tlVcf.iTermDate(0))) Then
                                            ilRet = btrGetPosition(hlVcf, llVcfRecPos)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilRet) & " Get Position VCF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                            Do
                                                'tmSRec = tlVcf
                                                'ilRet = gGetByKeyForUpdate("VCF", hlVcf, tmSRec)
                                                'tlVcf = tmSRec
                                                'If ilRet <> BTRV_ERR_NONE Then
                                                '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VCF"
                                                '    GoSub BuildVCFErrorMsg
                                                '    Exit Function
                                                'End If
                                                tlVcf.iTermDate(0) = ilTermDate0
                                                tlVcf.iTermDate(1) = ilTermDate1
                                                ilRet = btrUpdate(hlVcf, tlVcf, ilVcfRecLen)
                                                If ilRet = BTRV_ERR_CONFLICT Then
                                                    ilCRet = btrGetDirect(hlVcf, tlVcf, ilVcfRecLen, llVcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    If ilCRet <> BTRV_ERR_NONE Then
                                                        slMsg = "I/O Error:" & str$(ilCRet) & " Get Direct VCF"
                                                        '6/6/16: Replaced GoSub
                                                        'GoSub BuildVCFErrorMsg
                                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                        gBuildVCF = False
                                                        Exit Function
                                                    End If
                                                End If
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilRet) & " Update VCF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        ilRet = btrGetNext(hlVcf, tlVcf, ilVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Loop
                End If
            Next ilDay
        End If
    Next ilLoop
    'Build VCF from VLF
    ilNoXLinks = 0
    For ilIndex = LBound(tlVcf.iCSV) To UBound(tlVcf.iCSV) Step 1
        tlVcf.iCSV(ilIndex) = 0
        tlVcf.sCSD(ilIndex) = " "
        tlVcf.iCST(0, ilIndex) = 0
        tlVcf.iCST(1, ilIndex) = 0
        tlVcf.iCSP(ilIndex) = 0
    Next ilIndex
    ilVlfRecLen = Len(tlVlf)  'btrRecordLength(hlVef)  'Get and save record length
    For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
        If (tlVefSch(ilLoop).sOnAirSchStatus = "S") And (tlVefSch(ilLoop).sType = "S") Then
            'Read in all VLF for vehicle
            For ilDay = 0 To 2 Step 1
                tlVlfSrchKey0.iSellCode = tlVefSch(ilLoop).iVefCode
                Select Case ilDay
                    Case 0
                        tlVefSch(ilLoop).iDay = 0
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate0(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate0(1)
                    Case 1
                        tlVefSch(ilLoop).iDay = 6
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate6(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate6(1)
                    Case 2
                        tlVefSch(ilLoop).iDay = 7
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate7(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate7(1)
                End Select
                If (ilTermDate0 <> 0) Or (ilTermDate1 <> 0) Then
                    tlVlfSrchKey0.iSellDay = tlVefSch(ilLoop).iDay
                    tlVlfSrchKey0.iEffDate(0) = 0
                    tlVlfSrchKey0.iEffDate(1) = 0
                    tlVlfSrchKey0.iSellTime(0) = 0
                    tlVlfSrchKey0.iSellTime(1) = 0
                    tlVlfSrchKey0.iSellPosNo = 0
                    tlVlfSrchKey0.iSellSeq = 0
                    ilRet = btrGetGreaterOrEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                    Do While (ilRet = BTRV_ERR_NONE) And (tlVlf.iSellCode = tlVefSch(ilLoop).iVefCode) And (tlVlf.iSellDay = tlVefSch(ilLoop).iDay)
                        If tlVlf.sStatus = "P" Then
                            '12/7/17
                            'ilRet = btrGetPosition(hlVlf, llVlfRecPos)
                            llVlfCode = tlVlf.lCode
                            If ilNoXLinks > 0 Then
                                If ((tlVcf.iSellTime(0) <> tlVlf.iSellTime(0)) Or (tlVcf.iSellTime(1) <> tlVlf.iSellTime(1))) Or ((tlVcf.iEffDate(0) <> tlVlf.iEffDate(0)) Or (tlVcf.iEffDate(1) <> tlVlf.iEffDate(1))) Then
                                    'Check for duplicate Vcf
                                    'Insert record
                                    If mVcfInsertTest(hlVcf, tlVcf) Then
                                        tlVcf.lCode = 0
                                        ilRet = btrInsert(hlVcf, tlVcf, ilVcfRecLen, INDEXKEY1)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            slMsg = "I/O Error:" & str$(ilRet) & " Insert VCF"
                                            '6/6/16: Replaced GoSub
                                            'GoSub BuildVCFErrorMsg
                                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                            gBuildVCF = False
                                            Exit Function
                                        End If
                                    End If
                                    ilNoXLinks = 0
                                    For ilIndex = LBound(tlVcf.iCSV) To UBound(tlVcf.iCSV) Step 1
                                        tlVcf.iCSV(ilIndex) = 0
                                        tlVcf.sCSD(ilIndex) = " "
                                        tlVcf.iCST(0, ilIndex) = 0
                                        tlVcf.iCST(1, ilIndex) = 0
                                        tlVcf.iCSP(ilIndex) = 0
                                    Next ilIndex
                                End If
                            End If
                            'Any avail within 60s of the airtime will be considered the same avail
                            gUnpackTimeLong tlVlf.iAirTime(0), tlVlf.iAirTime(1), True, llAirTime
                            llAirTimeN60 = llAirTime - 60
                            llAirTimeP60 = llAirTime + 60
                            If llAirTimeN60 < 0 Then
                                llAirTimeN60 = 0
                            End If
                            '12/27/17: Check if Ghost Window Length defined.  If so, use it
                            ilVff = gBinarySearchVff(tlVlf.iAirCode)
                            If ilVff <> -1 Then
                                If tgVff(ilVff).iConflictWinLen > 0 Then
                                    llAirTimeN60 = llAirTime - tgVff(ilVff).iConflictWinLen
                                    llAirTimeP60 = llAirTime + tgVff(ilVff).iConflictWinLen
                                    If llAirTimeN60 < 0 Then
                                        llAirTimeN60 = 0
                                    End If
                                End If
                            End If

                            tlVlfSrchKey1.iAirCode = tlVlf.iAirCode 'tlVefSch(ilIndex).iVefCode
                            tlVlfSrchKey1.iAirDay = tlVlf.iAirDay
                            tlVlfSrchKey1.iEffDate(0) = tlVlf.iEffDate(0)
                            tlVlfSrchKey1.iEffDate(1) = tlVlf.iEffDate(1)
                            gPackTimeLong llAirTimeN60, tlVlfSrchKey1.iAirTime(0), tlVlfSrchKey1.iAirTime(1)
                            'tlVlfSrchKey1.iAirTime(0) = tlVlf.iAirTime(0)
                            'tlVlfSrchKey1.iAirTime(1) = tlVlf.iAirTime(1)
                            tlVlfSrchKey1.iAirPosNo = tlVlf.iAirPosNo
                            tlVlfSrchKey1.iAirSeq = 0
                            ilRet = btrGetGreaterOrEqual(hlVlf, tlAirVlf, ilVlfRecLen, tlVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                            Do While (ilRet = BTRV_ERR_NONE) And (tlAirVlf.iAirCode = tlVlf.iAirCode) And (tlAirVlf.iAirDay = tlVlf.iAirDay) And (tlAirVlf.iAirPosNo = tlVlf.iAirPosNo)
                                'no room in the while above
                                If (tlAirVlf.iEffDate(0) <> tlVlf.iEffDate(0)) Or (tlAirVlf.iEffDate(1) <> tlVlf.iEffDate(1)) Then
                                    Exit Do
                                End If
                                'no room in the while above
                                gUnpackTimeLong tlAirVlf.iAirTime(0), tlAirVlf.iAirTime(1), True, llAirTime
                                If llAirTime > llAirTimeP60 Then
                                    Exit Do
                                End If
                                'If (tlAirVlf.iAirTime(0) <> tlVlf.iAirTime(0)) Or (tlAirVlf.iAirTime(1) <> tlVlf.iAirTime(1)) Then
                                '    Exit Do
                                'End If
                                'Avoid same record
                                If (tlAirVlf.sStatus = "P") And ((tlAirVlf.iSellCode <> tlVlf.iSellCode) Or (tlAirVlf.iSellTime(0) <> tlVlf.iSellTime(0)) Or (tlAirVlf.iSellTime(1) <> tlVlf.iSellTime(1))) Then
                                    'Create VCF record
                                    If ilNoXLinks = 0 Then
                                        tlVcf.iSellCode = tlVlf.iSellCode
                                        tlVcf.iSellDay = tlVlf.iSellDay
                                        tlVcf.iSellTime(0) = tlVlf.iSellTime(0)
                                        tlVcf.iSellTime(1) = tlVlf.iSellTime(1)
                                        tlVcf.iSellPosNo = tlVlf.iSellPosNo
                                        tlVcf.iEffDate(0) = tlVlf.iEffDate(0)
                                        tlVcf.iEffDate(1) = tlVlf.iEffDate(1)
                                        tlVcf.iTermDate(0) = tlVlf.iTermDate(0)
                                        tlVcf.iTermDate(1) = tlVlf.iTermDate(1)
                                        tlVcf.sDelete = "N"
                                    End If
                                    'If ilNoXLinks < UBound(tlVcf.iCSV) Then
                                    If ilNoXLinks < UBound(tlVcf.iCSV) + 1 Then
                                        ilMatchFd = False
                                        For ilLink = 1 To ilNoXLinks Step 1
                                            If (tlVcf.iCSV(ilLink - 1) = tlAirVlf.iSellCode) And (tlVcf.sCSD(ilLink - 1) = Trim$(str$(tlAirVlf.iSellDay))) And (tlVcf.iCST(0, ilLink - 1) = tlAirVlf.iSellTime(0)) And (tlVcf.iCST(1, ilLink - 1) = tlAirVlf.iSellTime(1)) And (tlVcf.iCSP(ilLink - 1) = tlAirVlf.iSellPosNo) Then
                                                ilMatchFd = True
                                                Exit For
                                            End If
                                        Next ilLink
                                        If Not ilMatchFd Then   'If duplicate not found, add entry
                                            ilNoXLinks = ilNoXLinks + 1
                                            tlVcf.iCSV(ilNoXLinks - 1) = tlAirVlf.iSellCode
                                            tlVcf.sCSD(ilNoXLinks - 1) = Trim$(str$(tlAirVlf.iSellDay))
                                            tlVcf.iCST(0, ilNoXLinks - 1) = tlAirVlf.iSellTime(0)
                                            tlVcf.iCST(1, ilNoXLinks - 1) = tlAirVlf.iSellTime(1)
                                            tlVcf.iCSP(ilNoXLinks - 1) = tlAirVlf.iSellPosNo
                                        End If
                                    End If
                                End If
                                ilRet = btrGetNext(hlVlf, tlAirVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            Loop
                            '12/7/17
                            'ilRet = btrGetDirect(hlVlf, tlVlf, ilVlfRecLen, llVlfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            tlVlfSrchKey0.iSellCode = tlVlf.iSellCode
                            tlVlfSrchKey0.iSellDay = tlVlf.iSellDay
                            tlVlfSrchKey0.iEffDate(0) = tlVlf.iEffDate(0)
                            tlVlfSrchKey0.iEffDate(1) = tlVlf.iEffDate(1)
                            tlVlfSrchKey0.iSellTime(0) = 0
                            tlVlfSrchKey0.iSellTime(1) = 0
                            tlVlfSrchKey0.iSellPosNo = 0
                            tlVlfSrchKey0.iSellSeq = 0
                            ilRet = btrGetGreaterOrEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                            Do While (ilRet = BTRV_ERR_NONE) And (tlVlf.iSellCode = tlVefSch(ilLoop).iVefCode) And (tlVlf.iSellDay = tlVefSch(ilLoop).iDay)
                                If tlVlf.lCode = llVlfCode Then
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hlVlf, tlVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            If ilRet <> BTRV_ERR_NONE Then
                                slMsg = "I/O Error:" & str$(ilRet) & " Get Direct VCF"
                                '6/6/16: Replaced GoSub
                                'GoSub BuildVCFErrorMsg
                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                gBuildVCF = False
                                Exit Function
                            End If
                            'Replaced because of duplicate records now permitted- resulting in endless loop
                            'tlVlfSrchKey0.iSellCode = tlVlf.iSellCode
                            'tlVlfSrchKey0.iSellDay = tlVlf.iSellDay
                            'tlVlfSrchKey0.iEffDate(0) = tlVlf.iEffDate(0)
                            'tlVlfSrchKey0.iEffDate(1) = tlVlf.iEffDate(1)
                            'tlVlfSrchKey0.iSellTime(0) = tlVlf.iSellTime(0)
                            'tlVlfSrchKey0.iSellTime(1) = tlVlf.iSellTime(1)
                            'tlVlfSrchKey0.iSellPosNo = tlVlf.iSellPosNo
                            'tlVlfSrchKey0.iSellSeq = tlVlf.iSellSeq
                            'ilRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                        End If
                        ilRet = btrGetNext(hlVlf, tlVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Loop
                    If ilNoXLinks > 0 Then
                        'Insert record
                        If mVcfInsertTest(hlVcf, tlVcf) Then
                            tlVcf.lCode = 0
                            ilRet = btrInsert(hlVcf, tlVcf, ilVcfRecLen, INDEXKEY1)
                            If ilRet <> BTRV_ERR_NONE Then
                                slMsg = "I/O Error:" & str$(ilRet) & " Insert VCF"
                                '6/6/16: Replaced GoSub
                                'GoSub BuildVCFErrorMsg
                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                gBuildVCF = False
                                Exit Function
                            End If
                        End If
                    End If
                    ilNoXLinks = 0
                    For ilIndex = LBound(tlVcf.iCSV) To UBound(tlVcf.iCSV) Step 1
                        tlVcf.iCSV(ilIndex) = 0
                        tlVcf.sCSD(ilIndex) = " "
                        tlVcf.iCST(0, ilIndex) = 0
                        tlVcf.iCST(1, ilIndex) = 0
                        tlVcf.iCSP(ilIndex) = 0
                    Next ilIndex
                End If
            Next ilDay
        End If
    Next ilLoop
    'Terminate all Current, and convert pending to current
    For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
        If (tlVefSch(ilLoop).sOnAirSchStatus = "S") And (tlVefSch(ilLoop).sType = "S") Then
            'Read in all VLF for vehicle
            For ilDay = 0 To 2 Step 1
                ReDim lmVlfCode(0 To 0) As Long
                tlVlfSrchKey0.iSellCode = tlVefSch(ilLoop).iVefCode
                Select Case ilDay
                    Case 0
                        tlVefSch(ilLoop).iDay = 0
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate0(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate0(1)
                    Case 1
                        tlVefSch(ilLoop).iDay = 6
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate6(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate6(1)
                    Case 2
                        tlVefSch(ilLoop).iDay = 7
                        ilTermDate0 = tlVefSch(ilLoop).iTermDate7(0)
                        ilTermDate1 = tlVefSch(ilLoop).iTermDate7(1)
                End Select
                If (ilTermDate0 <> 0) Or (ilTermDate1 <> 0) Then
                    tlVlfSrchKey0.iSellDay = tlVefSch(ilLoop).iDay
                    tlVlfSrchKey0.iEffDate(0) = 0
                    tlVlfSrchKey0.iEffDate(1) = 0
                    tlVlfSrchKey0.iSellTime(0) = 0
                    tlVlfSrchKey0.iSellTime(1) = 0
                    tlVlfSrchKey0.iSellPosNo = 0
                    tlVlfSrchKey0.iSellSeq = 0
                    ilRet = btrGetGreaterOrEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                    Do While (ilRet = BTRV_ERR_NONE) And (tlVlf.iSellCode = tlVefSch(ilLoop).iVefCode) And (tlVlf.iSellDay = tlVefSch(ilLoop).iDay)
                        If tlVlf.sStatus = "P" Then
                            lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                            ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                        Else
                            If (tlVlf.iTermDate(0) = 0) And (tlVlf.iTermDate(1) = 0) Then
                                gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                                llStartDate = llStartDate + 1
                                gUnpackDateLong tlVlf.iEffDate(0), tlVlf.iEffDate(1), llDate
                                If llStartDate = llDate Then
                                    lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                                    ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                                Else
                                    If (ilTermDate1 < tlVlf.iEffDate(1)) Or ((ilTermDate1 = tlVlf.iEffDate(1)) And (ilTermDate0 < tlVlf.iEffDate(0))) Then
    '                                    lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
    '                                    ReDim Preserve lmVlfCode(1 To UBound(lmVlfCode) + 1) As Long
                                    Else
                                        lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                                        ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                                    End If
                                End If
                            Else    'remove any previously terminated before start
                                If (tlVlf.iTermDate(1) < tlVlf.iEffDate(1)) Or ((tlVlf.iTermDate(1) = tlVlf.iEffDate(1)) And (tlVlf.iTermDate(0) < tlVlf.iEffDate(0))) Then
                                    lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                                    ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                                Else
                                    gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                                    llStartDate = llStartDate + 1
                                    gUnpackDateLong tlVlf.iEffDate(0), tlVlf.iEffDate(1), llDate
                                    If llStartDate = llDate Then
                                        lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                                        ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                                    Else
                                        If (ilTermDate1 > tlVlf.iEffDate(1)) Or ((ilTermDate1 = tlVlf.iEffDate(1)) And (ilTermDate0 > tlVlf.iEffDate(0))) Then
                                            If (ilTermDate1 < tlVlf.iTermDate(1)) Or ((ilTermDate1 = tlVlf.iTermDate(1)) And (ilTermDate0 <= tlVlf.iTermDate(0))) Then
                                                lmVlfCode(UBound(lmVlfCode)) = tlVlf.lCode
                                                ReDim Preserve lmVlfCode(0 To UBound(lmVlfCode) + 1) As Long
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        ilRet = btrGetNext(hlVlf, tlVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Loop
                End If
                For ilVlf = LBound(lmVlfCode) To UBound(lmVlfCode) - 1 Step 1
                    'tlVlf.lCode = lmVlfCode(ilVlf)
                    'tmSRec = tlVlf
                    'ilRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                    'tlVlf = tmSRec
                    tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                    ilRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        slMsg = "I/O Error:" & str$(ilRet) & " GetEqual VLF"
                        '6/6/16: Replaced GoSub
                        'GoSub BuildVCFErrorMsg
                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                        gBuildVCF = False
                        Exit Function
                    End If
                    If tlVlf.sStatus = "P" Then
                        Do
                            'tmSRec = tlVlf
                            'ilRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                            'tlVlf = tmSRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    slMsg = "I/O Error:" & Str$(ilRet) & " Get by Key VLF"
                            '    GoSub BuildVCFErrorMsg
                            '    Exit Function
                            'End If
                            tlVlf.sStatus = "C"
                            ilRet = btrUpdate(hlVlf, tlVlf, ilVlfRecLen)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                'tlVlf.lCode = lmVlfCode(ilVlf)
                                'tmSRec = tlVlf
                                'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                'tlVlf = tmSRec
                                tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilCRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilCRet) & " GetEqual VLF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            slMsg = "I/O Error:" & str$(ilRet) & " Update VLF"
                            '6/6/16: Replaced GoSub
                            'GoSub BuildVCFErrorMsg
                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                            gBuildVCF = False
                            Exit Function
                        End If
                    Else
                        If (tlVlf.iTermDate(0) = 0) And (tlVlf.iTermDate(1) = 0) Then
                            gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                            llStartDate = llStartDate + 1
                            gUnpackDateLong tlVlf.iEffDate(0), tlVlf.iEffDate(1), llDate
                            If llStartDate = llDate Then
                                Do
                                    ilRet = btrDelete(hlVlf)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        'tlVlf.lCode = lmVlfCode(ilVlf)
                                        'tmSRec = tlVlf
                                        'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                        'tlVlf = tmSRec
                                        tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                        ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilCRet <> BTRV_ERR_NONE Then
                                            slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                            '6/6/16: Replaced GoSub
                                            'GoSub BuildVCFErrorMsg
                                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                            gBuildVCF = False
                                            Exit Function
                                        End If
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Delete VLF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                            Else
                                If (ilTermDate1 < tlVlf.iEffDate(1)) Or ((ilTermDate1 = tlVlf.iEffDate(1)) And (ilTermDate0 < tlVlf.iEffDate(0))) Then
                                    'Terminate before start- remove
                                    'The GetNext will still work even when record is deleted
                                    Do
                                        ilRet = btrDelete(hlVlf)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            'tlVlf.lCode = lmVlfCode(ilVlf)
                                            'tmSRec = tlVlf
                                            'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                            'tlVlf = tmSRec
                                            tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                            ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilCRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Delete VLF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                Else
                                    Do
                                        tlVlf.iTermDate(0) = ilTermDate0
                                        tlVlf.iTermDate(1) = ilTermDate1
                                        ilRet = btrUpdate(hlVlf, tlVlf, ilVlfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            'tlVlf.lCode = lmVlfCode(ilVlf)
                                            'tmSRec = tlVlf
                                            'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                            'tlVlf = tmSRec
                                            tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                            ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilCRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Update VLF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else    'remove any previously terminated before start
                            If (tlVlf.iTermDate(1) < tlVlf.iEffDate(1)) Or ((tlVlf.iTermDate(1) = tlVlf.iEffDate(1)) And (tlVlf.iTermDate(0) < tlVlf.iEffDate(0))) Then
                                'Terminate before start- remove
                                'The GetNext will still work even when record is deleted
                                Do
                                    ilRet = btrDelete(hlVlf)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        'tlVlf.lCode = lmVlfCode(ilVlf)
                                        'tmSRec = tlVlf
                                        'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                        'tlVlf = tmSRec
                                        tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                        ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilCRet <> BTRV_ERR_NONE Then
                                            slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                            '6/6/16: Replaced GoSub
                                            'GoSub BuildVCFErrorMsg
                                            mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                            gBuildVCF = False
                                            Exit Function
                                        End If
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    slMsg = "I/O Error:" & str$(ilRet) & " Delete VLF"
                                    '6/6/16: Replaced GoSub
                                    'GoSub BuildVCFErrorMsg
                                    mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                    gBuildVCF = False
                                    Exit Function
                                End If
                            Else
                                'Added code 5/25
                                gUnpackDateLong ilTermDate0, ilTermDate1, llStartDate
                                llStartDate = llStartDate + 1
                                gUnpackDateLong tlVlf.iEffDate(0), tlVlf.iEffDate(1), llDate
                                If llStartDate = llDate Then
                                    Do
                                        ilRet = btrDelete(hlVlf)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            'tlVlf.lCode = lmVlfCode(ilVlf)
                                            'tmSRec = tlVlf
                                            'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                            'tlVlf = tmSRec
                                            tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                            ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilCRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        slMsg = "I/O Error:" & str$(ilRet) & " Delete VLF"
                                        '6/6/16: Replaced GoSub
                                        'GoSub BuildVCFErrorMsg
                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                        gBuildVCF = False
                                        Exit Function
                                    End If
                                Else
                                    If (ilTermDate1 > tlVlf.iEffDate(1)) Or ((ilTermDate1 = tlVlf.iEffDate(1)) And (ilTermDate0 > tlVlf.iEffDate(0))) Then
                                        If (ilTermDate1 < tlVlf.iTermDate(1)) Or ((ilTermDate1 = tlVlf.iTermDate(1)) And (ilTermDate0 <= tlVlf.iTermDate(0))) Then
                                            Do
                                                tlVlf.iTermDate(0) = ilTermDate0
                                                tlVlf.iTermDate(1) = ilTermDate1
                                                ilRet = btrUpdate(hlVlf, tlVlf, ilVlfRecLen)
                                                If ilRet = BTRV_ERR_CONFLICT Then
                                                    'tlVlf.lCode = lmVlfCode(ilVlf)
                                                    'tmSRec = tlVlf
                                                    'ilCRet = gGetByKeyForUpdate("VLF", hlVlf, tmSRec)
                                                    'tlVlf = tmSRec
                                                    tlVlfSrchKey2.lCode = lmVlfCode(ilVlf)
                                                    ilCRet = btrGetEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                                    If ilCRet <> BTRV_ERR_NONE Then
                                                        slMsg = "I/O Error:" & str$(ilCRet) & " Get Equal VLF"
                                                        '6/6/16: Replaced GoSub
                                                        'GoSub BuildVCFErrorMsg
                                                        mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                        gBuildVCF = False
                                                        Exit Function
                                                    End If
                                                End If
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                slMsg = "I/O Error:" & str$(ilRet) & " Update VLF"
                                                '6/6/16: Replaced GoSub
                                                'GoSub BuildVCFErrorMsg
                                                mBuildVCFErrorMsg hlVcf, hlVlf, slMsg, lbcNotSchd
                                                gBuildVCF = False
                                                Exit Function
                                            End If
                                        End If
                                    End If

                                End If
                            End If
                        End If
                    End If
                Next ilVlf
            Next ilDay
        End If
    Next ilLoop
    Erase lmVlfCode
    ilRet = btrEndTrans(hlVcf)
    ilRet = btrClose(hlVlf)
    btrDestroy hlVlf
    ilRet = btrClose(hlVcf)
    btrDestroy hlVcf
    gBuildVCF = True
    Exit Function
'BuildVCFErrorMsg:
'    ilRet = btrAbortTrans(hlVcf)
'    If igBtrError > 0 Then
'        lbcNotSchd.AddItem "I/O Error" & str(igBtrError) & "/" & sgErrLoc & " for " & slMsg
'    Else
'        lbcNotSchd.AddItem slMsg
'    End If
'    Erase lmVlfCode
'    btrDestroy hlVlf
'    btrDestroy hlVcf
'    gBuildVCF = False
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildVehSchInfo                *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build table of vehicles         *
'*                     which indicates type of         *
'*                     scheduling required             *
'*                                                     *
'*******************************************************
Function gBuildVehSchInfo(ilOnAir As Integer, ilAlternate As Integer, tlVefSch() As VEFSCH) As Integer
'
'   ilRet = gBuildVehSchInfo(ilOnAir, ilAlternate, tlVefSch)
'
'   Where:
'       ilOnAir (I) - True=Include vehicles that are within LCF as "On Air"
'       ilAlternate (I)- True=Include vehicles that are within LCF as "Alternate"
'       tlVefSch (O) - Type structure containing schedule info for each vehicle
'                      sOnAirSchStatus- On Air Schedule status- "P"=Pending; "C"=Current; ""=Nothing defined for vehicle
'                      sAltSchStatus- On Air Schedule status- "P"=Pending; "C"=Current; ""=Nothing defined for vehicle
'                      iDay- used by selling and airing only (0=M-F; 6=Sa; 7=Su)
'                      iVefCode- Vehicle code
'                      sType- Vehicle type:"C"-Conventional; "S"-Selling; "A"-Airing; "D"-Delivery
'                      iGroup- Links selling and airing (same group number)
'       ilRet (O)- True if tlVefSch successfully made
'
    Dim ilRet As Integer
    Dim hlVef As Integer        'Vehicle file handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'VEF record length
    Dim hlLcf As Integer        'Library calendar file handle
    Dim tlLcf As LCF
    Dim tlLcfSrchKey As LCFKEY0  'Lcf key record image
    Dim ilLcfRecLen As Integer     'LcF record length
    Dim hlVlf As Integer        'Vehicle link file handle
    Dim tlVlf As VLF
    Dim ilVlfRecLen As Integer     'VLF record length
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim ilOffSet As Integer
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilLastGrpNoAssigned As Integer  'Last group number assigned
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim ilUpperBound As Integer
    Dim ilLimit As Integer
    Dim slDate As String
    Dim ilTmpGrpNo As Integer
    ReDim tlVefSch(0 To 0) As VEFSCH

    ilUpperBound = UBound(tlVefSch)
    tlVefSch(ilUpperBound).sOnAirSchStatus = " "
    tlVefSch(ilUpperBound).sAltSchStatus = " "
    tlVefSch(ilUpperBound).iTermDate0(0) = 0    'added since all days Mo-Fr, Sa, Su go into one record)
    tlVefSch(ilUpperBound).iTermDate0(1) = 0
    tlVefSch(ilUpperBound).iTermDate6(0) = 0
    tlVefSch(ilUpperBound).iTermDate6(1) = 0
    tlVefSch(ilUpperBound).iTermDate7(0) = 0
    tlVefSch(ilUpperBound).iTermDate7(1) = 0
    tlVefSch(ilUpperBound).iVefCode = -1
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gBuildVehSchInfo = False
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    ilVefRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
    ilRet = btrGetFirst(hlVef, tlVef, ilVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If (tlVef.sType <> "L") Then
            If (tlVef.sType = "S") Or (tlVef.sType = "A") Then
                ilLimit = 0 '2
            Else
                ilLimit = 0
            End If
            For ilLoop = 0 To ilLimit Step 1  'For airing and selling create three images
                tlVefSch(ilUpperBound).sOnAirSchStatus = " "
                tlVefSch(ilUpperBound).sAltSchStatus = " "
                tlVefSch(ilUpperBound).iTermDate0(0) = 0    'added since all days Mo-Fr, Sa, Su go into one record)
                tlVefSch(ilUpperBound).iTermDate0(1) = 0
                tlVefSch(ilUpperBound).iTermDate6(0) = 0
                tlVefSch(ilUpperBound).iTermDate6(1) = 0
                tlVefSch(ilUpperBound).iTermDate7(0) = 0
                tlVefSch(ilUpperBound).iTermDate7(1) = 0
                If (tlVef.sType = "S") Or (tlVef.sType = "A") Then
                    Select Case ilLoop
                        Case 0
                            tlVefSch(ilUpperBound).iDay = -1 '0
                        Case 1
                            tlVefSch(ilUpperBound).iDay = 6
                        Case 2
                            tlVefSch(ilUpperBound).iDay = 7
                    End Select
                Else
                    tlVefSch(ilUpperBound).iDay = -1
                End If
                tlVefSch(ilUpperBound).iVefCode = tlVef.iCode
                tlVefSch(ilUpperBound).sType = tlVef.sType
                tlVefSch(ilUpperBound).iGroup = -1
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tlVefSch(0 To ilUpperBound) As VEFSCH
            Next ilLoop
        End If
        ilRet = btrGetNext(hlVef, tlVef, ilVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        gBuildVehSchInfo = False
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    'Determine if any pending exist for vehicle
    hlLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gBuildVehSchInfo = False
        ilRet = btrClose(hlLcf)
        btrDestroy hlLcf
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    ilLcfRecLen = Len(tlLcf)  'btrRecordLength(hlLcf)  'Get and save record length
    If ilOnAir Then
        'For selling and Airing to be set as Requiring scheduling, the linkage
        'records must also be defined, skip this test for selling and airing
        'and only test VLF for pending- as of 1/13/95 this is not true- allow selling
        'and airing vehicle changes without links
        For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
            'If (tlVefSch(ilLoop).sType <> "S") And (tlVefSch(ilLoop).sType <> "A") Then
                tlLcfSrchKey.iType = 0
                tlLcfSrchKey.sStatus = "P"
                tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
                tlLcfSrchKey.iLogDate(0) = 257
                tlLcfSrchKey.iLogDate(1) = 1900
                tlLcfSrchKey.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                If (ilRet = BTRV_ERR_NONE) And (tlLcf.iType = 0) And (tlLcf.sStatus = "P") And (tlLcf.iVefCode = tlVefSch(ilLoop).iVefCode) Then
                    tlVefSch(ilLoop).sOnAirSchStatus = "P"
                Else
                    tlLcfSrchKey.iType = 0
                    tlLcfSrchKey.sStatus = "D"
                    tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
                    tlLcfSrchKey.iLogDate(0) = 257
                    tlLcfSrchKey.iLogDate(1) = 1900
                    tlLcfSrchKey.iSeqNo = 0
                    ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                    If (ilRet = BTRV_ERR_NONE) And (tlLcf.iType = 0) And (tlLcf.sStatus = "D") And (tlLcf.iVefCode = tlVefSch(ilLoop).iVefCode) Then
                        tlVefSch(ilLoop).sOnAirSchStatus = "P"
                    End If
                End If
            'End If
            If tlVefSch(ilLoop).sOnAirSchStatus = " " Then
                tlLcfSrchKey.iType = 0
                tlLcfSrchKey.sStatus = "C"
                tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
                tlLcfSrchKey.iLogDate(0) = 257
                tlLcfSrchKey.iLogDate(1) = 1900
                tlLcfSrchKey.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                If (ilRet = BTRV_ERR_NONE) And (tlLcf.iType = 0) And (tlLcf.sStatus = "C") And (tlLcf.iVefCode = tlVefSch(ilLoop).iVefCode) Then
                    tlVefSch(ilLoop).sOnAirSchStatus = "C"
                End If
            End If
        Next ilLoop
    End If
'    'Selling and airing vehicles don't have alternates
'    If ilAlternate Then
'        For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
'            If (tlVefSch(ilLoop).sType <> "S") And (tlVefSch(ilLoop).sType <> "A") Then
'                tlLcfSrchKey.sType = "A"
'                tlLcfSrchKey.sStatus = "P"
'                tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
'                tlLcfSrchKey.iLogDate(0) = 257
'                tlLcfSrchKey.iLogDate(1) = 1900
'                tlLcfSrchKey.iSeqNo = 0
'                ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
'                If (ilRet = BTRV_ERR_NONE) And (tlLcf.sType = "A") And (tlLcf.sStatus = "P") And (tlLcf.iVefCode = tlVefSch(ilLoop).iVefCode) Then
'                    tlVefSch(ilLoop).sAltSchStatus = "P"
'                Else
'                    tlLcfSrchKey.sType = "A"
'                    tlLcfSrchKey.sStatus = "C"
'                    tlLcfSrchKey.iVefCode = tlVefSch(ilLoop).iVefCode
'                    tlLcfSrchKey.iLogDate(0) = 257
'                    tlLcfSrchKey.iLogDate(1) = 1900
'                    tlLcfSrchKey.iSeqNo = 0
'                    ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
'                    If (ilRet = BTRV_ERR_NONE) And (tlLcf.sType = "A") And (tlLcf.sStatus = "C") And (tlLcf.iVefCode = tlVefSch(ilLoop).iVefCode) Then
'                        tlVefSch(ilLoop).sAltSchStatus = "C"
'                    End If
'                End If
'            End If
'        Next ilLoop
'    End If
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    'Group selling and airing vehicles together if airing pending
    hlVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlVlf, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gBuildVehSchInfo = False
        ilRet = btrClose(hlVlf)
        btrDestroy hlVlf
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    ilVlfRecLen = Len(tlVlf)  'btrRecordLength(hlVlf)  'Get and save record length
    llNoRec = gExtNoRec(ilVlfRecLen) 'Obtain number of records
    btrExtClear hlVlf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlVlf, tlVlf, ilVlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        Call btrExtSetBounds(hlVlf, llNoRec, -1, "UC", "VLF", "") 'Set extract limits (all records)
        tlCharTypeBuff.sType = "P"    'Extract all matching records
        ilOffSet = gFieldOffset("Vlf", "VLFSTATUS")
        ilRet = btrExtAddLogicConst(hlVlf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            gBuildVehSchInfo = False
            ilRet = btrClose(hlVlf)
            btrDestroy hlVlf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            Exit Function
        End If
        ilRet = btrExtAddField(hlVlf, 0, ilVlfRecLen)  'Extract record
        If ilRet <> BTRV_ERR_NONE Then
            gBuildVehSchInfo = False
            ilRet = btrClose(hlVef)
            btrDestroy hlVlf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlVlf)    'Extract record
        ilRet = btrExtGetNext(hlVlf, tlVlf, ilVlfRecLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                gBuildVehSchInfo = False
                ilRet = btrClose(hlVlf)
                btrDestroy hlVlf
                ilRet = btrClose(hlVef)
                btrDestroy hlVef
                Exit Function
            End If
            'ilRet = btrExtGetFirst(hlVlf, tlVlf, ilVlfRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlVlf, tlVlf, ilVlfRecLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
                    If (tlVlf.iSellCode = tlVefSch(ilLoop).iVefCode) Then 'And (tlVlf.iSellDay = tlVefSch(ilLoop).iDay) Then
                        For ilIndex = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
                            If (tlVlf.iAirCode = tlVefSch(ilIndex).iVefCode) Then 'And (tlVlf.iAirDay = tlVefSch(ilLoop).iDay) Then
                                tlVefSch(ilLoop).sOnAirSchStatus = "P"
                                gUnpackDate tlVlf.iEffDate(0), tlVlf.iEffDate(1), slDate
                                slDate = gDecOneDay(slDate)
                                Select Case tlVlf.iSellDay
                                    Case 0
                                        gPackDate slDate, tlVefSch(ilLoop).iTermDate0(0), tlVefSch(ilLoop).iTermDate0(1)
                                    Case 6
                                        gPackDate slDate, tlVefSch(ilLoop).iTermDate6(0), tlVefSch(ilLoop).iTermDate6(1)
                                    Case 7
                                        gPackDate slDate, tlVefSch(ilLoop).iTermDate7(0), tlVefSch(ilLoop).iTermDate7(1)
                                End Select
                                tlVefSch(ilIndex).sOnAirSchStatus = "P"
                                Select Case tlVlf.iAirDay
                                    Case 0
                                        gPackDate slDate, tlVefSch(ilIndex).iTermDate0(0), tlVefSch(ilIndex).iTermDate0(1)
                                    Case 6
                                        gPackDate slDate, tlVefSch(ilIndex).iTermDate6(0), tlVefSch(ilIndex).iTermDate6(1)
                                    Case 7
                                        gPackDate slDate, tlVefSch(ilIndex).iTermDate7(0), tlVefSch(ilIndex).iTermDate7(1)
                                End Select
                                If (tlVefSch(ilLoop).iGroup = -1) And (tlVefSch(ilIndex).iGroup = -1) Then
                                    ilLastGrpNoAssigned = 0
                                    For ilTest = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
                                        If tlVefSch(ilTest).iGroup > ilLastGrpNoAssigned Then
                                            ilLastGrpNoAssigned = tlVefSch(ilTest).iGroup
                                        End If
                                    Next ilTest
                                    tlVefSch(ilLoop).iGroup = ilLastGrpNoAssigned + 1
                                    tlVefSch(ilIndex).iGroup = ilLastGrpNoAssigned + 1
                                ElseIf (tlVefSch(ilLoop).iGroup <> -1) And (tlVefSch(ilIndex).iGroup = -1) Then
                                    tlVefSch(ilIndex).iGroup = tlVefSch(ilLoop).iGroup
                                ElseIf (tlVefSch(ilLoop).iGroup = -1) And (tlVefSch(ilIndex).iGroup <> -1) Then
                                    tlVefSch(ilLoop).iGroup = tlVefSch(ilIndex).iGroup
                                ElseIf tlVefSch(ilLoop).iGroup > tlVefSch(ilIndex).iGroup Then
                                    ilTmpGrpNo = tlVefSch(ilLoop).iGroup
                                    For ilTest = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
                                        If ilTmpGrpNo = tlVefSch(ilTest).iGroup Then
                                            tlVefSch(ilTest).iGroup = tlVefSch(ilIndex).iGroup
                                        ElseIf tlVefSch(ilTest).iGroup > ilTmpGrpNo Then
                                            tlVefSch(ilTest).iGroup = tlVefSch(ilTest).iGroup - 1
                                        End If
                                    Next ilTest
                                ElseIf tlVefSch(ilLoop).iGroup < tlVefSch(ilIndex).iGroup Then
                                    ilTmpGrpNo = tlVefSch(ilIndex).iGroup
                                    For ilTest = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
                                        If ilTmpGrpNo = tlVefSch(ilTest).iGroup Then
                                            tlVefSch(ilTest).iGroup = tlVefSch(ilLoop).iGroup
                                        ElseIf tlVefSch(ilTest).iGroup > ilTmpGrpNo Then
                                            tlVefSch(ilTest).iGroup = tlVefSch(ilTest).iGroup - 1
                                        End If
                                    Next ilTest
                                End If
                                Exit For
                            End If
                        Next ilIndex
                        Exit For    'ilLoop
                    End If
                Next ilLoop
                ilRet = btrExtGetNext(hlVlf, tlVlf, ilVlfRecLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlVlf, tlVlf, ilVlfRecLen, llRecPos)
                Loop
            Loop
        End If
    End If
    ilRet = btrClose(hlVlf)
    btrDestroy hlVlf
    'Test each pending selling, and group with current airing-code not completed- show selling as
    'separate vehicles requiring scheduling
    'For ilLoop = LBound(tlVefSch) To UBound(tlVefSch) - 1 Step 1
    '    If (tlVefSch(ilLoop).sType = "S") And (tlVefSch(ilLoop).sOnAirSchStatus = "P") And (tlVefSch(ilLoop).iGroup = -1) Then
    '        ilVpfIndex = -1
    '        For ilIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
    '            If tmVef.iCode = tgVpf(ilIndex).iVefKCode Then
    '                ilVpfIndex = ilIndex
    '                Exit For
    '            End If
    '        Next ilIndex
    '        If ilVpfIndex >= 0 Then
                'Test if matching airing- if so set as Current
    '        End If
    '    End If
    'Next ilLoop
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    'Check that all links are defined for grouped together vehicles
    'Get all contract avails for the vehicle (bypass local contract avails)
    'and check if a link is defined
    ' call gBuildEventDay
    ' cycle thru extended records to see if match exist- if not then link is
    ' missing
    gBuildVehSchInfo = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveLCFDate                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove LCF for specified date   *
'*                                                     *
'*******************************************************
Function gRemoveLCFDate(hlLcf As Integer, sLCP As String, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer) As Integer
'
'   gRemoveLCFDate hlLcf, slCP, ilVefCode, ilLogDate0, ilLogDate1
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       ilType (I)- 0=Regular Programming; 1->NN = Sports Programming
'       slCP (I)- "C" = Current; "P" = Pending
'       ilVefCode (I)- Vehicle code
'       ilLogDate0 (I)- Log date to be removed
'       ilLogDate1
'
                    'Remove current TFN so new one can be created
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim tlLcfSrchKey2 As LCFKEY2     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llLcfRecPos As Long
    Dim ilVefIndex As Integer
    Dim ilType As Integer

    ilVefIndex = gBinarySearchVef(ilVefCode)
    If ilVefIndex = -1 Then
        gRemoveLCFDate = False
        Exit Function
    End If
    ilLcfRecLen = Len(tlLcf)
    If tgMVef(ilVefIndex).sType <> "G" Then
        ilType = 0
        tlLcfSrchKey.iType = ilType
        tlLcfSrchKey.sStatus = sLCP
        tlLcfSrchKey.iVefCode = ilVefCode
        tlLcfSrchKey.iLogDate(0) = ilLogDate0
        tlLcfSrchKey.iLogDate(1) = ilLogDate1
        tlLcfSrchKey.iSeqNo = 1
        ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Else
        tlLcfSrchKey2.iVefCode = ilVefCode
        tlLcfSrchKey2.iLogDate(0) = ilLogDate0
        tlLcfSrchKey2.iLogDate(1) = ilLogDate1
        ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        ilType = tlLcf.iType
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) And (tlLcf.iLogDate(0) = ilLogDate0) And (tlLcf.iLogDate(1) = ilLogDate1)
        ilRet = btrGetPosition(hlLcf, llLcfRecPos)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            sgErrLoc = "gRemoveLCFDate-Get Position Lcf(1)"
            gRemoveLCFDate = False
            Exit Function
        End If
        Do
            'tmSRec = tlLcf
            'ilCRet = gGetByKeyForUpdate("LCF", hlLcf, tmSRec)
            'tlLcf = tmSRec
            'If ilCRet <> BTRV_ERR_NONE Then
            '    igBtrError = ilCRet
            '    sgErrLoc = "gRemoveLCFDate-Get by Key Lcf(2)"
            '    gRemoveLCFDate = False
            '    Exit Function
            'End If
            ilRet = btrDelete(hlLcf)
            If ilRet = BTRV_ERR_CONFLICT Then
                ilCRet = btrGetDirect(hlLcf, tlLcf, ilLcfRecLen, llLcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilCRet <> BTRV_ERR_NONE Then
                    igBtrError = ilCRet
                    sgErrLoc = "gRemoveLCFDate-Get Direct Lcf(3)"
                    gRemoveLCFDate = False
                    Exit Function
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            sgErrLoc = "gRemoveLCFDate-Delete Lcf(4)"
            gRemoveLCFDate = False
            Exit Function
        End If
        If tgMVef(ilVefIndex).sType <> "G" Then
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = sLCP
            tlLcfSrchKey.iVefCode = ilVefCode
            tlLcfSrchKey.iLogDate(0) = ilLogDate0
            tlLcfSrchKey.iLogDate(1) = ilLogDate1
            tlLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Else
            tlLcfSrchKey2.iVefCode = ilVefCode
            tlLcfSrchKey2.iLogDate(0) = ilLogDate0
            tlLcfSrchKey2.iLogDate(1) = ilLogDate1
            ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            ilType = tlLcf.iType
        End If
    Loop
    gRemoveLCFDate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeLCFFromLLC                 *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Make LCF for specified date     *
'*                                                     *
'*******************************************************
Private Function mMakeLCFFromLLC(hlLcf As Integer, ilType As Integer, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, tlLLC() As LLC) As Integer
'
'   mMakeLCFFromLLC hlLcf, ilType, ilVefCode, ilLogDate0, ilLogDate1, tlLLC
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       ilType (I)- 0=Regular Programming; 1->NN = Sports Programming (Game number)
'       ilVefCode (I)- Vehicle code
'       ilLogDate0 (I)- Log date to be removed
'       ilLogDate1
'       tlLLC (I)- LLC records (current and pending) to create LCF from
'
    Dim tlLcf As LCF                'LCF record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim ilRet As Integer
    Dim ilBuildHdLCF As Integer
    Dim ilLcfIndex As Integer
    Dim ilSeqNo As Integer
    Dim ilIndex As Integer

    ilLcfRecLen = Len(tlLcf)
    ilBuildHdLCF = True
    ilSeqNo = 1
    For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
        If ilBuildHdLCF Then
            tlLcf.iVefCode = ilVefCode
            tlLcf.iLogDate(0) = ilLogDate0
            tlLcf.iLogDate(1) = ilLogDate1
            tlLcf.iSeqNo = ilSeqNo
            tlLcf.iType = ilType
            tlLcf.sStatus = "C"
            tlLcf.sTiming = "N" 'Timing not started
            tlLcf.sAffPost = "N"
            tlLcf.iLastTime(0) = 0
            tlLcf.iLastTime(1) = 0
            tlLcf.iUrfCode = tgUrf(0).iCode
            ilBuildHdLCF = False
            For ilLcfIndex = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
                tlLcf.lLvfCode(ilLcfIndex) = 0
                tlLcf.iTime(0, ilLcfIndex) = 0
                tlLcf.iTime(1, ilLcfIndex) = 0
            Next ilLcfIndex
            ilLcfIndex = LBound(tlLcf.lLvfCode)
        End If
        tlLcf.lLvfCode(ilLcfIndex) = tlLLC(ilIndex).lLvfCode
        gPackTime tlLLC(ilIndex).sStartTime, tlLcf.iTime(0, ilLcfIndex), tlLcf.iTime(1, ilLcfIndex)
        ilLcfIndex = ilLcfIndex + 1
        If ilLcfIndex > UBound(tlLcf.lLvfCode) Then
            tlLcf.lCode = 0
            ilRet = btrInsert(hlLcf, tlLcf, ilLcfRecLen, INDEXKEY3)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                sgErrLoc = "mMakeLCFFromLLC-Insert Lcf(1)"
                mMakeLCFFromLLC = False
                Exit Function
            End If
            ilSeqNo = ilSeqNo + 1
            ilBuildHdLCF = True
        End If
    Next ilIndex
    If Not ilBuildHdLCF Then
        tlLcf.lCode = 0
        ilRet = btrInsert(hlLcf, tlLcf, ilLcfRecLen, INDEXKEY3)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            sgErrLoc = "mMakeLCFFromLLC-Insert Lcf(2)"
            mMakeLCFFromLLC = False
            Exit Function
        End If
    End If
    mMakeLCFFromLLC = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTFNExist                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine earliest date of LCF  *
'*                                                     *
'*******************************************************
Private Function mTFNExist(hlLcf As Integer, ilType As Integer, sLCP As String, ilVefCode As Integer) As Integer
'
'   ilTFN = mTFNExist(hlLcf, ilType, slCP, ilVefCode)
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       ilType (I)- 0=Regular Programming; 1->NN = Sports Programming (Game Number)
'       slCP (I)- "C" = Current; "P" = Pending
'       ilVefCode (I)- Vehicle code
'       ilTFN (O)- True= TFN exist; False = No TFN
'
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim ilRet As Integer
    ilLcfRecLen = Len(tlLcf)
    tlLcfSrchKey.iType = ilType
    tlLcfSrchKey.sStatus = sLCP
    tlLcfSrchKey.iVefCode = ilVefCode
    tlLcfSrchKey.iLogDate(0) = 0  'Year 1/1/1900
    tlLcfSrchKey.iLogDate(1) = 0
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) Then
        If (tlLcf.iLogDate(0) <= 7) And (tlLcf.iLogDate(1) = 0) Then
            mTFNExist = True
        Else
            mTFNExist = False
        End If
    Else
        mTFNExist = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gVcfInsertTest                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test for duplicate VCF          *
'*                                                     *
'*******************************************************
Private Function mVcfInsertTest(hlVcf As Integer, tlVcf As VCF) As Integer
    Dim tlVcfSrchKey0 As VCFKEY0  'Vcf key record image
    Dim ilVcfRecLen As Integer     'VEF record length
    Dim ilInsert As Integer
    Dim tlTestVcf As VCF        'Check for duplicate VCF
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim ilMatch As Integer
    ilInsert = True
    ilVcfRecLen = Len(tlTestVcf)
    tlVcfSrchKey0.iSellCode = tlVcf.iSellCode
    tlVcfSrchKey0.iSellDay = tlVcf.iSellDay
    tlVcfSrchKey0.iEffDate(0) = tlVcf.iEffDate(0)
    tlVcfSrchKey0.iEffDate(1) = tlVcf.iEffDate(1)
    tlVcfSrchKey0.iSellTime(0) = tlVcf.iSellTime(0)
    tlVcfSrchKey0.iSellTime(1) = tlVcf.iSellTime(1)
    tlVcfSrchKey0.iSellPosNo = tlVcf.iSellPosNo
    ilRet = btrGetEqual(hlVcf, tlTestVcf, ilVcfRecLen, tlVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
    'Check all header type field because of the btrGetNext test
    Do While (ilRet = BTRV_ERR_NONE) And (tlTestVcf.iSellCode = tlVcf.iSellCode) And (tlTestVcf.iSellDay = tlVcf.iSellDay) And (tlTestVcf.iEffDate(0) = tlVcf.iEffDate(0)) And (tlTestVcf.iEffDate(1) = tlVcf.iEffDate(1)) And (tlTestVcf.iSellTime(0) = tlVcf.iSellTime(0)) And (tlTestVcf.iSellTime(1) = tlVcf.iSellTime(1)) And (tlTestVcf.iSellPosNo = tlVcf.iSellPosNo)
        ilMatch = False
        If (tlTestVcf.iTermDate(0) = tlVcf.iTermDate(0)) And (tlTestVcf.iTermDate(1) = tlVcf.iTermDate(1)) And (tlTestVcf.sDelete = tlVcf.sDelete) Then
            For ilIndex = LBound(tlVcf.iCSV) To UBound(tlVcf.iCSV) Step 1
                ilMatch = False
                For ilTest = LBound(tlTestVcf.iCSV) To UBound(tlTestVcf.iCSV) Step 1
                    If (tlTestVcf.iCSV(ilTest) = tlVcf.iCSV(ilIndex)) And (tlTestVcf.sCSD(ilTest) = tlVcf.sCSD(ilIndex)) And (tlTestVcf.iCST(0, ilTest) = tlVcf.iCST(0, ilIndex)) And (tlTestVcf.iCST(1, ilTest) = tlVcf.iCST(1, ilIndex)) And (tlTestVcf.iCSP(ilTest) = tlVcf.iCSP(ilIndex)) Then
                        ilMatch = True
                        tlTestVcf.iCSV(ilTest) = -1 'So not rechecked
                        Exit For
                    End If
                Next ilTest
                If Not ilMatch Then
                    Exit For
                End If
            Next ilIndex
        End If
        If Not ilMatch Then
            ilRet = btrGetNext(hlVcf, tlTestVcf, ilVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilInsert = False
            Exit Do
        End If
    Loop
    mVcfInsertTest = ilInsert
    Exit Function
End Function

Private Sub mSetAdjMax(hlLcf As Integer, hlLvf As Integer, sLCP As String, ilVefCode As Integer, ilEvtType() As Integer)
'
'   mSetAdjMax hlLcf, hlLvf, slType, ilVefCode, ilLogDate0, ilLogDate1
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       hlLnf (I)- LVF Handle
'       slType (I)- "O" = On Air; "A" = Alternate
'       slCP (I)- "C" = Current; "P" = Pending
'       ilVefCode (I)- Vehicle code
    Dim tlLvf As LVF                'LVF record image
    Dim tlLvfVersion As LVF                'LVF record image
    Dim tlLvfSrchKey0 As LONGKEY0     'LVF key record image
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim tlLvfSrchKey1 As LVFKEY1     'LVF key record image
    Dim ilLvfRecLen As Integer         'LVF record length
    Dim ilTFNDay As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer

    ilLvfRecLen = Len(tlLvf)
    ReDim tmLibAdjInfo(0 To 0) As LIBADJINFO
    For ilTFNDay = 1 To 7 Step 1
        ReDim tlLLC(0 To 0) As LLC  'Merged library names
        Select Case ilTFNDay
            Case 1
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
            Case 2
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
            Case 3
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
            Case 4
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
            Case 5
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
            Case 6
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
            Case 7
                ilRet = gBuildEventDay(0, "B", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
        End Select
        For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
            'Determine the max for each unique title
            tlLvfSrchKey0.lCode = tlLLC(ilIndex).lLvfCode
            ilRet = btrGetEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                ilFound = False
                For ilLoop = 0 To UBound(tmLibAdjInfo) - 1 Step 1
                    If tmLibAdjInfo(ilLoop).iLtfCode = tlLvf.iLtfCode Then
                        ilFound = True
                        If tlLvf.iVersion = tmLibAdjInfo(ilLoop).iMaxVersion Then
                            tmLibAdjInfo(ilLoop).iMaxExist = True
                        End If
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    tlLvfSrchKey1.iLtfCode = tlLvf.iLtfCode
                    tlLvfSrchKey1.iVersion = 32000
                    ilRet = btrGetGreaterOrEqual(hlLvf, tlLvfVersion, ilLvfRecLen, tlLvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    If (ilRet = BTRV_ERR_NONE) And (tlLvfVersion.iLtfCode = tlLvf.iLtfCode) Then
                        tmLibAdjInfo(UBound(tmLibAdjInfo)).iLtfCode = tlLvf.iLtfCode
                        tmLibAdjInfo(UBound(tmLibAdjInfo)).iMaxVersion = tlLvfVersion.iVersion
                        If tlLvf.iVersion = tmLibAdjInfo(ilLoop).iMaxVersion Then
                            tmLibAdjInfo(UBound(tmLibAdjInfo)).iMaxExist = True
                        Else
                            tmLibAdjInfo(UBound(tmLibAdjInfo)).iMaxExist = False
                        End If
                        ReDim Preserve tmLibAdjInfo(0 To UBound(tmLibAdjInfo) + 1) As LIBADJINFO
                    End If
                End If
            End If
        Next ilIndex
    Next ilTFNDay
End Sub

Private Sub mBuildLCFErrorMsg(hlLcf As Integer, hlLvf As Integer, hlSsf As Integer, hlSdf As Integer, hlSmf As Integer, tlVefSch() As VEFSCH, ilLoop As Integer, slMsg As String, lbcNotSchd As control)
    Dim ilRet As Integer
    ilRet = btrAbortTrans(hlLcf)
    tmVefSrchKey.iCode = tlVefSch(ilLoop).iVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmVef.sName = "Vehicle Name Missing"
    End If
    If igBtrError > 0 Then
        lbcNotSchd.AddItem "Error #" & str(igBtrError) & "/" & sgErrLoc & " for " & Trim$(tmVef.sName) & ": " & slMsg
    Else
        lbcNotSchd.AddItem Trim$(tmVef.sName) & ": " & slMsg
    End If
    btrDestroy hlLcf
    btrDestroy hlLvf
    btrDestroy hlSsf
    btrDestroy hlSdf
    btrDestroy hlSmf
    btrDestroy hmVef
    Erase sgSSFErrorMsg
    'gBuildLCF = False
    gCloseSchFiles
End Sub

Private Sub mBuildVCFErrorMsg(hlVcf As Integer, hlVlf As Integer, slMsg As String, lbcNotSchd As control)
    Dim ilRet As Integer
    
    ilRet = btrAbortTrans(hlVcf)
    If igBtrError > 0 Then
        lbcNotSchd.AddItem "I/O Error" & str(igBtrError) & "/" & sgErrLoc & " for " & slMsg
    Else
        lbcNotSchd.AddItem slMsg
    End If
    Erase lmVlfCode
    btrDestroy hlVlf
    btrDestroy hlVcf
    'gBuildVCF = False
End Sub
