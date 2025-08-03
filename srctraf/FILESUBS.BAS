Attribute VB_Name = "FILESUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FILESUBS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gGetLatestVLFDate             gVpfFindHd                    gGetEarliestAffSpotDate   *
'*                                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FileSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the File subs and functions
Option Explicit
Option Compare Text

Global tgPriceCff As CFF
Dim imLowLimit As Integer
Dim lmLowLimit As Long
' Dan M 4/20/09 limit user and limit changes to site form.
Public bgInternalGuide As Boolean
Public igChangesAllowed As Integer
'Following dim removed as not required in gGetOffset of gGetByKeyForUpdate
'Dim tmAdf As ADF
'Dim tmAgf As AGF
'Dim tmAnf As ANF
'Dim tmArf As ARF
'Dim tmBOf As BOF
'Dim tmBsf As BSF
'Dim tmBvf As BVF
'Dim tmCcf As CCF
'Dim tmCdf As CDF
'Dim tmCef As CEF
'Dim tmCff As CFF
'Dim tmChf As CHF
'Dim tmCif As CIF
'Dim tmClf As CLF
'Dim tmCmf As CMF
'Dim tmCnf As CNF
'Dim tmCof As COF
'Dim tmCpf As CPF
'Dim tmCrf As CRF
'Dim tmCsf As CSF
''Dim tmCtf As CTF
'Dim tmCxf As CXF
'Dim tmCyf As CYF
''Dim tmDaf As DAF
'Dim tmDlf As DLF
'Dim tmDnf As DNF
'Dim tmDrf As DRF
'Dim tmDsf As DSF
''Dim tmElf As ELF
'Dim tmEnf As ENF
'Dim tmEtf As ETF
''Dim tmFsf As FSF
''Dim tmFxf As FXF
''Dim tmGmf As GMF
'Dim tmLcf As LCF
'Dim tmLef As LEF
''Dim tmLhf As LHF   'in XTRAFILE.BAS
''Dim tmLgf As LGF
'Dim tmLtf As LTF
'Dim tmLvf As LVF
'Dim tmMcf As MCF
'Dim tmMnf As MNF
'Dim tmPhf As RVF    'PHF
'Dim tmPjf As PJF
'Dim tmPnf As PNF    'PHF
'Dim tmPrf As PRF
'Dim tmPvf As PVF
'Dim tmRcf As RCF
'Dim tmRdf As RDF
''Dim tmRgf As RGF
'Dim tmRif As RIF
''Dim tmRpf As RPF
'Dim tmRvf As RVF
'Dim tmSbf As SBF
'Dim tmSdf As SDF
''Dim tmSff As SFF
''Dim tmShf As SHF   'Spot History
'Dim tmSif As SIF
'Dim tmSlf As SLF
'Dim tmSmf As SMF
'Dim tmSof As SOF
'Dim tmSpf As SPF
'Dim tmStf As STF
'Dim tmSvf As SVF
'Dim tmTzf As TZF
'Dim tmUrf As URF
'Dim tmVcf As VCF
'Dim tmVlf As VLF
'Dim tmVpf As VPF
Dim tmVsf As VSF

Dim tmLongSrchKey As LONGKEY0    'Key record image

'Alert File
Dim tmAufSrchKey0 As LONGKEY0
Dim tmAufSrchKey1 As AUFKEY1
Dim tmAufSrchKey2 As AUFKEY2
Dim tmAufSrchKey3 As AUFKEY3
Dim tmAuf As AUF
Dim imAufRecLen As Integer
Dim tmUlfSrchKey0 As LONGKEY0
Dim tmUlf As ULF
Dim imUlfRecLen As Integer

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertFound                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Test if Alert exist       *
'*                                                     *
'*******************************************************
Public Function gAlertFound(slType As String, slSubType As String, llChfCode As Long, ilVefCode As Integer, slInWeekDate As String) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim llMoWeekDate As Long
    Dim tlAuf As AUF

    If slType <> "C" Then
        llMoWeekDate = gDateValue(slInWeekDate)
        Do While gWeekDayLong(llMoWeekDate) <> 0
            llMoWeekDate = llMoWeekDate - 1
        Loop
    Else
        llMoWeekDate = 0
    End If

    ilFound = False
    imAufRecLen = Len(tlAuf)
    '7/15/14
    If slType = "L" Or slType = "P" Then
        tmAufSrchKey2.sType = slType
        tmAufSrchKey2.sStatus = "R"
        tmAufSrchKey2.sSubType = slSubType
        tmAufSrchKey2.iVefCode = ilVefCode
        gPackDateLong llMoWeekDate, tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
        ilRet = btrGetEqual(hgAuf, tlAuf, imAufRecLen, tmAufSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Else
        tmAufSrchKey3.sType = slType
        tmAufSrchKey3.sStatus = "R"
        tmAufSrchKey3.sSubType = slSubType
        tmAufSrchKey3.lChfCode = llChfCode
        ilRet = btrGetEqual(hgAuf, tlAuf, imAufRecLen, tmAufSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (slType = tlAuf.sType) And (tlAuf.sStatus = "R")
        If slType = "C" Then
            If (llChfCode = tlAuf.lChfCode) And (slSubType = tlAuf.sSubType) Then
                ilFound = True
            End If
            Exit Do
        ElseIf slType = "L" Then
            If (ilVefCode = tlAuf.iVefCode) And (slSubType = tlAuf.sSubType) Then
                gUnpackDateLong tlAuf.iMoWeekDate(0), tlAuf.iMoWeekDate(1), llDate
                If llDate = llMoWeekDate Then
                    ilFound = True
                End If
            End If
            Exit Do
        ElseIf slType = "P" Then
            If (ilVefCode = tlAuf.iVefCode) And (slSubType = tlAuf.sSubType) Then
                gUnpackDateLong tlAuf.iMoWeekDate(0), tlAuf.iMoWeekDate(1), llDate
                If llDate = llMoWeekDate Then
                    ilFound = True
                End If
            End If
            Exit Do
        End If
        ilRet = btrGetNext(hgAuf, tlAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get next record
    Loop
    If ilFound Then
        gAlertFound = True
    Else
        gAlertFound = False
    End If

End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gAnyRepDef                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Any market set as Cluster 'Yes' *
'*                                                     *
'*******************************************************
Sub gAnyRepDef()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'
'   Set sgRepDef to N or Y depending on if any market is set to cluster Yes
'   also, create igRepVefCode which contains VefCode for each vehicle that is Type = "R"
'
    Dim ilRet As Integer
    Dim ilVeh As Integer
    Dim ilClusterFound As Integer
    Dim ilCluster As Integer

    gSpfRead
    'gSetUsingTraffic
    ilRet = gObtainVef()
    sgRepDef = "N"
    ReDim igRepVefCode(0 To 0) As Integer
    ilRet = gObtainMnfForType("H3", sgMktMnfStamp, tgMktMnf())
    If ilRet = False Then
        Exit Sub
    End If
    'If tgSpf.sMktBase <> "Y" Then
    If (Asc(tgSpf.sUsingFeatures) And USINGREP) <> USINGREP Then
        Exit Sub
    End If
    'For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
    '    If Trim$(tgMktMnf(ilLoop).sRPU) <> "Y" Then
            For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If (tgMVef(ilVeh).sType = "R") Then
    '                If tgMVef(ilVeh).iMnfVehGp3Mkt = tgMktMnf(ilLoop).iCode Then
                        ilClusterFound = False
                        If sgMktClusterDef <> "N" Then
                            For ilCluster = LBound(igMktVefCode) To UBound(igMktVefCode) - 1 Step 1
                                If igMktVefCode(ilCluster) = tgMVef(ilVeh).iCode Then
                                    ilClusterFound = True
                                    Exit For
                                End If
                            Next ilCluster
                        End If
                        If Not ilClusterFound Then
                            sgRepDef = "Y"
                            igRepVefCode(UBound(igRepVefCode)) = tgMVef(ilVeh).iCode
                            ReDim Preserve igRepVefCode(0 To UBound(igRepVefCode) + 1) As Integer
                        End If
    '                End If
                End If
            Next ilVeh
    '    End If
    'Next ilLoop
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gAnyClustersDef                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Any market set as Cluster 'Yes' *
'*                                                     *
'*******************************************************
Sub gAnyClustersDef()
'
'
'   Set sgMktClusterDef to N or Y depending on if any market is set to cluster Yes
'   also, create igMktVefCode which contains VefCode for each vehicle that is associated with Market Cluster = Yes
'   This will be only for clients that are Market Based.
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVeh As Integer

    gSpfRead
    'gSetUsingTraffic
    ilRet = gObtainVef()
    sgMktClusterDef = "N"
    ReDim igMktVefCode(0 To 0) As Integer
    ilRet = gObtainMnfForType("H3", sgMktMnfStamp, tgMktMnf())
    If ilRet = False Then
        Exit Sub
    End If
    If tgSpf.sMktBase <> "Y" Then
        Exit Sub
    End If
    For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
        If Trim$(tgMktMnf(ilLoop).sRPU) = "Y" Then
            For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '9/7/02- Jim only Rep vehicles will be Cluster or Non-Cluster
                'If (tgMVef(ilVeh).sType = "C") Or (tgMVef(ilVeh).sType = "S") Then
                '    If tgMVef(ilVeh).iMnfVehGp3Mkt = tgMktMnf(ilLoop).iCode Then
                '        sgMktClusterDef = "Y"
                '        igMktVefCode(UBound(igMktVefCode)) = tgMVef(ilVeh).iCode
                '        ReDim Preserve igMktVefCode(0 To UBound(igMktVefCode) + 1) As Integer
                '    End If
                'ElseIf (tgMVef(ilVeh).sType = "R") Then
                If (tgMVef(ilVeh).sType = "R") Then
                    If tgMVef(ilVeh).iMnfVehGp3Mkt = tgMktMnf(ilLoop).iCode Then
                        sgMktClusterDef = "Y"
                        igMktVefCode(UBound(igMktVefCode)) = tgMVef(ilVeh).iCode
                        ReDim Preserve igMktVefCode(0 To UBound(igMktVefCode) + 1) As Integer
                    End If
                End If
            Next ilVeh
        End If
    Next ilLoop
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:gBtrvErrorMsg                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show error btrieve message      *
'*                                                     *
'*******************************************************
Sub gBtrvErrorMsg(ilErrCode As Integer, slMess As String, frm As Form)
'
'   slMess = "Open Arf"
'   gBtrvErrorMsg ilRet, slMess, MainForm
'   Where:
'       ilErrCode (I)- Error number from btrieve
'       slMess (I) - Message to be added to the btrieve error message
'       MainForm (I)- Name of Form to unload if error exists
'

    Dim slErrMsg As String

    If ilErrCode = BTRV_ERR_NONE Then
        Exit Sub
    End If
    If ilErrCode = BTRV_ERR_REJECT_COUNT Then
        Exit Sub
    End If
    If ilErrCode = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    End If
    If ilErrCode = BTRV_ERR_FILTER_LIMIT Then
        Exit Sub
    End If
    slErrMsg = gGetBtrvErrorMsg(ilErrCode)  '"Btrieve: "
    Screen.MousePointer = vbDefault
    slErrMsg = slErrMsg & sgLF & frm.Caption & ": " & slMess
'    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
'    If ilErrCode = 30000 Then
'        MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
'        Exit Sub
'    Else
'        gErrorApplStop "FileSubs: " & slErrMsg
'    End If
    If ilErrCode = 30000 Then
        'If igBkgdProg = 0 Then
        '    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
        'ElseIf igBkgdProg = 1 Then
        '    gLogMsg slErrMsg, "Bkgd_Schd.Txt", False
        'ElseIf igBkgdProg = 2 Then
        '    gLogMsg slErrMsg, "Set_Credit.Txt", False
        'Else
        '    gLogMsg slErrMsg, "TrafficErrors.Txt", False
        'End If
        gMsgBox slErrMsg, vbOKOnly + vbCritical + vbApplicationModal, "Btrieve Error"
        Exit Sub
    Else
        gErrorApplStop "FileSubs: " & slErrMsg
    End If
'    If TypeOf Frm Is Traffic Then
'        igManUnload = True
'        Unload Frm
'        igManUnload = False
'        Error ERRORCODEBASE
'
''        End
'    Else
''        Stop
'        'igManUnload = True
'        'Unload Frm
'        'igManUnload = False
'        Error ERRORCODEBASE
'    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gChrRefExist                    *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified character*
'*                     exists in specified file        *
'*                                                     *
'*******************************************************
Function gChrRefExist(frm As Form, slString As String, slFileName As String, slFieldName As String, Optional blIsNotEqual As Boolean = False) As Integer
' optional added for ttp 5329. needed way to test if idc was being used.
'   ilRet = gChrRefExist (MainForm, slString, sFileName, sFieldName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       slString (I)- Char or string to be checked
'       sFileName (I)- File name
'       sFieldName (I)- field name of integer item in the record
'       ilRet (O)- True = Found
'                       False = not found
'

    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    'Dim tlStrTypeBuff As POPSTRINGTYPE   'Type field record
    'ttp 5329 change type to hold more than 1 char.
    'Dim tlStrTypeBuff As POPCHARTYPE   'Type field record
    Dim tlStrTypeBuff As SORTCODE
    Dim tlRecBuff As ARRAYREC '
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilStrOffset As Integer
    
    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slFile = Left$(slFileName, ilPos - 1)
    Else
        slFile = slFileName
    End If
    ilStrOffset = gFieldOffset(slFile, slFieldName)
    If ilStrOffset < 0 Then
        gChrRefExist = True
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gChrRefExistErr
    gBtrvErrorMsg ilRet, "gChrRefExist (btrOpen):" & slFileName, frm
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
        On Error GoTo gChrRefExistErr
        gBtrvErrorMsg ilRet, "gChrRefExist (btrReset):" & slFileName, frm
        On Error GoTo 0
        btrDestroy hlFile
        gChrRefExist = False
        Exit Function
    Else
        On Error GoTo gChrRefExistErr
        gBtrvErrorMsg ilRet, "gChrRefExist (btrGetFirst):" & slFileName, frm
        On Error GoTo 0
    End If
    llNoRec = gExtNoRec(ilRecLen)  'Force cach to be large
    ilPos = InStr(slFileName, ".")
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", Left$(slFileName, ilPos - 1), "") 'Set extract limits (all records)
    tlStrTypeBuff.sKey = slString    'Extract all matching records
    If blIsNotEqual Then
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilStrOffset, Len(slString), BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slString))
    Else
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilStrOffset, Len(slString), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slString))
    End If
    On Error GoTo gChrRefExistErr
    gBtrvErrorMsg ilRet, "gChrRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract character field
    If ddExtRecordCount(hlFile) > 0 Then
        gChrRefExist = True
    Else
        gChrRefExist = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gChrRefExistErr
    gBtrvErrorMsg ilRet, "gChrRefExist (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gChrRefExistErr:
    gDbg_HandleError "FileSubs: gChrRefExist"
'    gChrRefExist = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCodeChrRefExist                *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified code &   *
'*                     character exists in specified   *
'*                     file                            *
'*                                                     *
'*******************************************************
Function gCodeChrRefExist(frm As Form, slFileName As String, ilMatchCode As Integer, slCodeFieldName As String, slMatchChr As String, slChrFieldName As String) As Integer
'
'   ilRet = gCodeChrRefExist (MainForm, slFileName, ilMatchCode, slCodeFieldName, slMatchChr, sChrFieldName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       sFileName (I)- File name
'       ilMatchCode (I)- code to be matched
'       slCodeFieldName (I)- field name of integer item in the record
'       slMatchChr (I)- Char  to be checked
'       sFieldName (I)- field name of character item in the record
'       ilRet (O)- True = Found
'                       False = not found
'

    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    'Dim tlChrTypeBuff As POPSTRINGTYPE   'Type field record
    Dim tlIntTypeBuff As POPICODE   'Type field record
    Dim tlChrTypeBuff As POPCHARTYPE   'Type field record
    Dim tlRecBuff As ARRAYREC
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilCodeOffset As Integer
    Dim ilChrOffset As Integer

    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slFile = Left$(slFileName, ilPos - 1)
    Else
        slFile = slFileName
    End If
    ilCodeOffset = gFieldOffset(slFile, slCodeFieldName)
    If ilCodeOffset < 0 Then
        gCodeChrRefExist = True
        Exit Function
    End If
    ilChrOffset = gFieldOffset(slFile, slChrFieldName)
    If ilChrOffset < 0 Then
        gCodeChrRefExist = True
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCodeChrRefExistErr
    gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrOpen):" & slFileName, frm
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
        On Error GoTo gCodeChrRefExistErr
        gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrReset):" & slFileName, frm
        On Error GoTo 0
        btrDestroy hlFile
        gCodeChrRefExist = False
        Exit Function
    Else
        On Error GoTo gCodeChrRefExistErr
        gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrGetFirst):" & slFileName, frm
        On Error GoTo 0
    End If
    llNoRec = gExtNoRec(ilRecLen)  'Force large cach area
    ilPos = InStr(slFileName, ".")
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", Left$(slFileName, ilPos - 1), "") 'Set extract limits (all records)
    tlIntTypeBuff.iCode = ilMatchCode    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
    On Error GoTo gCodeChrRefExistErr
    gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    tlChrTypeBuff.sType = slMatchChr    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilChrOffset, Len(slMatchChr), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlChrTypeBuff, Len(slMatchChr))
    On Error GoTo gCodeChrRefExistErr
    gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract character field
    On Error GoTo gCodeChrRefExistErr
    gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    If ddExtRecordCount(hlFile) > 0 Then
        gCodeChrRefExist = True
    Else
        gCodeChrRefExist = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gCodeChrRefExistErr
    gBtrvErrorMsg ilRet, "gCodeChrRefExist (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gCodeChrRefExistErr:
    gDbg_HandleError "FileSubs: gCodeChrRefExist"
'    gCodeChrRefExist = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCodeInUser                     *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified vehicle  *
'*                     or salesperson code number      *
'*                     exists in user file             *
'*                                                     *
'*******************************************************
Function gCodeInUser(frm As Form, slFilter As String, ilCodeNo As Integer) As Integer


'   ilRet = gVefInUser (MainForm, slFilter, iCodeNo)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       slFilter (I)- "V" =Vehicle or "S" = salesperson
'       lCodeNo (I)- Code number to be checked
'       ilRet (O)- True = Found
'                  False = not found
'

    Dim hlUrf As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim tlUrf As URF   'file record

    hlUrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCodeInUserErr
    gBtrvErrorMsg ilRet, "gCodeInUser (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlUrf) 'Get and save record length
    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        gUrfDecrypt tlUrf
        If tlUrf.sDelete <> "Y" Then
            If slFilter = "V" Then
                If (ilCodeNo = tlUrf.iVefCode) Then
                    ilRet = btrClose(hlUrf)
                    btrDestroy hlUrf
                    gCodeInUser = True
                    Exit Function
                End If
            ElseIf slFilter = "S" Then
                If (ilCodeNo = tlUrf.iSlfCode) Then
                    ilRet = btrClose(hlUrf)
                    btrDestroy hlUrf
                    gCodeInUser = True
                    Exit Function
                End If
            End If
        End If
        ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo gCodeInUserErr
        gBtrvErrorMsg ilRet, "gCodeInUser (btrGetFirst or Next):" & "Urf.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    gCodeInUser = False
    Exit Function
gCodeInUserErr:
    gDbg_HandleError "FileSubs: gCodeInUser"
'    gCodeInUser = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCPErrorMsg                     *
'*                                                     *
'*             Created:4/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show Counterpoint Software      *
'*                     error message                   *
'*                                                     *
'*******************************************************
Sub gCPErrorMsg(ilErrCode As Integer, slMess As String, frm As Form)
'
'   slMess = "Code number"
'   gCPErrorMsg ilErrCode, slMess, MainForm
'   Where:
'       ilErrCode (I)- error code set by Counterpoint Function/Sub
'       slMess (I) - Message to be added to standard error message
'       MainForm (I)- Name of Form to unload if error
'

    Dim slErrMsg As String
    If (ilErrCode = CP_MSG_NONE) Or (ilErrCode = CP_MSG_POPREQ) Or (ilErrCode = CP_MSG_NOPOPREQ) Then
        Exit Sub
    End If
    slErrMsg = "Counterpoint: "
    Select Case ilErrCode
        Case CP_MSG_PARSE    '1
            slErrMsg = slErrMsg & "Parse error (1)"
    End Select
    slErrMsg = slErrMsg & sgLF & frm.Caption & ": " & slMess
    If ilErrCode <> CP_MSG_NOSHOW Then
        'If igBkgdProg = 0 Then
        '    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Counterpoint Error"
        'ElseIf igBkgdProg = 1 Then
        '    gLogMsg slErrMsg, "Bkgd_Schd.Txt", False
        'ElseIf igBkgdProg = 2 Then
        '    gLogMsg slErrMsg, "Set_Credit.Txt", False
        'Else
        '    gLogMsg slErrMsg, "TrafficErrors.Txt", False
        'End If
        gMsgBox slErrMsg, vbOKOnly + vbCritical + vbApplicationModal, "Counterpoint Error"
    Else
        gErrorApplStop "FileSubs:" & slErrMsg
    End If
'    If TypeOf Frm Is Traffic Then
'        igManUnload = True
'        Unload Frm
'        igManUnload = False
'        Error ERRORCODEBASE
''       End
'    Else
'        igManUnload = True
'        Unload Frm
'        igManUnload = False
'        Error ERRORCODEBASE
'    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gDecryptField                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Decrypt data field              *
'*                                                     *
'*******************************************************
Function gDecryptField(slEncryptName As String) As String
    Dim slDecrypt As String
    Dim slName As String
    Dim ilLen As Integer
    Dim ilLoop As Integer

    slName = slEncryptName
    ilLen = Len(slEncryptName)
    slDecrypt = ""
    For ilLoop = 1 To ilLen Step 1
        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    gDecryptField = slDecrypt
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gFieldOffset                    *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the offset to a field *
'*                     within a record                 *
'*                                                     *
'*******************************************************
Function gFieldOffset(slFile As String, slField As String) As Integer
'
'   sFile = "ADF"
'   sField = "slfFirstName"
'   ilOffset = gFieldOffset(sFile, sField)
'   Where:
'       sFile (I)- Name of the file
'       sField (I)- Name of the field as in the DDF
'       ilOffset (O)- The offset of the start of the field from 0 (-1 if not found)
'

    Dim slMsg As String
    Dim ilOffSet As Integer

    slMsg = ""
    ilOffSet = csiGetOffset(UCase$(slFile), UCase$(slField))
    If ilOffSet >= 0 Then
        gFieldOffset = ilOffSet
    Else
        slMsg = "Offset to field " & slField & " missing from file " & slFile
    End If
'    Select Case UCase$(slFile)
'        Case "ADF"
'            Select Case UCase$(slField)
'                Case "ADFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iCode)
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iCode) '0
'                Case "ADFNAME"
'                    gFieldOffset = 2
'                Case "ADFABBR"
'                    gFieldOffset = 32
'                Case "ADFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iSlfCode) '49
'                Case "ADFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iAgfCode) '51
'                Case "ADFMNFCOMP1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iMnfComp(0)) '103
'                Case "ADFMNFCOMP2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iMnfComp(1)) '105
'                Case "ADFMNFEXCL1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iMnfExcl(0)) '107
'                Case "ADFMNFEXCL2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iMnfExcl(1)) '109
'                Case "ADFCREDITRESTR"
'                    gFieldOffset = 185
'                Case "ADFCREDITLIMIT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.lCreditLimit) '109
'                Case "ADFPAYMRATING"
'                    gFieldOffset = 190
'                Case "ADFMNFSORT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iMnfSort) '160
'                Case "ADFBILAGYDIR"
'                    gFieldOffset = 194
'                Case "ADFARFLKCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iArfLkCode) '313
'                Case "ADFARFCONTRCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iArfCntrCode) '339
'                Case "ADFARFINVCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iArfInvCode) '341
'                Case "ADFPNFBUYER"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iPnfBuyer) '265
'                Case "ADFPNFPAY"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iPnfPay) '265
'                Case "ADFPCT90"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iPct90) '341
'                Case "ADFCURRAR"
'                    gFieldOffset = 480
'                Case "ADFUNBILLED"
'                    gFieldOffset = 486
'                Case "ADFDATELSTPAYM"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iDateLstPaym(0)) '35
'                Case "ADFAVGTOPAY"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iAvgToPay)
'                Case "ADFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iUrfCode) '395
'                Case "ADFSTATE"
'                    gFieldOffset = 533
'                Case "ADFREMOTEID"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iRemoteID) '35
'                Case "ADFAUTOCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iAutoCode) '35
'                Case "ADFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAdf, tmAdf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "AGF"
'            Select Case UCase$(slField)
'                Case "AGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iCode) '0
'                Case "AGFNAME"
'                    gFieldOffset = 2
'                Case "AGFABBR"
'                    gFieldOffset = 42
'                Case "AGFCITY"
'                    gFieldOffset = 47
'                Case "AGFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iSlfCode) '35
'                Case "AGFCREDITRESTR"
'                    gFieldOffset = 96
'                Case "AGFCREDITLIMIT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.lCreditLimit) '35
'                Case "AGFPAYMRATING"
'                    gFieldOffset = 101
'                Case "AGFMNFSORT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iMnfSort) '85
'                Case "AGFARFLKCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iArfLkCode) '237
'                Case "AGFARFCONTRCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iArfCntrCode) '263
'                Case "AGFARFINVCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iArfInvCode) '265
'                Case "AGFPNFBUYER"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iPnfBuyer) '265
'                Case "AGFPNFPAY"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iPnfPay) '265
'                Case "AGFPCT90"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iPct90) '265
'                Case "AGFCURRAR"
'                    gFieldOffset = 390
'                Case "AGFUNBILLED"
'                    gFieldOffset = 396
'                Case "AGFDATELSTPAYM"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iDateLstPaym(0)) '35
'                Case "AGFAVGTOPAY"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iAvgToPay)
'                Case "AGFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iUrfCode) '318
'                Case "AGFSTATE"
'                    gFieldOffset = 438
'                Case "AGFREMOTEID"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iRemoteID) '35
'                Case "AGFAUTOCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iAutoCode) '35
'                Case "AGFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAgf, tmAgf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "ANF"
'            Select Case UCase$(slField)
'                Case "ANFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmAnf, tmAnf.iCode) '0
'                Case "ANFNAME"
'                    gFieldOffset = 2
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "ARF"
'            Select Case UCase$(slField)
'                Case "ARFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmArf, tmArf.iCode) '0
'                Case "ARFTYPE"
'                    gFieldOffset = 2
'                Case "ARFID"
'                    gFieldOffset = 3
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "BOF"
'            Select Case UCase$(slField)
'                Case "BOFTYPE"
'                    gFieldOffset = 4 '0
'                Case "BOFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBOf, tmBOf.iAdfCode) '0
'                Case "BOFENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBOf, tmBOf.iEndDate(0)) '0
'                Case "BOFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBOf, tmBOf.iUrfCode) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "BSF"
'            Select Case UCase$(slField)
'                Case "BSFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBsf, tmBsf.iSlfCode) '0
'                Case "BSFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBsf, tmBsf.iYear) '0
'                Case "BSFSEQNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBsf, tmBsf.iSeqNo) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "BVF"
'            Select Case UCase$(slField)
'                Case "BVFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.lCode) '0
'                Case "BVFSOFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.iSofCode) '0
'                Case "BVFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.iVefCode) '0
'                Case "BVFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.iYear) '0
'                Case "BVFMNFBUDGET"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.iMnfBudget) '0
'                Case "BVFSEQNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmBvf, tmBvf.iSeqNo) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CCF"
'            Select Case UCase$(slField)
'                Case "CCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCode) '0
'                Case "CCFCIFCODE1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCifCode(0)) '2
'                Case "CCFCIFCODE2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCifCode(1)) '4
'                Case "CCFCIFCODE3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCifCode(2)) '6
'                Case "CCFCIFCODE4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCifCode(3)) '8
'                Case "CCFCPFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCcf, tmCcf.lCpfCode) '10
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CDF"
'            Select Case UCase$(slField)
'                Case "CDFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.lCode) '0
'                Case "CDFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iAgfCode) '0
'                Case "CDFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iAdfCode) '2
'                Case "CDFACTIONDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iActionDate(0)) '4
'                Case "CDFENTRYDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iDateEntrd(0)) '8
'                Case "CDFENTRYTIME"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iTimeEntrd(0)) '8
'                Case "CDFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCdf, tmCdf.iUrfCode) '8
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CEF"
'            Select Case UCase$(slField)
'                Case "CEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCef, tmCef.lCode) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CFF"
'            Select Case UCase$(slField)
'                Case "CFFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCff, tmCff.lChfCode) '0
'                Case "CFFCLFLINE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCff, tmCff.iClfLine) '4
'                Case "CFFCNTREVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCff, tmCff.iCntRevNo) '6
'                Case "CFFSTARTDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCff, tmCff.iStartDate(0)) '8
'                Case "CFFDELETE"
'                    gFieldOffset = 41
'                Case "CFFPROPVER"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCff, tmCff.iPropVer) '6
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CHF"
'            Select Case UCase$(slField)
'                Case "CHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lCode) '0
'                Case "CHFCNTRNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lCntrNo) '4
'                Case "CHFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lVefCode) '224
'                Case "CHFEXTREVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iExtRevNo) '59
'                Case "CHFOHDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iOHDDate(0)) '220
'                Case "CHFCNTREVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iCntRevNo) '59
'                Case "CHFPROPDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iPropDate(0)) '220
'                Case "CHFTYPE"
'                    gFieldOffset = 44
'                Case "CHFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iAdfCode) '35
'                Case "CHFPRODUCT"
'                    gFieldOffset = 47
'                Case "CHFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iAgfCode) '57
'                Case "CHFSLFCODE1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iSlfCode(0)) '59
'                Case "CHFMNFDEMO1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfDemo(0)) '59
'                Case "CHFMNFCOMP1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfComp(0)) '101
'                Case "CHFMNFCOMP2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfComp(1)) '103
'                Case "CHFMNFEXCL1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfExcl(0)) '105
'                Case "CHFMNFEXCL2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfExcl(1)) '107
'                Case "CHFMNFREVBK1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfRevSet(0)) '147
'                Case "CHFMNFREVBK2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfRevSet(1)) '149
'                Case "CHFMNFREVBK3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfRevSet(2)) '147
'                Case "CHFMNFREVBK4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfRevSet(3)) '147
'                Case "CHFMNFREVBK5"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfRevSet(4)) '147
'                Case "CHFRCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iRcfCode) '194
'                Case "CHFCMFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lCxfCode) '207
'                Case "CHFCXFINT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lCxfInt) '207
'                Case "CHFPROPVER"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iPropVer) '59
'                Case "CHFSTATUS"
'                    gFieldOffset = 288
'                Case "CHFMNFBUS"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfBus) '207
'                Case "CHFMNFCMPY0"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(0)) '207
'                Case "CHFMNFCMPY1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(1)) '207
'                Case "CHFMNFCMPY2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(2)) '207
'                Case "CHFMNFCMPY3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(3)) '207
'                Case "CHFMNFCMPY4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(4)) '207
'                Case "CHFMNFCMPY5"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(5)) '207
'                Case "CHFMNFCMPY6"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfCmpy(6)) '207
'                Case "CHFMNFPOTNTYPE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfPotnType) '207
'                Case "CHFSCHSTATUS"
'                    gFieldOffset = 369
'                Case "CHFSTARTDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iStartDate(0)) '220
'                Case "CHFENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iEndDate(0)) '220
'                Case "CHFMNFSEG"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iMnfSeg)
'                Case "CHFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iUrfCode) '226
'                Case "CHFSELLNET"
'                    gFieldOffset = 383
'                Case "CHFDELETE"
'                    gFieldOffset = 401
'                Case "CHFSIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.lSifCode) '226
'                Case "CHFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmChf, tmChf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CIF"
'            Select Case UCase$(slField)
'                Case "CIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.lCode) '0
'                Case "CIFMCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iMcfCode) '4
'                Case "CIFNAME"
'                    gFieldOffset = 6
'                Case "CIFCUT"
'                    gFieldOffset = 11
'                Case "CIFREEL"
'                    gFieldOffset = 12
'                Case "CIFLEN"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iLen) '22
'                Case "CIFETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iEtfCode) '24
'                Case "CIFENFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iEnfCode) '26
'                Case "CIFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iAdfCode) '28
'                Case "CIFCPFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.lCpfCode) '30
'                Case "CIFMNFCOMP1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iMnfComp(0)) '34
'                Case "CIFMNFCOMP2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iMnfComp(1)) '36
'                Case "CIFMNFANN"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iMnfAnn) '38
'                Case "CIFCSFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.lCsfCode) '40
'                Case "CIFPURGED"
'                    gFieldOffset = 50
'                Case "CIFROTSTARTDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iRotStartDate(0)) '61
'                Case "CIFROTENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iRotEndDate(0)) '65
'                Case "CIFPURGEDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iPurgeDate(0)) '65
'                Case "CIFENTRYDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iDateEntrd(0)) '65
'                Case "CIFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCif, tmCif.iUrfCode) '69
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CLF"
'            Select Case UCase$(slField)
'                Case "CLFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.lChfCode) '0
'                Case "CLFLINE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iLine) '4
'                Case "CLFCNTREVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iCntRevNo) '6
'                Case "CLFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iVefCode) '8
'                Case "CLFRPFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iRpfCode) '10
'                Case "CLFCXFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.lCxfCode) '77
'                Case "CLFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iUrfCode) '82
'                Case "CLFSCHSTATUS"
'                    gFieldOffset = 76
'                Case "CLFDELETE"
'                    gFieldOffset = 79
'                Case "CLFPROPVER"
'                    gFieldOffset = 0 'GetOffSetForInt(tmClf, tmClf.iPropVer) '6
'                Case "CLFTYPE"
'                    gFieldOffset = 118
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CMF"
'            Select Case UCase$(slField)
'                Case "CMFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCmf, tmCmf.lCode) '0
'                Case "CMFTITLE"
'                    gFieldOffset = 4
'                Case "CMFCOMTYPE"
'                    gFieldOffset = 24
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CNF"
'            Select Case UCase$(slField)
'                Case "CNFCRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCnf, tmCnf.lCrfCode) '0
'                Case "CNFINSTRNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCnf, tmCnf.iInstrNo) '4
'                Case "CNFCIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCnf, tmCnf.lCifCode) '6
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "COF"
'            Select Case UCase$(slField)
'                Case "COFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCof, tmCof.iYear) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CPF"
'            Select Case UCase$(slField)
'                Case "CPFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCpf, tmCpf.lCode) '0
'                Case "CPFNAME"
'                    gFieldOffset = 4
'                Case "CPFROTENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCpf, tmCpf.iRotEndDate(0)) '64
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CRF"
'            Select Case UCase$(slField)
'                Case "CRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.lCode) '0
'                Case "CRFETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iEtfCode) '4
'                Case "CRFENFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iEnfCode) '6
'                Case "CRFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iAdfCode) '8
'                Case "CRFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.lChfCode) '10
'                Case "CRFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iVefCode) '14
'                Case "CRFROTTYPE"
'                    gFieldOffset = 16
'                Case "CRFROTNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iRotNo) '16
'                Case "CRFSTARTDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iStartDate(0)) '18
'                Case "CRFENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iEndDate(0)) '22
'                Case "CRFLEN"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iLen) '41
'                Case "CRFANFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iAnfCode) '44
'                Case "CRFCSFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.lCsfCode) '44
'                Case "CRFFEEDSTATUS"
'                    gFieldOffset = 63
'                Case "CRFZONE"
'                    gFieldOffset = 68
'                Case "CRFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iUrfCode) '55
'                Case "CRFSIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.lSifCode) '44
'                Case "CRFAFFFDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCrf, tmCrf.iAffFdDate(0)) '18
'                Case "CRFAFFFDSTATUS"
'                    gFieldOffset = 111
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CSF"
'            Select Case UCase$(slField)
'                Case "CSFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCsf, tmCsf.lCode) '0
'                Case "CSFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCsf, tmCsf.iAdfCode) '4
'                Case "CSFTYPE"
'                    gFieldOffset = 6
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "CTF"
'        '    Select Case UCase$(slField)
'        '       Case "CTFYEAR"
'        '           gFieldOffset = 0 'GetOffSetForInt(tmCtf, tmCtf.iYear)
'        '       Case "CTFADFCODE"
'        '           gFieldOffset = 0 'GetOffSetForInt(tmCtf, tmCtf.iAdfCode)
'        '       Case "CTFAGFCODE"
'        '           gFieldOffset = 0 'GetOffSetForInt(tmCtf, tmCtf.iAgfCode)
'        '    End Select
'        Case "CXF"
'            Select Case UCase$(slField)
'                Case "CXFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCxf, tmCxf.lCode) '0
'                Case "CXFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCxf, tmCxf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "CYF"
'            Select Case UCase$(slField)
'                Case "CYFCIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCyf, tmCyf.lCifCode) '0
'                Case "CYFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmCyf, tmCyf.iVefCode) '4
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "DAF"
'        '    Select Case UCase$(slField)
'        '        Case "DAFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmDaf, tmDaf.iVefCode) '0
'        '        Case "DAFRDFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmDaf, tmDaf.iRdfCode) '0
'        '        Case "DAFSTARTDATE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmDaf, tmDaf.iStartDate(0)) '0
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "DLF"
'            Select Case UCase$(slField)
'                Case "DLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDlf, tmDlf.lCode) '0
'                Case "DLFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDlf, tmDlf.iVefCode) '0
'                Case "DLFETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDlf, tmDlf.iEtfCode) '4
'                Case "DLFENFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDlf, tmDlf.iEnfCode) '4
'                Case "DLFMNFFEED"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDlf, tmDlf.iMnfFeed) '4
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "DNF"
'            Select Case UCase$(slField)
'                Case "DNFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDnf, tmDnf.iCode) '0
'                Case "DNFBOOKNAME"
'                    gFieldOffset = 2
'                Case "DNFBOOKDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDnf, tmDnf.iBookDate(0)) '4
'                Case "DNFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDnf, tmDnf.iUrfCode) '4
'                Case "DNFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDnf, tmDnf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "DRF"
'            Select Case UCase$(slField)
'                Case "DRFDNFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDrf, tmDrf.iDnfCode) '0
'                Case "DRFDEMODATATYPE"
'                    gFieldOffset = 6
'                Case "DRFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDrf, tmDrf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "DSF"
'            Select Case UCase$(slField)
'                Case "DSFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDsf, tmDsf.lCode) '0
'                Case "DSFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmDsf, tmDsf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "ELF"
'        '    Select Case UCase$(slField)
'        '        Case "ELFDATE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmElf, tmElf.iDate(0)) '0
'        '        Case "ELFTIME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmElf, tmElf.iTime(0)) '4
'        '        Case "ELFURFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmElf, tmElf.iUrfCode) '8
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "ENF"
'            Select Case UCase$(slField)
'                Case "ENFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEnf, tmEnf.iCode) '0
'                Case "ENFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEnf, tmEnf.iVefCode) '2
'                Case "ENFETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEnf, tmEnf.iEtfCode) '4
'                Case "ENFNAME"
'                    gFieldOffset = 6
'                Case "ENFMNFGENRE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEnf, tmEnf.iMnfGenre) '36
'                Case "ENFCEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEnf, tmEnf.lCefCode) '47
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "ETF"
'            Select Case UCase$(slField)
'                Case "ETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmEtf, tmEtf.iCode) '0
'                Case "ETFTYPE"
'                    gFieldOffset = 2
'                Case "ETFNAME"
'                    gFieldOffset = 3
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "FSF"
'        '    Select Case UCase$(slField)
'        '        Case "FSFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.lCode) '0
'        '        Case "FSFFXFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.iFxfCode) '4
'        '        Case "FSFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.iVefCode) '6
'        '        Case "FSFMNFCOMP1"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.iMnfComp(0)) '50
'        '        Case "FSFMNFCOMP2"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.iMnfComp(1)) '52
'        '        Case "FSFURFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFsf, tmFsf.iUrfCode) '55
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        'Case "FXF"
'        '    Select Case UCase$(slField)
'        '        Case "FXFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFxf, tmFxf.iCode) '0
'        '        Case "FXFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFxf, tmFxf.iVefCode) '2
'        '        Case "FXFMNFFEED"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFxf, tmFxf.iMnfFeed) '4
'        '        Case "FXFFEEDDAY"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFxf, tmFxf.iFeedDay) '6
'        '        Case "FXFFEEDTIME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmFxf, tmFxf.iFeedTime(0)) '8
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        'Case "GMF"
'        '    Select Case UCase$(slField)
'        '        Case "GMFMNFGAME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmGmf, tmGmf.iMnfGame) '0
'        '        Case "GMFNUMBER"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmGmf, tmGmf.iNumber) '2
'        '        Case "GMFNAME"
'        '            gFieldOffset = 12
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "LCF"
'            Select Case UCase$(slField)
'                Case "LCFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLcf, tmLcf.iVefCode) '0
'                Case "LCFLOGDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLcf, tmLcf.iLogDate(0)) '2
'                Case "LCFSEQNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLcf, tmLcf.iSeqNo) '6
'                Case "LCFTYPE"
'                    gFieldOffset = 8
'                Case "LCFLVF1"      'Use to compute 1-50
'                    gFieldOffset = 0 'GetOffSetForInt(tmLcf, tmLcf.lLvfCode(1)) '16
'                Case "LCFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLcf, tmLcf.iUrfCode) '416
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "LEF"
'            Select Case UCase$(slField)
'                Case "LEFLVFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.lLvfCode) '0
'                Case "LEFSTARTTIME"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iStartTime(0)) '4
'                Case "LEFSEQNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iSeqNo) '8
'                Case "LEFETFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iEtfCode) '10
'                Case "LEFENFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iEnfCode) '12
'                Case "LEFMNFEXCL1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iMnfExcl(0)) '14
'                Case "LEFMNFEXCL2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iMnfExcl(1)) '16
'                Case "LEFANFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iAnfCode) '18
'                Case "LEFCIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.lEvtIDCefCode) '27
'                Case "LEFCEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.lCefCode) '32
'                Case "LEFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLef, tmLef.iUrfCode) '36
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
''        Case "LGF"
''            Select Case UCase$(slField)
''                Case "LGFVEFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iVefCode) '0
''                Case "LGFETFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iEtfCode) '2
''                Case "LGFENFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iEnfCode) '4
''                Case "LGFANFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iAnfCode) '2
''                Case "LGFADFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iAdfCode) '4
''                Case "LGFCIFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.lCifCode) '2
''                Case "LGFURFCODE"
''                    gFieldOffset = 0 'GetOffSetForInt(tmLgf, tmLgf.iUrfCode) '4
''                Case Else
''                    slMsg = "Offset to field " & slField & " missing from file " & slFile
''            End Select
'        'Case "LHF"
'        '    Select Case UCase$(slField)
'        '        Case "LHFDATE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmLhf, tmLhf.iDate(0)) '0
'        '        Case "LHFSTARTTIME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmLhf, tmLhf.iStartTime(0)) '4
'        '        Case "LHFSEQNO"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmLhf, tmLhf.iSeqNo) '8
'        '        Case "LHFENFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmLhf, tmLhf.iEnfCode) '10
'        '        Case "LHFANFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmLhf, tmLhf.iAnfCode) '12
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "LTF"
'            Select Case UCase$(slField)
'                Case "LTFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLtf, tmLtf.iCode) '0
'                Case "LTFTYPE"
'                    gFieldOffset = 2
'                Case "LTFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLtf, tmLtf.iVefCode) '4
'                Case "LTFNAME"
'                    gFieldOffset = 5
'                Case "LTFVAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLtf, tmLtf.iVar) '26
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "LVF"
'            Select Case UCase$(slField)
'                Case "LVFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLvf, tmLvf.lCode) '0
'                Case "LVFLTFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLvf, tmLvf.iLtfCode) '0
'                Case "LVFVERSION"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLvf, tmLvf.iVersion) '28
'                Case "LVFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmLvf, tmLvf.iUrfCode) '36
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "MCF"
'            Select Case UCase$(slField)
'                Case "MCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmMcf, tmMcf.iCode) '0
'                Case "MCFNAME"
'                    gFieldOffset = 2
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "MNF"
'            Select Case UCase$(slField)
'                Case "MNFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmMnf, tmMnf.iCode) '0
'                Case "MNFTYPE"
'                    gFieldOffset = 2
'                Case "MNFNAME"
'                    gFieldOffset = 3
'                Case "MNFUNITTYPE"
'                    gFieldOffset = 28
'                Case "MNFGROUPNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmMnf, tmMnf.iGroupNo) '40
'                Case "MNFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmMnf, tmMnf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "PHF"
'            Select Case UCase$(slField)
'                Case "PHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.lCode) '0
'                Case "PHFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iAgfCode) '0
'                Case "PHFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iAdfCode) '2
'                Case "PHFPRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.lPrfCode) '2
'                Case "PHFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iSlfCode) '4
'                Case "PHFCNTRNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.lCntrNo) '6
'                Case "PHFINVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.lInvNo) '10
'                Case "PHFAIRVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iAirVefCode) '18
'                Case "PHFTRANDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iTranDate(0)) '24
'                Case "PHFTRANTYPE"
'                    gFieldOffset = 36
'                Case "PHFAGEPERIOD"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iAgePeriod) '24
'                Case "PHFAGEYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iAgingYear) '24
'                Case "PHFCASHTRADE"
'                    gFieldOffset = 55
'                Case "PHFPURGEDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iPurgeDate(0)) '24
'                Case "PHFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iUrfCode) '49
'                Case "PHFBILLVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPhf, tmPhf.iBillVefCode) '18
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "PJF"
'            Select Case UCase$(slField)
'                Case "PJFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.lCode) '0
'                Case "PJFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iSlfCode) '0
'                Case "PJFSOFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iSofCode) '0
'                Case "PJFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iYear) '0
'                Case "PJFMNFBUS"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iMnfBus) '0
'                Case "PJFEFFDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iEffDate(0)) '0
'                Case "PJFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iAdfCode) '0
'                Case "PJFPRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.lPrfCode) '0
'                Case "PJFMNFDEMO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iMnfDemo) '0
'                Case "PJFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iVefCode) '0
'                Case "PJFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.lChfCode) '0
'                Case "PJFROLLOVERDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iRolloverDate(0)) '0
'                Case "PJFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPjf, tmPjf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "PNF"
'            Select Case UCase$(slField)
'                Case "PNFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPnf, tmPnf.iCode) '0
'                Case "PNFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPnf, tmPnf.iAdfCode) '0
'                Case "PNFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPnf, tmPnf.iAgfCode) '0
'                Case "PNFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPnf, tmPnf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "PRF"
'            Select Case UCase$(slField)
'                Case "PRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.lCode) '0
'                Case "PRFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iAdfCode) '0
'                Case "PRFNAME"
'                    gFieldOffset = 6
'                Case "PRFMNFCOMP1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iMnfComp(0)) '103
'                Case "PRFMNFCOMP2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iMnfComp(1)) '105
'                Case "PRFMNFEXCL1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iMnfExcl(0)) '107
'                Case "PRFMNFEXCL2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iMnfExcl(1)) '109
'                Case "PRFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iUrfCode) '22
'                Case "PRFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPrf, tmPrf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "PVF"
'            Select Case UCase$(slField)
'                Case "PVFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.lCode) '0
'                Case "PVFNAME"
'                    gFieldOffset = 4
'                Case "PVFVEFCODE1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(0)) '43
'                Case "PVFVEFCODE2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(1)) '43
'                Case "PVFVEFCODE3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(2)) '43
'                Case "PVFVEFCODE4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(3)) '43
'                Case "PVFVEFCODE5"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(4)) '43
'                Case "PVFVEFCODE6"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(5)) '43
'                Case "PVFVEFCODE7"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(6)) '43
'                Case "PVFVEFCODE8"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(7)) '43
'                Case "PVFVEFCODE9"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(8)) '43
'                Case "PVFVEFCODE10"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(9)) '43
'                Case "PVFVEFCODE11"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(10)) '43
'                Case "PVFVEFCODE12"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(11)) '43
'                Case "PVFVEFCODE13"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(12)) '43
'                Case "PVFVEFCODE14"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(13)) '43
'                Case "PVFVEFCODE15"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(14)) '43
'                Case "PVFVEFCODE16"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(15)) '43
'                Case "PVFVEFCODE17"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(16)) '43
'                Case "PVFVEFCODE18"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(17)) '43
'                Case "PVFVEFCODE19"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(18)) '43
'                Case "PVFVEFCODE20"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(19)) '43
'                Case "PVFVEFCODE21"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(20)) '43
'                Case "PVFVEFCODE22"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(21)) '43
'                Case "PVFVEFCODE23"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(22)) '43
'                Case "PVFVEFCODE24"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(23)) '43
'                Case "PVFVEFCODE25"
'                    gFieldOffset = 0 'GetOffSetForInt(tmPvf, tmPvf.iVefCode(24)) '43
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "RCF"
'            Select Case UCase$(slField)
'                Case "RCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iCode) '0
'                Case "RCFNAME"
'                    gFieldOffset = 2
'                Case "RCFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iVefCode) '6
'                Case "RCFSTARTDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iStartDate(0)) '8
'                Case "RCFENDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iEndDate(0)) '12
'                Case "RCFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iYear) '585
'                Case "RCFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iUrfCode) '585
'                Case "RCFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRcf, tmRcf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "RDF"
'            Select Case UCase$(slField)
'                Case "RDFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRdf, tmRdf.iCode) '0
'                Case "RDFNAME"
'                    gFieldOffset = 2
'                Case "RDFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRdf, tmRdf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "RGF"
'        '    Select Case UCase$(slField)
'        '        Case "RGFRPFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRgf, tmRgf.iRpfCode) '0
'        '        Case "RGFLEN"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRgf, tmRgf.iLen) '2
'        '        Case "RGFSPOTFREQ"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRgf, tmRgf.iSpotFreq) '4
'        '        Case "RGFWEEKFREQ"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRgf, tmRgf.iWeekFreq) '6
'        '        Case "RGFURFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRgf, tmRgf.iUrfCode) '68
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "RIF"
'            Select Case UCase$(slField)
'                Case "RIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.lCode) '0
'                Case "RIFRCFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.iRcfCode) '0
'                Case "RIFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.iVefCode) '0
'                Case "RIFRDFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.iRdfCode) '0
'                Case "RIFYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.iYear) '0
'                Case "RIFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRif, tmRif.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "RPF"
'        '    Select Case UCase$(slField)
'        '        Case "RPFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iCode) '0
'        '        Case "RPFRCFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iRcfCode) '2
'        '        Case "RPFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iVefCode) '4
'        '        Case "RPFNAME"
'        '            gFieldOffset = 6
'        '        Case "RPFLTFCODE1"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iLtfCode(0)) '28
'        '        Case "RPFLTFCODE2"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iLtfCode(1)) '32
'        '        Case "RPFLTFCODE3"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iLtfCode(2)) '36
'        '        Case "RPFSTARTTIME1"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iStartTime(0, 1))'40   'Used for all start times
'        '        Case "RPFENDTIME1"  'Used for all end times
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iEndTime(0, 1))'68
'        '        Case "RPFMO1"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.sWkDays(1, 1))'96   'Used for all days
'        '        Case "RPFANFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iAnfCode) '167
'        '        Case "RPFURFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmRpf, tmRpf.iUrfCode) '179
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "RVF"
'            Select Case UCase$(slField)
'                Case "RVFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.lCode) '0
'                Case "RVFAGFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iAgfCode) '0
'                Case "RVFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iAdfCode) '2
'                Case "RVFPRFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.lPrfCode) '2
'                Case "RVFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iSlfCode) '4
'                Case "RVFCNTRNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.lCntrNo) '6
'                Case "RVFINVNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.lInvNo) '10
'                Case "RVFAIRVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iAirVefCode) '18
'                Case "RVFTRANDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iTranDate(0)) '24
'                Case "RVFTRANTYPE"
'                    gFieldOffset = 36
'                Case "RVFAGEPERIOD"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iAgePeriod) '24
'                Case "RVFAGEYEAR"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iAgingYear) '24
'                Case "RVFCASHTRADE"
'                    gFieldOffset = 55
'                Case "RVFPURGEDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iPurgeDate(0)) '24
'                Case "RVFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iUrfCode) '49
'                Case "RVFBILLVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iBillVefCode) '18
'                Case "RVFPKLINENO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iPkLineNo) '18
'                Case "RVFINVDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmRvf, tmRvf.iInvDate(0)) '24
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SBF"
'            Select Case UCase$(slField)
'                Case "SBFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.lCode) '0
'                Case "SBFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.lChfCode) '0
'                Case "SBFDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.iDate(0)) '4
'                Case "SBFTRANTYPE"
'                    gFieldOffset = 12
'                Case "SBFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.iVefCode) '9
'                Case "SBFMNFITEM"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.iMnfItem) '11
'                Case "SBFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSbf, tmSbf.iUrfCode) '50
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SDF", "PSF"
'            Select Case UCase$(slField)
'                Case "SDFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.lCode) '0
'                Case "SDFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iVefCode) '0
'                Case "SDFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.lChfCode) '2
'                Case "SDFLINENO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iLineNo) '6
'                Case "SDFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iAdfCode) '8
'                Case "SDFDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iDate(0)) '10
'                Case "SDFTIME"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iTime(0)) '14
'                Case "SDFSCHSTATUS"
'                    gFieldOffset = 22
'                Case "SDFMNFMISSED"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iMnfMissed) '19
'                Case "SDFTRACER"
'                    gFieldOffset = 25
'                Case "SDFAFFCHG"
'                    gFieldOffset = 26
'                Case "SDFPOINTER"
'                    gFieldOffset = 27
'                Case "SDFCOPY"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.lCopyCode) '23
'                Case "SDFLEN"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iLen) '19
'                Case "SDFPRICETYPE"
'                    gFieldOffset = 34
'                Case "SDFSPOTTYPE"
'                    gFieldOffset = 35
'                Case "SDFBILL"
'                    gFieldOffset = 36
'                Case "SDFROTNO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iRotNo) '31
'                Case "SDFSMFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.lSmfCode) '31
'                Case "SDFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSdf, tmSdf.iUrfCode) '31
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "SFF"
'        '    Select Case UCase$(slField)
'        '        Case "SFFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSff, tmSff.iVefCode) '0
'        '        Case "SFFETFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSff, tmSff.iEtfCode) '2
'        '        Case "SFFENFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSff, tmSff.iEnfCode) '4
'        '        Case "SFFMNFFEED"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSff, tmSff.iMnfFeed) '4
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        'Case "SHF"
'        '    Select Case UCase$(slField)
'        '        Case "SHFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.lCode) '0
'        '        Case "SHFCHFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.lChfCode) '0
'        '        Case "SHFLINENO"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.iLineNo) '4
'        '        Case "SHFADFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.iAdfCode) '6
'        '        Case "SHFDATE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.iDate(0)) '8
'        '        Case "SHFTIME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.iTime(0)) '12
'        '        Case "SHFCPFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmShf, tmShf.lCpfCode) '17
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "SIF"
'            Select Case UCase$(slField)
'                Case "SIFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSif, tmSif.lCode) '0
'                Case "SIFADFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSif, tmSif.iAdfCode) '0
'                Case "SIFNAME"
'                    gFieldOffset = 6
'                Case "SIFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSif, tmSif.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SLF"
'            Select Case UCase$(slField)
'                Case "SLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSlf, tmSlf.iCode) '0
'                Case "SLFFIRSTNAME"
'                    gFieldOffset = 2
'                Case "SLFLASTNAME"
'                    gFieldOffset = 22
'                Case "SLFSOFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSlf, tmSlf.iSofCode) '42
'                Case "SLFJOBTITLE"
'                    gFieldOffset = 72
'                Case "SLFSTATE"
'                    gFieldOffset = 73
'                Case "SLFMNFSLSTEAM"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSlf, tmSlf.iMnfSlsTeam) '68
'                Case "SLFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSlf, tmSlf.iUrfCode) '459
'                Case "SLFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSlf, tmSlf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SMF"
'            Select Case UCase$(slField)
'                Case "SMFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.lCode) '0
'                Case "SMFSDFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.lSdfCode) '0
'                Case "SMFCHFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.lChfCode) '0
'                Case "SMFLINENO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.iLineNo) '4
'                Case "SMFMISSEDDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.iMissedDate(0)) '6
'                Case "SMFACTUALDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.iActualDate(0)) '6
'                Case "SMFMGSOURCE"
'                    gFieldOffset = 53
'                Case "SMFSCHSTATUS"
'                    gFieldOffset = 54
'                Case "SMFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSmf, tmSmf.iUrfCode) '43
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SOF"
'            Select Case UCase$(slField)
'                Case "SOFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSof, tmSof.iCode) '0
'                Case "SOFMNFSSCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSof, tmSof.iMnfSSCode) '2
'                Case "SOFMNFREGION"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSof, tmSof.iMnfRegion) '4
'                Case "SOFNAME"
'                    gFieldOffset = 6
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "SPF"
'            Select Case UCase$(slField)
'                Case "SPFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSpf, tmSpf.iCode) '0
'                Case "SPFURFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmSpf, tmSpf.iUrfGCode) '2
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "STF"
'            Select Case UCase$(slField)
'                Case "STFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmStf, tmStf.lCode) '0
'                Case "STFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmStf, tmStf.iVefCode) '0
'                Case "STFCREATEDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmStf, tmStf.iCreateDate(0)) '6
'                Case "STFLOGDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmStf, tmStf.iLogDate(0)) '6
'                Case "STFPRINT"
'                    gFieldOffset = 29
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        'Case "SVF"
'        '    Select Case UCase$(slField)
'        '        Case "SVFUSER"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.iUrfCode) '0
'        '        Case "SVFVEFCODE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.iVefCode) '0
'        '        Case "SVFWKSTARTDATE"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.iWkStartdate(0)) '8
'        '        Case "SVFSTARTTIME"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.iStartTime(0)) '12
'        '        Case "SVFLVFCODE1"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(1)) '4
'        '        Case "SVFLVFCODE2"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(2)) '4
'        '        Case "SVFLVFCODE3"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(3)) '4
'        '        Case "SVFLVFCODE4"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(4)) '4
'        '        Case "SVFLVFCODE5"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(5)) '4
'        '        Case "SVFLVFCODE6"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(6)) '4
'        '        Case "SVFLVFCODE7"
'        '            gFieldOffset = 0 'GetOffSetForInt(tmSvf, tmSvf.lLvfCode(7)) '4
'        '        Case Else
'        '            slMsg = "Offset to field " & slField & " missing from file " & slFile
'        '    End Select
'        Case "TZF"
'            Select Case UCase$(slField)
'                Case "TZFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCode) '0
'                Case "TZFCIFZONE1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(1)) '4
'                Case "TZFCIFZONE2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(2)) '4
'                Case "TZFCIFZONE3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(3)) '4
'                Case "TZFCIFZONE4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(4)) '4
'                Case "TZFCIFZONE5"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(5)) '4
'                Case "TZFCIFZONE6"
'                    gFieldOffset = 0 'GetOffSetForInt(tmTzf, tmTzf.lCifZone(6)) '4
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "URF"
'            Select Case UCase$(slField)
'                Case "URFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmUrf, tmUrf.iCode) '0
'                Case "URFNAME"
'                    gFieldOffset = 2
'                Case "URFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmUrf, tmUrf.iVefCode) '22
'                Case "URFREPT"
'                    gFieldOffset = 24
'                Case "URFSLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmUrf, tmUrf.iSlfCode) '44
'                Case "URFDEFAULTVEH"
'                    gFieldOffset = 0 'GetOffSetForInt(tmUrf, tmUrf.iDefVeh) '127
'                Case "URFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmUrf, tmUrf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "VCF"
'            Select Case UCase$(slField)
'                Case "VCFSELLCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVcf, tmVcf.iSellCode) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "VEF"
'            Select Case UCase$(slField)
'                Case "VEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iCode) '0
'                Case "VEFNAME"
'                    gFieldOffset = 2
'                Case "VEFTYPE"
'                    gFieldOffset = 167
'                Case "VEFMNFGROUP"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfGroup(1)) '0
'                Case "VEFMNFGP2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfVehGp2) '0
'                Case "VEFMNFGP3MKT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfVehGp3Mkt) '0
'                Case "VEFMNFGP4FMT"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfVehGp4Fmt) '0
'                Case "VEFMNFGP5RSCH"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfVehGp5Rsch) '0
'                Case "VEFMNFGP6SUB"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfVehGp6Sub) '0
'                Case "VEFMNFDEMO"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iMnfDemo) '0
'                Case "VEFVEFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iVefCode) '0
'                Case "VEFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVef, tmVef.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "VLF"
'            Select Case UCase$(slField)
'                Case "VLFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVlf, tmVlf.lCode) '0
'                Case "VLFSELLCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVlf, tmVlf.iSellCode) '0
'                Case "VLFSTATUS"
'                    gFieldOffset = 16
'                Case "VLFAIRCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVlf, tmVlf.iAirCode) '0
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "VPF"
'            Select Case UCase$(slField)
'                Case "VPFVEFKCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVpf, tmVpf.iVefKCode) '0
'                Case "VPFURFGCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVpf, tmVpf.iUrfGCode) '7
'                Case "VPFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVpf, tmVpf.iSyncDate(0)) '35
'                Case "VPFSYNCDATE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVpf, tmVpf.iSyncDate(0)) '35
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'        Case "VSF"
'            Select Case UCase$(slField)
'                Case "VSFCODE"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.lCode) '0
'                Case "VSFTYPE"
'                    gFieldOffset = 4
'                Case "VSFNAME"
'                    gFieldOffset = 5
'                Case "VSFFSCODE1"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(0)) '43
'                Case "VSFFSCODE2"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(1)) '45
'                Case "VSFFSCODE3"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(2)) '47
'                Case "VSFFSCODE4"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(3)) '49
'                Case "VSFFSCODE5"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(4)) '51
'                Case "VSFFSCODE6"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(5)) '53
'                Case "VSFFSCODE7"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(6)) '55
'                Case "VSFFSCODE8"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(7)) '57
'                Case "VSFFSCODE9"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(8)) '59
'                Case "VSFFSCODE10"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(9)) '61
'                Case "VSFFSCODE11"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(10)) '43
'                Case "VSFFSCODE12"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(11)) '45
'                Case "VSFFSCODE13"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(12)) '47
'                Case "VSFFSCODE14"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(13)) '49
'                Case "VSFFSCODE15"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(14)) '51
'                Case "VSFFSCODE16"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(15)) '53
'                Case "VSFFSCODE17"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(16)) '55
'                Case "VSFFSCODE18"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(17)) '57
'                Case "VSFFSCODE19"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(18)) '59
'                Case "VSFFSCODE20"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(19)) '61
'                Case "VSFFSCODE21"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(20)) '43
'                Case "VSFFSCODE22"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(21)) '45
'                Case "VSFFSCODE23"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(22)) '47
'                Case "VSFFSCODE24"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(23)) '49
'                Case "VSFFSCODE25"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(24)) '51
'                Case "VSFFSCODE26"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(25)) '53
'                Case "VSFFSCODE27"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(26)) '55
'                Case "VSFFSCODE28"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(27)) '57
'                Case "VSFFSCODE29"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(28)) '59
'                Case "VSFFSCODE30"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(29)) '61
'                Case "VSFFSCODE31"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(30)) '43
'                Case "VSFFSCODE32"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(31)) '45
'                Case "VSFFSCODE33"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(32)) '47
'                Case "VSFFSCODE34"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(33)) '49
'                Case "VSFFSCODE35"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(34)) '51
'                Case "VSFFSCODE36"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(35)) '53
'                Case "VSFFSCODE37"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(36)) '55
'                Case "VSFFSCODE38"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(37)) '57
'                Case "VSFFSCODE39"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(38)) '59
'                Case "VSFFSCODE40"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(39)) '61
'                Case "VSFFSCODE41"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(40)) '43
'                Case "VSFFSCODE42"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(41)) '45
'                Case "VSFFSCODE43"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(42)) '47
'                Case "VSFFSCODE44"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(43)) '49
'                Case "VSFFSCODE45"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(44)) '51
'                Case "VSFFSCODE46"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(45)) '53
'                Case "VSFFSCODE47"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(46)) '55
'                Case "VSFFSCODE48"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(47)) '57
'                Case "VSFFSCODE49"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(48)) '59
'                Case "VSFFSCODE50"
'                    gFieldOffset = 0 'GetOffSetForInt(tmVsf, tmVsf.iFSCode(49)) '61
'                Case "VSFSOURCE"
'                    gFieldOffset = 445
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'    End Select
    If slMsg = "" Then
        Exit Function
    End If
'    MsgBox slMsg, vbOkOnly + vbCritical + vbApplicationModal, "Offset Error"
    'If igBkgdProg = 0 Then
    '    MsgBox slMsg, vbOkOnly + vbCritical + vbApplicationModal, "Offset Error"
    'ElseIf igBkgdProg = 1 Then
    '    gLogMsg slMsg, "Bkgd_Schd.Txt", False
    'ElseIf igBkgdProg = 2 Then
    '    gLogMsg slMsg, "Set_Credit.Txt", False
    'Else
    '    gLogMsg slMsg, "TrafficErrors.Txt", False
    'End If
    gMsgBox slMsg, vbOKOnly + vbCritical + vbApplicationModal, "Offset Error"
    gFieldOffset = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveSmf                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove MG Spec for a spot      *
'*                                                     *
'*******************************************************
Function gFindSmf(tlSdf As SDF, hlSmf As Integer, tlSmf As SMF) As Integer
'
'   ilRet = gRemoveSmf(tlSdf, hlSmf, tlSmf)
'   Where:
'       tlSdf(I/O)- Sdf record
'       hlSmf(I)- Smf open handle
'       tlSmf(O) - Spot MG record if found
'
'
    Dim tlWSmf As SMF
    Dim ilRet As Integer
    Dim tlSmfSrchKey2 As LONGKEY0   'SdfCode
    Dim ilSmfRecLen As Integer

    tlSmf.lChfCode = 0
    tlSmf.iOrigSchVef = tlSdf.iVefCode
    tlSmf.sMGSource = "A"
    gFindSmf = False
    If (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        'tlSmfSrchKey.lChfCode = tlSdf.lChfCode
        'tlSmfSrchKey.iLineNo = tlSdf.iLineNo
        'tlSmfSrchKey.iMissedDate(0) = 0 'sch date =tlSdf.iDate(0)
        'tlSmfSrchKey.iMissedDate(1) = 0 'sch date =tlSdf.iDate(1)
        'ilSmfRecLen = Len(tlWSmf)
        'ilRet = btrGetGreaterOrEqual(hlSmf, tlWSmf, ilSmfRecLen, tlSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        'Do While (ilRet = BTRV_ERR_NONE) And (tlWSmf.lChfCode = tlSdf.lChfCode) And (tlWSmf.iLineNo = tlSdf.iLineNo)
        '    'If (tlWSmf.sSchStatus = tlSdf.sSchStatus) And (tlWSmf.iActualDate(0) = tlSdf.iDate(0)) And (tlWSmf.iActualDate(1) = tlSdf.iDate(1)) And (tlWSmf.iActualTime(0) = tlSdf.iTime(0)) And (tlWSmf.iActualTime(1) = tlSdf.iTime(1)) Then
        '    If tlWSmf.lSdfCode = tlSdf.lCode Then
        '        tlSmf = tlWSmf 'return SMF
        '        gFindSmf = True
        '        Exit Function
        '    End If
        '    ilRet = btrGetNext(hlSmf, tlWSmf, ilSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
        ilSmfRecLen = Len(tlWSmf)
        tlSmfSrchKey2.lCode = tlSdf.lCode
        ilRet = btrGetEqual(hlSmf, tlWSmf, ilSmfRecLen, tlSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tlSmf = tlWSmf 'return SMF
            gFindSmf = True
            Exit Function
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFindVehicleLatestDate         *
'*                                                     *
'*             Created:8/31/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find latest Vehicle date        *
'*                     (Last log date or Now)          *
'*                                                     *
'*******************************************************
Function gFindVehicleLatestDate(frm As Form, ilVefCode As Integer) As String
'
'   slDate = gFindVehicleLatestDate (MainForm, ilCode)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       ilCode (I)- Vehicle code number (combo allowed)
'       ilRet (O)- Index into tgVpf  for record associated with Vehicle
'                  Note: If Vehicle preference doesn't exist, it will be created
'

    Dim slDate As String
    Dim slTest As String
    Dim hlVsf As Integer
    'Dim tlVsf As VSF
    Dim ilRecLen As Integer     'Vsf record length
    Dim tlSrchKey As INTKEY0
    Dim ilLoop As Integer
    Dim ilDestroy As Integer
    Dim ilVpfIndex As Integer
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim blFirstCallToVpfFind As Boolean

    blFirstCallToVpfFind = True
    slDate = ""
    If ilVefCode > 0 Then
        If blFirstCallToVpfFind Then
            ilVpfIndex = gVpfFind(frm, ilVefCode)
            blFirstCallToVpfFind = False
        Else
            ilVpfIndex = gVpfFindIndex(ilVefCode)
        End If
        If ilVpfIndex = -1 Then
            slDate = Format$(gNow(), "m/d/yy")
        Else
            If tgVpf(ilVpfIndex).sGenLog = "N" Then
                slNowDate = Format$(gNow(), "m/d/yy")
                gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slDate
                If Not gValidDate(slDate) Then
                    slDate = slNowDate
                Else
                    If gDateValue(slNowDate) > gDateValue(slDate) Then
                        slDate = slNowDate
                    End If
                End If
            Else
                gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slDate
                If Not gValidDate(slDate) Then
                    slDate = Format$(gNow(), "m/d/yy")
                End If
            End If
        End If
    Else
        slNowDate = Format$(gNow(), "m/d/yy")
        hlVsf = CBtrvTable(ONEHANDLE)
        ilDestroy = True
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gFindLatestDateErr
        gBtrvErrorMsg ilRet, "gFindLatestDate (btrOpen)" & "Vsf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tmVsf)  'btrRecordLength(hlVpf)  'Get and save record length
        tlSrchKey.iCode = -ilVefCode
        ilRet = btrGetEqual(hlVsf, tmVsf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo gFindLatestDateErr
        gBtrvErrorMsg ilRet, "gFindLatestDate (btrGetEqual)", frm
        On Error GoTo 0
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
        ilDestroy = False
        For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
            If tmVsf.iFSCode(ilLoop) > 0 Then
                If blFirstCallToVpfFind Then
                    ilVpfIndex = gVpfFind(frm, tmVsf.iFSCode(ilLoop))
                    blFirstCallToVpfFind = False
                Else
                    ilVpfIndex = gVpfFindIndex(tmVsf.iFSCode(ilLoop))
                End If
                If ilVpfIndex >= 0 Then
                    If tgVpf(ilVpfIndex).sGenLog = "N" Then
                        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slTest
                        If Not gValidDate(slTest) Then
                            slTest = slNowDate
                        Else
                            If gDateValue(slNowDate) > gDateValue(slTest) Then
                                slTest = slNowDate
                            End If
                        End If
                    Else
                        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slTest
                    End If
                    If gValidDate(slTest) Then
                        If slDate = "" Then
                            slDate = Trim$(slTest)
                        Else
                            If gDateValue(slTest) > gDateValue(slDate) Then
                                slDate = slTest
                            End If
                        End If
                    End If
                End If
            End If
        Next ilLoop
        If slDate = "" Then
            slDate = Format$(gNow(), "m/d/yy")
        End If
    End If
    slNowDate = Format$(gNow(), "m/d/yy")
    If gDateValue(slNowDate) > gDateValue(slDate) Then
        gFindVehicleLatestDate = slNowDate
    Else
        If tgVpf(ilVpfIndex).sMoveLLD = "Y" Then  'User can alter contracts and move spots between Now and Last Log Date
            gFindVehicleLatestDate = slNowDate
        Else
            gFindVehicleLatestDate = slDate
        End If
    End If
    Exit Function
gFindLatestDateErr:
    On Error GoTo 0
    If ilDestroy Then
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
    End If
    gFindVehicleLatestDate = ""
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBtrvErrorMsg                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show error btrieve message      *
'*                                                     *
'*******************************************************
Function gGetBtrvErrorMsg(ilErrCode As Integer) As String
    Dim slErrMsg As String
    Dim ilRet As Integer

    slErrMsg = "Btrieve: "
    Select Case ilErrCode
        Case BTRV_ERR_INVALID_OP    '1
            slErrMsg = slErrMsg & "Invalid Operation (1)"
        Case BTRV_ERR_IO_ERR        '2
            slErrMsg = slErrMsg & "I/O Error (2)"
        Case BTRV_ERR_NOT_OPEN      '3
            slErrMsg = slErrMsg & "File not Opened (3)"
        Case BTRV_ERR_KEY_NOT_FOUND '4
            slErrMsg = slErrMsg & "Key Value not Found (4)"
        Case BTRV_ERR_DUPLICATE_KEY '5
            slErrMsg = slErrMsg & "Duplicate Key Value (5)"
        Case BTRV_ERR_INVALID_KEY   '6
            slErrMsg = slErrMsg & "Invalid Key Number (6)"
        Case BTRV_ERR_DIFF_KEY      '7
            slErrMsg = slErrMsg & "Different Key Number (7)"
        Case BTRV_ERR_INVALID_POS '8
            slErrMsg = slErrMsg & "Invalid Positioning (8)"
        Case BTRV_ERR_END_OF_FILE '9
            slErrMsg = slErrMsg & "End of File (9)"
        Case BTRV_ERR_MOD_KEY_VALUE '10
            slErrMsg = slErrMsg & "Modifiable Key Value Error (10)"
        Case BTRV_ERR_INVALID_FNAME '11
            slErrMsg = slErrMsg & "Invalid File Name (11)"
        Case BTRV_ERR_FILE_NOT_FOUND    '12
            slErrMsg = slErrMsg & "File Not Found (12)"
        Case BTRV_ERR_EXT_FILE      '13
            slErrMsg = slErrMsg & "Extended File Error (13)"
        Case BTRV_ERR_PREIMAGE_OPEN '14
            slErrMsg = slErrMsg & "Per-image Open Error (14)"
        Case BTRV_ERR_PREIMAGE_IO   '15
            slErrMsg = slErrMsg & "Pre-image I/O Error (15)"
        Case BTRV_ERR_EXPANSION '16
            slErrMsg = slErrMsg & "Expansion Error (16)"
        Case BTRV_ERR_CLOSE '17
            slErrMsg = slErrMsg & "Close Error (17)"
        Case BTRV_ERR_DISK_FULL '18
            slErrMsg = slErrMsg & "Disk Full (18)"
        Case BTRV_ERR_UNRECOVERABLE '19
            slErrMsg = slErrMsg & "Unrecoverable Error (19)"
        Case BTRV_ERR_MGR_INACTIVE  '20
            slErrMsg = slErrMsg & "Record Manager Inactive (20)"
        Case BTRV_ERR_KEYBUF_LENGTH '21
            slErrMsg = slErrMsg & "Key Buffer Too Short (21)"
        Case BTRV_ERR_DATABUF_LENGTH    '22
            slErrMsg = slErrMsg & "Data Buffer Length (22)"
        Case BTRV_ERR_POSBLK_LENGTH '23
            slErrMsg = slErrMsg & "Position Block Length (23)"
        Case BTRV_ERR_PAGESIZE  '24
            slErrMsg = slErrMsg & "Page Size Error (24)"
        Case BTRV_ERR_CREATE_IO '25
            slErrMsg = slErrMsg & "Create I/O Error"
        Case BTRV_ERR_NUMKEYS   '26
            slErrMsg = slErrMsg & "Number of Keys (26)"
        Case BTRV_ERR_INVALID_KEYPOS    '27
            slErrMsg = slErrMsg & "Invalid Key Position (27)"
        Case BTRV_ERR_REC_LENGTH    '28
            slErrMsg = slErrMsg & "Invalid Record Length (28)"
        Case BTRV_ERR_KEY_LENGTH    '29
            slErrMsg = slErrMsg & "Invalid Key Length (29)"
        Case BTRV_ERR_NOT_BTRV_FILE '30
            slErrMsg = slErrMsg & "Not a Btrieve File (30)"
        Case BTRV_ERR_ALREADY_EXTD  '31
            slErrMsg = slErrMsg & "File Already Extended (31)"
        Case BTRV_ERR_EXTD_IO   '32
            slErrMsg = slErrMsg & "Extend I/O Error (32)"
        Case BTRV_ERR_INVALID_EXT_NAME  '34
            slErrMsg = slErrMsg & "Invalid Extension Name (34)"
        Case BTRV_ERR_DIRECTORY '35
            slErrMsg = slErrMsg & "Directory Error (35)"
        Case BTRV_ERR_TRANSACTION   '36
            slErrMsg = slErrMsg & "Transaction Error"
        Case BTRV_ERR_TRANS_ACTIVE  '37
            slErrMsg = slErrMsg & "Transaction in Active (37)"
        Case BTRV_ERR_TRANS_FILE_IO '38
            slErrMsg = slErrMsg & "Transaction Control File I/O Error (38)"
        Case BTRV_ERR_END_ABORT_TRANS   '39
            slErrMsg = slErrMsg & "End/Abort Transaction Error (39)"
        Case BTRV_ERR_TRANS_MAX_FILES   '40
            slErrMsg = slErrMsg & "Transaction Max Files (40)"
        Case BTRV_ERR_OP_NOT_ALLOWED    '41
            slErrMsg = slErrMsg & "Operation Not allowed (41)"
        Case BTRV_ERR_ACCEL_ACCESS  '42
            slErrMsg = slErrMsg & "Incomplete Accelerated Access (42)"
        Case BTRV_ERR_INVALID_REC_ADDR  '43
            slErrMsg = slErrMsg & "Invalid Record Address (43)"
        Case BTRV_ERR_NULL_KEY_PATH '44
            slErrMsg = slErrMsg & "Null Key Path (44)"
        Case BTRV_ERR_INCON_KEY_FLAGS   '45
            slErrMsg = slErrMsg & "Inconsistent Key Flags (45)"
        Case BTRV_ERR_ACCESS_DENIED '46
            slErrMsg = slErrMsg & "Access to File Denied (46)"
        Case BTRV_ERR_MAX_OPEN_FILES    '47
            slErrMsg = slErrMsg & "Maximum Open Files (47)"
        Case BTRV_ERR_INVALID_ALT_SEQ   '48
            slErrMsg = slErrMsg & "Invalid Alternate Sequence Definition (48)"
        Case BTRV_ERR_KEY_TYPE  '49
            slErrMsg = slErrMsg & "Key Type Error"
        Case BTRV_ERR_OWNER_SET '50
            slErrMsg = slErrMsg & "Owner Already Set (50)"
        Case BTRV_ERR_INVALID_OWNER '51
            slErrMsg = slErrMsg & "Invalid Owner (51)"
        Case BTRV_ERR_WRITE_CACHE   '52
            slErrMsg = slErrMsg & "Error Writing Cache (52)"
        Case BTRV_ERR_INVALID_INTF  '53
            slErrMsg = slErrMsg & "Invalid Interface (53)"
        Case BTRV_ERR_VARIABLE_PAGE '54
            slErrMsg = slErrMsg & "Variable Page Error (54)"
        Case 55
            slErrMsg = slErrMsg & "Autoincrement Error (55)"
        Case BTRV_ERR_INCOMPLT_INDEX    '56
            slErrMsg = slErrMsg & "Incomplete Index"
        Case BTRV_ERR_EXPAND_MEM    '57
            slErrMsg = slErrMsg & "Expanded Memory Error (57)"
        Case BTRV_ERR_COMPBUF_SIZE  '58
            slErrMsg = slErrMsg & "Compression Buffer Too Short (58)"
        Case BTRV_ERR_FILE_EXISTS   '59
            slErrMsg = slErrMsg & "File Already Exists (59)"
        Case BTRV_ERR_REJECT_COUNT  '60
            slErrMsg = slErrMsg & "Reject Count Reached (60)"
        Case BTRV_ERR_WORK_SPACE_SIZE   '61
            slErrMsg = slErrMsg & "Work Space Too Small (61)"
        Case BTRV_ERR_INCORRECT_DESCP   '62
            slErrMsg = slErrMsg & "Incorrect Description (62)"
        Case BTRV_ERR_INVALID_EXTDBUF   '63
            slErrMsg = slErrMsg & "Invalid Extended Insert Buffer (63)"
        Case BTRV_ERR_FILTER_LIMIT  '64
            slErrMsg = slErrMsg & "Filter Limit Reached (64)"
        Case BTRV_ERR_INCOR_FLD_OFFSET  '65
            slErrMsg = slErrMsg & "Incorrect Field Offset (65)"
        Case BTRV_ERR_AUTO_TRANS_ABORT  '74
            slErrMsg = slErrMsg & "Automatic Transaction Abort (74)"
        Case BTRV_ERR_DEADLOCK_DETECT   '78
            slErrMsg = slErrMsg & "Deadlock Detected (78)"
        Case BTRV_ERR_CONFLICT  '80
            slErrMsg = slErrMsg & "Conflict Error (80)"
        Case BTRV_ERR_LOCK  '81
            slErrMsg = slErrMsg & "Lock Error (81)"
        Case BTRV_ERR_LOST_POS  '82
            slErrMsg = slErrMsg & "Lost Position (82)"
        Case BTRV_ERR_READ_TRANS    '83
            slErrMsg = slErrMsg & "Read Outside Transaction (83)"
        Case BTRV_ERR_REC_LOCKED    '84
            slErrMsg = slErrMsg & "File in Use [Re-press Save or Repeat Operation] (84)"
        Case BTRV_ERR_FILE_LOCKED   '85
            slErrMsg = slErrMsg & "File in Use [Re-press Save or Repeat Operation] (85)"
        Case BTRV_ERR_FILE_TBL_FULL '86
            slErrMsg = slErrMsg & "File Table Full (86)"
        Case BTRV_ERR_HNDL_TBL_FULL '87
            slErrMsg = slErrMsg & "Handle Table Full (87)"
        Case BTRV_ERR_INCOM_MODE    '88
            slErrMsg = slErrMsg & "Incompatible Mode Error (88)"
        Case BTRV_ERR_REDIR_DEV_FULL    '90
            slErrMsg = slErrMsg & "Redirection Device Table Full (90)"
        Case BTRV_ERR_SERVER    '91
            slErrMsg = slErrMsg & "Serve Error (91)"
        Case BTRV_ERR_TRANS_TBL_FULL    '92
            slErrMsg = slErrMsg & "Transaction Table Full (92)"
        Case BTRV_ERR_INCOM_LOCK    '93
            slErrMsg = slErrMsg & "Incompatible Lock Type (93)"
        Case BTRV_ERR_PERMISSION    '94
            slErrMsg = slErrMsg & "Permission Error (94)"
        Case BTRV_ERR_SESSION   '95
            slErrMsg = slErrMsg & "Session No Longer Valid (95)"
        Case BTRV_ERR_COMM_ENV  '96
            slErrMsg = slErrMsg & "Communications Environment Error (96)"
        Case BTRV_ERR_DATA_MSGSIZE  '97
            slErrMsg = slErrMsg & "Data Message too Small (97)"
        Case BTRV_ERR_INTERNAL_TRANS    '98
            slErrMsg = slErrMsg & "Internal Transaction Error (98)"
        'Case 30000      'Open error on dual system
        '    ilRet = csiHandleValue(0, 7)
        '    slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Save and Erase Prohibited:Error#=" & Str$(ilRet)
        'Case 30001      'Update or Insert or Delete error on dual system
        '    ilRet = csiHandleValue(0, 7)
        '    slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Update Prohibited:Error#=" & Str$(ilRet)
        'Case 30002      'Update or Insert or Delete error on dual system
        '    ilRet = csiHandleValue(0, 7)
        '    slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Insert Prohibited:Error#=" & Str$(ilRet)
        'Case 30003      'Update or Insert or Delete error on dual system
        '    ilRet = csiHandleValue(0, 7)
        '    slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Delete Prohibited:Error#=" & Str$(ilRet)
        Case 30000      'Open error on dual system
            ilRet = csiHandleValue(0, 7)
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Save and Erase Prohibited:Error#=" & str$(ilRet)
            Else
                slErrMsg = slErrMsg & "Call Counterpoint as System Error, Save and Erase Prohibited:Error#=" & str$(ilRet)
            End If
        Case 30001      'Update or Insert or Delete error on dual system
            ilRet = csiHandleValue(0, 7)
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Update Prohibited:Error#=" & str$(ilRet)
            Else
                slErrMsg = slErrMsg & "Call Counterpoint as System Error, Update Prohibited:Error#=" & str$(ilRet)
            End If
        Case 30002      'Update or Insert or Delete error on dual system
            ilRet = csiHandleValue(0, 7)
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Insert Prohibited:Error#=" & str$(ilRet)
            Else
                slErrMsg = slErrMsg & "Call Counterpoint as System Error, Insert Prohibited:Error#=" & str$(ilRet)
            End If
        Case 30003      'Update or Insert or Delete error on dual system
            ilRet = csiHandleValue(0, 7)
            If (Len(sgMDBPath) <> 0) And (Len(sgSDBPath) <> 0) Then
                slErrMsg = slErrMsg & "Call Counterpoint as Dual System Not Linked, Delete Prohibited:Error#=" & str$(ilRet)
            Else
                slErrMsg = slErrMsg & "Call Counterpoint as System Error, Delete Prohibited:Error#=" & str$(ilRet)
            End If
        Case Else
            slErrMsg = "Error #" & str$(ilErrCode)
    End Select
    gGetBtrvErrorMsg = slErrMsg
End Function



Function gGetByKeyForUpdateSSF(hlFile As Integer, tlSsf As SSF) As Integer
    Dim tlSsfSrchKey As SSFKEY0
    Dim tlSsfSrchKey2 As SSFKEY2
    Dim ilRecLen As Integer
    Dim ilRet As Integer
    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilVefCode As Integer

    If tlSsf.iType = 0 Then
        tlSsfSrchKey.iType = tlSsf.iType
        tlSsfSrchKey.iVefCode = tlSsf.iVefCode
        tlSsfSrchKey.iDate(0) = tlSsf.iDate(0)
        tlSsfSrchKey.iDate(1) = tlSsf.iDate(1)
        tlSsfSrchKey.iStartTime(0) = tlSsf.iStartTime(0)
        tlSsfSrchKey.iStartTime(1) = tlSsf.iStartTime(1)
        ilRecLen = Len(tlSsf)
        ilRet = gSSFGetEqual(hlFile, tlSsf, ilRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gGetByKeyForUpdateSSF = ilRet
        Exit Function
    Else
        ilVefCode = tlSsf.iVefCode
        ilType = tlSsf.iType
        ilDate0 = tlSsf.iDate(0)
        ilDate1 = tlSsf.iDate(1)
        tlSsfSrchKey2.iVefCode = ilVefCode
        tlSsfSrchKey2.iDate(0) = ilDate0
        tlSsfSrchKey2.iDate(1) = ilDate1
        ilRecLen = Len(tlSsf)
        ilRet = gSSFGetEqualKey2(hlFile, tlSsf, ilRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tlSsf.iVefCode = ilVefCode)
            If (tlSsf.iDate(0) = ilDate0) And (tlSsf.iDate(1) = ilDate1) Then
                If tlSsf.iType = ilType Then
                    gGetByKeyForUpdateSSF = ilRet
                    Exit Function
                End If
            End If
            ilRecLen = Len(tlSsf)
            ilRet = gSSFGetNext(hlFile, tlSsf, ilRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    gGetByKeyForUpdateSSF = BTRV_ERR_KEY_NOT_FOUND
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetCrfVefCode                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get rotation vehicle            *
'*                     If Hidden line get package line *
'*                     vehicle                         *
'*                                                     *
'*******************************************************
Function gGetCrfVefCode(hlClf As Integer, tlSdf As SDF, ilSchVefCode As Integer, ilPkgVefCode As Integer, ilLnVefCode As Integer, slLive As String, ilRdfCode As Integer) As Integer
    Dim ilLineNo As Integer
    Dim ilRet As Integer
    Dim tlClfSrchKey As CLFKEY0 'CLF key record image
    Dim ilClfRecLen As Integer  'CLF record length
    Dim tlClf As CLF            'CLF record image

    ilPkgVefCode = 0
    ilLnVefCode = 0
    slLive = "R"
    ilSchVefCode = tlSdf.iVefCode    'Use spot vehicle instead of line 9/22/98
    tlClfSrchKey.lChfCode = tlSdf.lChfCode
    tlClfSrchKey.iLine = tlSdf.iLineNo
    tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
    tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
    ilClfRecLen = Len(tlClf)
    ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = tlSdf.iLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F"))  'And (tlClf.sSchStatus = "A")
        ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = tlSdf.iLineNo) Then
        'ilLnVefCode = tlClf.iVefCode   'Remove 9/22/98- using spot vehicle and package vehicle
        ilLnVefCode = tlClf.iVefCode    '11/15/05- Added tp handle MG, Outsides and Fills
        ilRdfCode = tlClf.iRdfCode
        If (tlClf.sLiveCopy = "L") Or (tlClf.sLiveCopy = "M") Or (tlClf.sLiveCopy = "S") Or (tlClf.sLiveCopy = "P") Or (tlClf.sLiveCopy = "Q") Then
            slLive = tlClf.sLiveCopy
        End If
        If tlClf.sType = "H" Then
            ilLineNo = tlClf.iPkLineNo
            tlClfSrchKey.lChfCode = tlSdf.lChfCode
            tlClfSrchKey.iLine = ilLineNo
            tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
            tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
            ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = ilLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F"))   'And (tlClf.sSchStatus = "A")
                ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = ilLineNo) Then
                ilPkgVefCode = tlClf.iVefCode
            End If
        Else
            ilPkgVefCode = 0
        End If
        gGetCrfVefCode = True
    Else
        gGetCrfVefCode = False
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetEarliestLCFDate             *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine earliest date of LCF  *
'*                                                     *
'*******************************************************
Function gGetEarliestLCFDate(hlLcf As Integer, sLCP As String, ilVefCode As Integer) As Long
'
'   llDate = gGetEarliestLCFDate(hlLcf, slCP, ilVefCode)
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       slCP (I)- "C" = Current; "P" = Pending; "D" = Deleted
'       ilVefCode (I)- Vehicle code
'       llDate (O)- Earliest date on -1 if none
'
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim tlLcfSrchKey2 As LCFKEY2     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilType As Integer
    Dim ilVef As Integer

    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType <> "G" Then
            ilType = 0
            ilLcfRecLen = Len(tlLcf)
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = sLCP
            tlLcfSrchKey.iVefCode = ilVefCode
            tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tlLcfSrchKey.iLogDate(1) = 1900
            tlLcfSrchKey.iSeqNo = 1
            ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) Then
                gUnpackDate tlLcf.iLogDate(0), tlLcf.iLogDate(1), slDate
                gGetEarliestLCFDate = gDateValue(slDate)
            Else
                gGetEarliestLCFDate = -1
            End If
        Else
            ilLcfRecLen = Len(tlLcf)
            tlLcfSrchKey2.iVefCode = ilVefCode
            tlLcfSrchKey2.iLogDate(0) = 257  'Year 1/1/1900
            tlLcfSrchKey2.iLogDate(1) = 1900
            ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) Then
                gUnpackDate tlLcf.iLogDate(0), tlLcf.iLogDate(1), slDate
                gGetEarliestLCFDate = gDateValue(slDate)
            Else
                gGetEarliestLCFDate = -1
            End If
        End If
    Else
        gGetEarliestLCFDate = -1
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetFlightPrice                 *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Flight price                *
'*                                                     *
'*******************************************************
Function gGetFlightPrice(tlSdf As SDF, tlClf As CLF, hlCff As Integer, hlSmf As Integer, slPrice As String) As Integer
'
'   ilRet = gGetFlightPrice(tlSdf, hlCff, hlSmf, slPrice)
'   Where:
'       tlSdf(I)- Spot image to get flight for
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       hlSmf(I)- Open handle to SMF
'       slPrice(O)- Spot price
'       tgPriceCff(O)- Flight associated with sdf
'       ilRet = True if found; False if error
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slBonusOnInv As String
    'Dim tlCff As CFF

    If tlSdf.sSpotType = "X" Then
        slBonusOnInv = "Y"
        'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
            'If tlSdf.iAdfCode = tgCommAdf(ilLoop).iCode Then
            ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
            If ilLoop <> -1 Then
                slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
            '    Exit For
            End If
        'Next ilLoop
        'If tlSdf.sPriceType <> "N" Then
        If tlSdf.sPriceType = "+" Then
            slPrice = "+ Fill"  '"> Fill"
        ElseIf tlSdf.sPriceType = "-" Then
            slPrice = "- Fill"  '"< Fill"
        Else
            If slBonusOnInv <> "N" Then
                slPrice = "+ Fill"
            Else
                slPrice = "- Fill"
            End If
        End If
        gGetFlightPrice = True
        Exit Function
    ElseIf tlSdf.sSpotType = "O" Or tlSdf.sSpotType = "C" Then      '12-11-12
        slPrice = ".00"
        gGetFlightPrice = True
        Exit Function
    End If
    ilRet = gGetSpotFlight(tlSdf, tlClf, hlCff, hlSmf, tgPriceCff)  'tlCff)
    If ilRet = True Then
        Select Case tgPriceCff.sPriceType   'tlCff.sPriceType
            Case "T"    'True
                slPrice = gLongToStrDec(tgPriceCff.lActPrice, 2)    'tlCff.lActPrice, 2)
            Case "N"    'No Charge
                slPrice = "N/C"
            Case "M"    'MG Line
                slPrice = "MG"
            Case "B"    'Bonus
                slPrice = "Bonus"
            Case "S"    'Spinoff
                slPrice = "Spinoff"
            Case "P"    'Package
                slPrice = ".00"
            Case "R"    'Recapturable
                slPrice = "Recapturable"
            Case "A"    'ADU
                slPrice = "ADU"
        End Select
        gGetFlightPrice = True
    Else
        slPrice = ".00"
        gGetFlightPrice = False
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetLatestLCFDate               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine latest date of LCF    *
'*                                                     *
'*******************************************************
Function gGetLatestLCFDate(hlLcf As Integer, sLCP As String, ilVefCode As Integer) As Long
'
'   llDate = gGetLatestLCFDate(hlLcf, slCP, ilVefCode)
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       slCP (I)- "C" = Current; "P" = Pending; "D" = Deleted
'       ilVefCode (I)- Vehicle code
'       llDate (O)- Latest date on -1 if none
'
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim tlLcfSrchKey2 As LCFKEY2     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilType As Integer
    Dim ilVef As Integer

    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType <> "G" Then
            ilType = 0
            ilLcfRecLen = Len(tlLcf)
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = sLCP
            tlLcfSrchKey.iVefCode = ilVefCode
            tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
            tlLcfSrchKey.iLogDate(1) = 2100
            tlLcfSrchKey.iSeqNo = 1
            ilRet = btrGetLessOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) Then
                If tlLcf.iLogDate(1) > 0 Then   'Test if only TFN image
                    gUnpackDate tlLcf.iLogDate(0), tlLcf.iLogDate(1), slDate
                    gGetLatestLCFDate = gDateValue(slDate)
                Else
                    gGetLatestLCFDate = -1
                End If
            Else
                gGetLatestLCFDate = -1
            End If
        Else
            ilLcfRecLen = Len(tlLcf)
            tlLcfSrchKey2.iVefCode = ilVefCode
            tlLcfSrchKey2.iLogDate(0) = 257  'Year 1/1/1900
            tlLcfSrchKey2.iLogDate(1) = 2100
            ilRet = btrGetLessOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) Then
                If tlLcf.iLogDate(1) > 0 Then   'Test if only TFN image
                    gUnpackDate tlLcf.iLogDate(0), tlLcf.iLogDate(1), slDate
                    gGetLatestLCFDate = gDateValue(slDate)
                Else
                    gGetLatestLCFDate = -1
                End If
            Else
                gGetLatestLCFDate = -1
            End If
        End If
    Else
        gGetLatestLCFDate = -1
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetLatestVLfDate               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine latest date of Vlf    *
'*                                                     *
'*******************************************************
Function gGetLatestVLFDate(hlVlf As Integer, slType As String, ilVefCode As Integer, ilDay As Integer, ilPending As Integer) As Long 'VBC NR
'
'   llDate = gGetLatestVlfDate(hlVlf, slType, ilVefCode, ilDay As Integer)
'   Where:
'       hlVlf (I)- VLF handle (obtained from CBtrvTable)
'       slType(I)- S=Selling; A=Airing
'       ilVefCode (I)- Vehicle code
'       ilDay(I)- 0=M-F; 6=Sat; 7=Sun
'       ilPendinf(O)- True if pending; False if current
'       llDate (O)- Latest date on -1 if none
'
    Dim tlVlf As VLF                'VLF record image 'VBC NR
    Dim tlVlfSSrchKey As VLFKEY0     'VLF key record image 'VBC NR
    Dim tlVlfASrchKey As VLFKEY1     'VLF key record image 'VBC NR
    Dim ilVlfRecLen As Integer         'VLF record length 'VBC NR
    Dim slDate As String 'VBC NR
    Dim ilRet As Integer 'VBC NR

    ilVlfRecLen = Len(tlVlf) 'VBC NR
    ilPending = False 'VBC NR
    If slType = "S" Then 'VBC NR
        tlVlfSSrchKey.iSellCode = ilVefCode 'VBC NR
        tlVlfSSrchKey.iSellDay = ilDay 'VBC NR
        tlVlfSSrchKey.iEffDate(0) = 257  'Year 1/1/1900 'VBC NR
        tlVlfSSrchKey.iEffDate(1) = 2100 'VBC NR
        tlVlfSSrchKey.iSellTime(0) = 0 'VBC NR
        tlVlfSSrchKey.iSellTime(1) = 0 'VBC NR
        tlVlfSSrchKey.iSellPosNo = 0 'VBC NR
        tlVlfSSrchKey.iSellSeq = 0 'VBC NR
        ilRet = btrGetLessOrEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfSSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date 'VBC NR
        If (ilRet = BTRV_ERR_NONE) And (tlVlf.iSellCode = ilVefCode) And (tlVlf.iSellDay = ilDay) Then 'VBC NR
            If tlVlf.iEffDate(1) > 0 Then   'Test if only TFN image 'VBC NR
                If tlVlf.sStatus = "P" Then 'VBC NR
                    ilPending = True 'VBC NR
                End If 'VBC NR
                gUnpackDate tlVlf.iEffDate(0), tlVlf.iEffDate(1), slDate 'VBC NR
                gGetLatestVLFDate = gDateValue(slDate) 'VBC NR
            Else 'VBC NR
                gGetLatestVLFDate = -1 'VBC NR
            End If 'VBC NR
        Else 'VBC NR
            gGetLatestVLFDate = -1 'VBC NR
        End If 'VBC NR
    Else 'VBC NR
        tlVlfASrchKey.iAirCode = ilVefCode 'VBC NR
        tlVlfASrchKey.iAirDay = ilDay 'VBC NR
        tlVlfASrchKey.iEffDate(0) = 257  'Year 1/1/1900 'VBC NR
        tlVlfASrchKey.iEffDate(1) = 2100 'VBC NR
        tlVlfASrchKey.iAirTime(0) = 0 'VBC NR
        tlVlfASrchKey.iAirTime(1) = 0 'VBC NR
        tlVlfASrchKey.iAirPosNo = 0 'VBC NR
        tlVlfASrchKey.iAirSeq = 0 'VBC NR
        ilRet = btrGetLessOrEqual(hlVlf, tlVlf, ilVlfRecLen, tlVlfASrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date 'VBC NR
        If (ilRet = BTRV_ERR_NONE) And (tlVlf.iAirCode = ilVefCode) And (tlVlf.iAirDay = ilDay) Then 'VBC NR
            If tlVlf.iEffDate(1) > 0 Then   'Test if only TFN image 'VBC NR
                If tlVlf.sStatus = "P" Then 'VBC NR
                    ilPending = True 'VBC NR
                End If 'VBC NR
                gUnpackDate tlVlf.iEffDate(0), tlVlf.iEffDate(1), slDate 'VBC NR
                gGetLatestVLFDate = gDateValue(slDate) 'VBC NR
            Else 'VBC NR
                gGetLatestVLFDate = -1 'VBC NR
            End If 'VBC NR
        Else 'VBC NR
            gGetLatestVLFDate = -1 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    Exit Function 'VBC NR
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetProdOrShtTitle              *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get either Product Name or     *
'*                      Short Title                    *
'*                                                     *
'*******************************************************
Function gGetProdOrShtTitle(hlSif As Integer, llSifCode As Long, tlChf As CHF, tlAdf As ADF, Optional ilMaxAbbrChar = -1) As String
    Dim ilRet As Integer
    Dim tlSif As SIF
    Dim ilSifRecLen As Integer
    Dim tlSifSrchKey As LONGKEY0  'SIF key record image

    ilSifRecLen = Len(tlSif)
    If (tgSpf.sUseProdSptScr <> "P") Then
        If ilMaxAbbrChar = -1 Then
            gGetProdOrShtTitle = Trim$(tlAdf.sAbbr) & "," & Trim$(tlChf.sProduct)
        Else
            gGetProdOrShtTitle = Trim$(Left$(Trim$(tlAdf.sAbbr), ilMaxAbbrChar)) & "," & Trim$(tlChf.sProduct)
        End If
        Exit Function
    Else
        If llSifCode <= 0 Then
            If tlChf.lSifCode <= 0 Then
                gGetProdOrShtTitle = "" 'Trim$(tlAdf.sAbbr) & "," & Trim$(tlChf.sProduct)
                Exit Function
            Else
                tlSifSrchKey.lCode = tlChf.lSifCode
            End If
        Else
            tlSifSrchKey.lCode = llSifCode
        End If
        ilRet = btrGetEqual(hlSif, tlSif, ilSifRecLen, tlSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            gGetProdOrShtTitle = Trim$(tlSif.sName)
        Else
            gGetProdOrShtTitle = "" 'Trim$(tlAdf.sAbbr) & "," & Trim$(tlChf.sProduct)
        End If
    End If
    Exit Function

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetShortTitle                  *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get associated short title for  *
'*                     a spot                          *
'*                     Note: First check Rotation if   *
'*                           defined, if not use       *
'*                           one from contract         *
'*            Same code in gBuildODFSpotCode           *
'*                                                     *
'*******************************************************
Function gGetShortTitle(hlVsf As Integer, hlClf As Integer, hlSif As Integer, tlChf As CHF, tlAdf As ADF, tlSdf As SDF) As String
    'Copy rotation record information
    Dim tlVsfSrchKey0 As LONGKEY0 'VSF key record image
    Dim ilVsfReclen As Integer  'VSF record length
    Dim tlVsf As VSF            'VSF record image
    Dim ilSchVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim slLive As String
    Dim ilRet As Integer
    Dim ilVsf As Integer
    Dim llSifCode As Long
    Dim ilFound As Integer
    Dim ilRdfCode As Integer

    ilVsfReclen = Len(tlVsf)
    llSifCode = 0
    If tgSpf.sUseProdSptScr = "P" Then
        If tlChf.lVefCode < 0 Then
            ilFound = False
            tlVsfSrchKey0.lCode = -tlChf.lVefCode
            ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                ilRet = gGetCrfVefCode(hlClf, tlSdf, ilSchVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                For ilVsf = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                    If tlVsf.iFSCode(ilVsf) = ilSchVefCode Then
                        ilFound = True
                        If tlVsf.lFSComm(ilVsf) > 0 Then
                            llSifCode = tlVsf.lFSComm(ilVsf)
                        End If
                        Exit For
                    End If
                Next ilVsf
                If ilFound Then
                    Exit Do
                End If
                If tlVsf.lLkVsfCode <= 0 Then
                    Exit Do
                End If
                tlVsfSrchKey0.lCode = tlVsf.lLkVsfCode
                ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        gGetShortTitle = gGetProdOrShtTitle(hlSif, llSifCode, tlChf, tlAdf)
    Else
        gGetShortTitle = gGetProdOrShtTitle(hlSif, llSifCode, tlChf, tlAdf)
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetSpotFlight                  *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get spot flight                 *
'*                                                     *
'*******************************************************
Function gGetSpotFlight(tlSdf As SDF, tlClf As CLF, hlCff As Integer, hlSmf As Integer, tlCff As CFF) As Integer
'
'   ilRet = gGetSpotFlight(tlSdf, tlClf, hlCff, hlSmf, tlCff)
'   Where:
'       tlSdf(I)- Spot image to get flight for
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       hlSmf(I)- Open handle to SMF
'       tlCff(O)- Spot flight
'       ilRet = True if found; False if error
'
    Dim tlSmf As SMF
    Dim llSpotDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlCffSrchKey As CFFKEY0 'CFF key record image
    Dim ilCffRecLen As Integer     'CFF record length
    Dim ilRet As Integer

    gGetSpotFlight = False
    If (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        ilRet = gFindSmf(tlSdf, hlSmf, tlSmf)
        If ilRet = False Then
            Exit Function
        End If
        gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), llSpotDate
    Else
        gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llSpotDate
    End If
    tlCffSrchKey.lChfCode = tlSdf.lChfCode  'llChfCode
    tlCffSrchKey.iClfLine = tlClf.iLine 'tlSdf.iLineNo using line so avg price can be obtained for package line which bill by airing
    tlCffSrchKey.iCntRevNo = tlClf.iCntRevNo
    tlCffSrchKey.iPropVer = tlClf.iPropVer
    tlCffSrchKey.iStartDate(0) = 0
    tlCffSrchKey.iStartDate(1) = 0
    ilCffRecLen = Len(tlCff)
    ilRet = btrGetGreaterOrEqual(hlCff, tlCff, ilCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = tlSdf.lChfCode) And (tlCff.iClfLine = tlClf.iLine)
        If (tlCff.iCntRevNo = tlClf.iCntRevNo) And (tlCff.iPropVer = tlClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
            gUnpackDateLong tlCff.iStartDate(0), tlCff.iStartDate(1), llStartDate    'Week Start date
            gUnpackDateLong tlCff.iEndDate(0), tlCff.iEndDate(1), llEndDate    'Week Start date
            If (llSpotDate >= llStartDate) And (llSpotDate <= llEndDate) Then
                gGetSpotFlight = True
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hlCff, tlCff, ilCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetSpotPrice                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Spot price (this includes   *
'*                     correction for spot within a    *
'*                     conventional vehicle sold to a  *
'*                     virtual vehicle)                *
'*                                                     *
'*******************************************************
Function gGetSpotPrice(tlSdf As SDF, tlClf As CLF, hlCff As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer, slPrice As String) As Integer
'
'   ilRet = gGetSpotPrice(tlSdf, hlCff, hlSmf, hlVef, hlVsf, slPrice)
'   Where:
'       tlSdf(I)- Spot image to get flight for
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       hlSmf(I)- Open handle to SMF
'       hlVef(I)- Open handle to VEF
'       hlVsf(I)- Open handle to VSF
'       slPrice(O)- Spot price
'       Note: gGetFlightPrice sets tgPriceCff if not Extra or Fill
'       ilRet = True if found; False if error
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slBonusOnInv As String
    Dim tlSmf As SMF

    gGetSpotPrice = False
    If tlSdf.sSpotType = "X" Then
        slBonusOnInv = "Y"
        'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
        '    If tlSdf.iAdfCode = tgCommAdf(ilLoop).iCode Then
            ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
            If ilLoop <> -1 Then
                slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
        '        Exit For
            End If
        'Next ilLoop
        'If tlSdf.sPriceType <> "N" Then
        If tlSdf.sPriceType = "+" Then
            slPrice = "+ Fill"  '"> Fill"
        ElseIf tlSdf.sPriceType = "-" Then
            slPrice = "- Fill"  '"< Fill"
        Else
            If slBonusOnInv <> "N" Then
                slPrice = "+ Fill"
            Else
                slPrice = "- Fill"
            End If
        End If
        gGetSpotPrice = True
        Exit Function
    ElseIf tlSdf.sSpotType = "O" Or tlSdf.sSpotType = "C" Then          '12-11-12
        slPrice = ".00"
        gGetSpotPrice = True
        Exit Function
    End If
    ilRet = gGetFlightPrice(tlSdf, tlClf, hlCff, hlSmf, slPrice)
    If ilRet = True Then
        If (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
            ilRet = gFindSmf(tlSdf, hlSmf, tlSmf)
            If ilRet = False Then
                Exit Function
            End If
            slPrice = gVirtVefSpotPrice(hlVef, hlVsf, tlClf.iVefCode, tlSmf.iOrigSchVef, slPrice)
        Else
            slPrice = gVirtVefSpotPrice(hlVef, hlVsf, tlClf.iVefCode, tlSdf.iVefCode, slPrice)
        End If
        gGetSpotPrice = True
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetWkDollar                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Dollars for Week         *
'*                                                     *
'*******************************************************
Function gGetWkDollars(ilType As Integer, slInpDate As String, llWkDollars() As Long) As Long
'
'   llDollar = gGetWkDollars(ilType, slDate, llWkDollars())
'   Where:
'       ilType (I)- 0=Standard month; 1= Regular month; 4=Corp
'       slDate (I)- Date to obtain week dollars from buckets
'       llWkDollars(I)- Bucket dollars (0 to 53 week index values)
'       llDollar(O)- Week dollars
'
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    gObtainWkNo ilType, slInpDate, ilWkNo, ilFirstLastWk
    If (ilWkNo = 1) And ilFirstLastWk Then
        gGetWkDollars = llWkDollars(0) + llWkDollars(1)
    ElseIf (ilWkNo = 52) And ilFirstLastWk Then
        gGetWkDollars = llWkDollars(52) + llWkDollars(53)
    Else
        If (ilWkNo >= 0) And (ilWkNo <= 53) Then
            gGetWkDollars = llWkDollars(ilWkNo)
        Else
            gGetWkDollars = 0
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gHlfRead                        *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the screen help file       *
'*                                                     *
'*******************************************************
Sub gHlfRead(slIDName As String, tlHlf() As HLF)
'
'   gHlfRead slName, tlHlf()
'   Where:
'       slName (I)- Name of the help area to read
'       tlHlf (O)- where the user records are to be stored
'                   Note: tlHlf must be defined as Dim tlHlf() as HLF
'

    Dim ilRecLen As Integer     'HLF record length
    Dim tlSrchKey As HLFKEY0    'Hlf key record image
    Dim ilRet As Integer
    Dim ilUpperBound As Integer

    ReDim tlHlf(0 To 0) As HLF
    'Orlando meeting- ignore help messages
    Exit Sub
    ilRecLen = Len(tlHlf(0))  'btrRecordLength(hlHlf)  'Get and save record length
    ilUpperBound = UBound(tlHlf)
    tlSrchKey.sIDName = slIDName
    tlSrchKey.iMsgNo = 0
    tlHlf(ilUpperBound).sIDName = ""
    ilRet = btrGetGreaterOrEqual(hgHlf, tlHlf(ilUpperBound), ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (StrComp(Trim$(slIDName), Trim$(tlHlf(ilUpperBound).sIDName), 1) = 0) And (ilRet <> BTRV_ERR_END_OF_FILE)
        ilUpperBound = ilUpperBound + 1
        ReDim Preserve tlHlf(0 To ilUpperBound) As HLF
        tlHlf(ilUpperBound).sIDName = ""
        ilRet = btrGetNext(hgHlf, tlHlf(ilUpperBound), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If ilUpperBound > 0 Then
        ilUpperBound = ilUpperBound - 1
        ReDim Preserve tlHlf(0 To ilUpperBound) As HLF
    Else
        tlHlf(0).sIDName = ""
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gIFSCodeRefExist                *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified code     *
'*                     number exists in specified      *
'*                     combo file                      *
'*                                                     *
'*******************************************************
Function gIFSCodeRefExist(frm As Form, slFilter As String, ilCodeNo As Integer) As Integer
'
'   ilRet = gIFSCodeRefExist (MainForm, smFilter, iCodeNo)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       smFilter (I)- "V" (vehicle combo) or "S" (salesperson combo)
'       iCodeNo (I)- Code number to be checked
'       ilRet (O)- True = Found
'                  False = not found
'
    Dim hlVsf As Integer
    'Dim tlVsf As VSF
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer

    hlVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIFSCodeRefExistErr
    gBtrvErrorMsg ilRet, "gIFSCodeRefExist (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tmVsf) 'Get and save record length
    ilRet = btrGetFirst(hlVsf, tmVsf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        If StrComp(Trim$(tmVsf.sType), slFilter, 1) = 0 Then
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) = ilCodeNo Then
                    ilRet = btrClose(hlVsf)
                    btrDestroy hlVsf
                    gIFSCodeRefExist = True
                    Exit Function
                End If
            Next ilLoop
        End If
        ilRet = btrGetNext(hlVsf, tmVsf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo gIFSCodeRefExistErr
        gBtrvErrorMsg ilRet, "gIFSCodeRefExist (btrGetFirst or Next):" & "Vsf.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    gIFSCodeRefExist = False
    Exit Function
gIFSCodeRefExistErr:
    gDbg_HandleError "FileSubs: gIFSCodeRefExist"
'    gIFSCodeRefExist = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIICodeRefExist                 *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified code     *
'*                     number exists in specified file *
'*                                                     *
'*******************************************************
Function gIICodeRefExist(frm As Form, ilCodeNo As Integer, slFileName As String, slFieldName As String) As Integer
'
'   ilRet = gIICodeRefExist (MainForm, iCodeNo, sFileName, sFieldName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       iCodeNo (I)- Code number to be checked
'       sFileName (I)- File name
'       sFieldName (I)- field name of integer item in the record
'       ilRet (O)- True = Found
'                       False = not found
'

    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlTypeBuff As POPICODE   'Type field record
    Dim tlRecBuff As ARRAYREC
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilCodeOffset As Integer

    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slFile = Left$(slFileName, ilPos - 1)
    Else
        slFile = slFileName
    End If
    ilCodeOffset = gFieldOffset(slFile, slFieldName)
    If ilCodeOffset < 0 Then
        gIICodeRefExist = True
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIICodeRefExistErr
    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrOpen):" & slFileName, frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    If ilRecLen > Len(tlRecBuff) Then
        ilRecLen = Len(tlRecBuff)
    End If
    llNoRec = btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    'Remove getfirst because of variable length records blows up
    'ilRet = btrGetFirst(hlFile, tlRecBuff, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If ilRet = BTRV_ERR_END_OF_FILE Then
    '    ilRet = btrClose(hlFile)
    '    On Error GoTo gIICodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrReset):" & slFileName, Frm
    '    On Error GoTo 0
    '    btrDestroy hlFile
    '    gIICodeRefExist = False
    '    Exit Function
    'ElseIf ilRet <> BTRV_ERR_DATABUF_LENGTH Then    'Variable length records- return a 22 error- ignore
    '    On Error GoTo gIICodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrGetFirst):" & slFileName, Frm
    '    On Error GoTo 0
    'End If
'    ilCode = Asc(tlRecBuff.sChar(ilCodeOffset)) + 256 * Asc(tlRecBuff.sChar(ilCodeOffset + 1))'Rec definition starts with 2 bytes prior to the array
'    If ilCode = ilCodeNo Then
'        ilRet = btrClose(hlFile)
'        On Error GoTo gIICodeRefExistErr
'        gBtrvErrorMsg ilRet, "gIICodeRefExist (btrReset):" & slFileName, Frm
'        On Error GoTo 0
'        btrDestroy hlFile
'        gIICodeRefExist = True
'        Exit Function
'    End If
    llNoRec = gExtNoRec(ilRecLen)
    ilPos = InStr(slFileName, ".")
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", Left$(slFileName, ilPos - 1), "") 'Set extract limits (all records)
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTypeBuff, 2)
    On Error GoTo gIICodeRefExistErr
    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract iCode field
    On Error GoTo gIICodeRefExistErr
    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    If ddExtRecordCount(hlFile) > 0 Then
        gIICodeRefExist = True
    Else
        gIICodeRefExist = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gIICodeRefExistErr
    gBtrvErrorMsg ilRet, "gIICodeRefExist (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gIICodeRefExistErr:
    gDbg_HandleError "FileSubs: gIICodeRefExist"
'    gIICodeRefExist = True
End Function





'*******************************************************
'*                                                     *
'*      Procedure Name:gIsCntrRemote                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if Contract is        *
'*                     associated with Market Cluster  *
'*                     set to Yes                      *
'*                                                     *
'*******************************************************
Function gIsCntrRemote(tlChf As CHF, hlVsf As Integer) As Integer
'
'   Return:  False if No vehicle has Market Cluster = "Y";  True if vehicle associated with contract has Market Cluster = "Y"
'
    Dim ilVeh As Integer
    Dim ilVefCode As Integer
    Dim llLkVsfCode As Long
    Dim ilVsfReclen As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer

    gIsCntrRemote = False
    If (sgMktClusterDef = "N") Then
        Exit Function
    End If
    If (UBound(igMktVefCode) <= 0) Then
        Exit Function
    End If
    ilVsfReclen = Len(tmVsf)
    If tlChf.lVefCode > 0 Then
        For ilVeh = 0 To UBound(igMktVefCode) - 1 Step 1
            ilVefCode = igMktVefCode(ilVeh)
            If tlChf.lVefCode = ilVefCode Then
                gIsCntrRemote = True
                Exit Function
            End If
        Next ilVeh
    ElseIf tlChf.lVefCode < 0 Then
        llLkVsfCode = -tlChf.lVefCode
        Do While llLkVsfCode > 0
            tmLongSrchKey.lCode = llLkVsfCode
            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Function
            End If
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) > 0 Then
                    For ilVeh = 0 To UBound(igMktVefCode) - 1 Step 1
                        ilVefCode = igMktVefCode(ilVeh)
                        If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                            gIsCntrRemote = True
                            Exit Function
                        End If
                    Next ilVeh
                End If
            Next ilLoop
            llLkVsfCode = tmVsf.lLkVsfCode
        Loop
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLCFTFNExist               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if LCF exist for TFN  *
'*                                                     *
'*******************************************************
Function gLCFTFNExist(hlLcf As Integer, sLCP As String, ilVefCode As Integer) As Integer
'
'   ilRet = gLCFTFNExist(hlLcf, ilType, slCP, ilVefCode)
'   Where:
'       hlLcf (I)- LCF handle (obtained from CBtrvTable)
'       ilType (I)- 0=Regular Programming; 1-NN=Sports Programming (Game Number)
'       slCP (I)- "C" = Current; "P" = Pending; "D" = Deleted
'       ilVefCode (I)- Vehicle code
'       ilRet (O)- True if TFN exist; False if it dose not exist
'
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim ilRet As Integer
    Dim ilVefIndex As Integer
    Dim ilType As Integer

    ilVefIndex = gBinarySearchVef(ilVefCode)
    If ilVefIndex = -1 Then
        gLCFTFNExist = False
        Exit Function
    End If
    If tgMVef(ilVefIndex).sType <> "G" Then
        ilLcfRecLen = Len(tlLcf)
        ilType = 0
        tlLcfSrchKey.iType = ilType
        tlLcfSrchKey.sStatus = sLCP
        tlLcfSrchKey.iVefCode = ilVefCode
        tlLcfSrchKey.iLogDate(0) = 7  '1=Monday; 2=Tuesday;...7=Sunday
        tlLcfSrchKey.iLogDate(1) = 0
        tlLcfSrchKey.iSeqNo = 1
        ilRet = btrGetLessOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tlLcf.sStatus = sLCP) And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = ilType) Then
            If tlLcf.iLogDate(1) = 0 Then   'Test if TFN image
                gLCFTFNExist = True
            Else
                gLCFTFNExist = False
            End If
        Else
            gLCFTFNExist = False
        End If
    Else
        gLCFTFNExist = False
    End If
    Exit Function
End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:gLibCodeRefExist                *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified library  *
'*                     code number exists in library   *
'*                     calendar file                   *
'*                                                     *
'*******************************************************
Function gLibCodeRefExist(frm As Form, slFilter As String, llCodeNo As Long) As Integer


'   ilRet = gLibCodeRefExist (MainForm, smFilter, lCodeNo)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       smFilter (I)- "C" (current) or "P" (pending) or "B" (both)
'       lCodeNo (I)- Code number to be checked
'       ilRet (O)- True = Found
'                  False = not found
'

    Dim hlLcf As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim tlLcf As LCF   'file record
    Dim ilLoop As Integer

    hlLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gLibCodeRefExistErr
    gBtrvErrorMsg ilRet, "gLibCodeRefExist (btrOpen):" & "Lcf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlLcf) 'Get and save record length
    ilRet = btrGetFirst(hlLcf, tlLcf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        If (StrComp(Trim$(tlLcf.sStatus), slFilter, 1) = 0) Or (slFilter = "B") Then
            For ilLoop = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
                If tlLcf.lLvfCode(ilLoop) = llCodeNo Then
                    ilRet = btrClose(hlLcf)
                    btrDestroy hlLcf
                    gLibCodeRefExist = True
                    Exit Function
                End If
            Next ilLoop
        End If
        ilRet = btrGetNext(hlLcf, tlLcf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo gLibCodeRefExistErr
        gBtrvErrorMsg ilRet, "gLibCodeRefExist (btrGetFirst or Next):" & "Lcf.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    gLibCodeRefExist = False
    Exit Function
gLibCodeRefExistErr:
    gDbg_HandleError "FileSubs: gLibCodeRefExist"
'    gLibCodeRefExist = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLICodeRefExist                 *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified code     *
'*                     number exists in specified file *
'*                                                     *
'*                                                     *
'*******************************************************
Function gLICodeRefExist(frm As Form, ilCodeNo As Integer, slFileName As String, slFieldName As String) As Integer
'
'   ilRet = gLICodeRefExist (MainForm, iCodeNo, sFileName, sFieldName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       iCodeNo (I)- Code number to be checked
'       sFileName (I)- File name
'       sFieldName (I)- field name of long item in the record
'       ilRet (O)- True = Found
'                  False = not found
'

    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlTypeBuff As POPICODE   'Type field record
    Dim tlRecBuff As ARRAYREC
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilCodeOffset As Integer

    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slFile = Left$(slFileName, ilPos - 1)
    Else
        slFile = slFileName
    End If
    ilCodeOffset = gFieldOffset(slFile, slFieldName)
    If ilCodeOffset < 0 Then
        gLICodeRefExist = True
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gLICodeRefExistErr
    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrOpen):" & slFileName, frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    If ilRecLen > Len(tlRecBuff) Then
        ilRecLen = Len(tlRecBuff)
    End If
    llNoRec = btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    'Remove getfirst because of variable length records blows up
    'ilRet = btrGetFirst(hlFile, tlRecBuff, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If ilRet = BTRV_ERR_END_OF_FILE Then
    '    ilRet = btrClose(hlFile)
    '    On Error GoTo gLICodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrReset):" & slFileName, Frm
    '    On Error GoTo 0
    '    btrDestroy hlFile
    '    gLICodeRefExist = False
    '    Exit Function
    'ElseIf ilRet <> BTRV_ERR_DATABUF_LENGTH Then    'Variable length records- return a 22 error- ignore
    '    On Error GoTo gLICodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrGetFirst):" & slFileName, Frm
    '    On Error GoTo 0
    'End If
    'Offset adjusted for first bytes in the record prior to the array
'    ilCode = Asc(tlRecBuff.sChar(ilCodeOffset)) + 256 * Asc(tlRecBuff.sChar(ilCodeOffset + 1)) 'Rec definition starts with 2 bytes prior to the array
'    If ilCode = ilCodeNo Then
'        ilRet = btrClose(hlFile)
'        On Error GoTo gLICodeRefExistErr
'        gBtrvErrorMsg ilRet, "gLICodeRefExist (btrReset):" & slFileName, Frm
'        On Error GoTo 0
'        btrDestroy hlFile
'        gLICodeRefExist = True
'        Exit Function
'    End If
    llNoRec = gExtNoRec(ilRecLen)
    ilPos = InStr(slFileName, ".")
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", Left$(slFileName, ilPos - 1), "") 'Set extract limits (all records)
    tlTypeBuff.iCode = ilCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTypeBuff, 2)
    On Error GoTo gLICodeRefExistErr
    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract lCode field
    On Error GoTo gLICodeRefExistErr
    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    If ddExtRecordCount(hlFile) > 0 Then
        gLICodeRefExist = True
    Else
        gLICodeRefExist = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gLICodeRefExistErr
    gBtrvErrorMsg ilRet, "gLICodeRefExist (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gLICodeRefExistErr:
    gDbg_HandleError "FileSubs: gLICodeRefExist"
'    gLICodeRefExist = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLLCodeRefExist                 *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified long code*
'*                     number exists in specified file *
'*                                                     *
'*                                                     *
'*******************************************************
Function gLLCodeRefExist(frm As Form, llCodeNo As Long, slFileName As String, slFieldName As String) As Integer
'
'   ilRet = gLLCodeRefExist (MainForm, lCodeNo, sFileName, sFieldName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       lCodeNo (I)- Code number to be checked
'       sFileName (I)- File name
'       sFieldName (I)- field name of long item in the record
'       ilRet (O)- True = Found
'                  False = not found
'

    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlTypeBuff As POPLCODE   'Type field record
    Dim tlRecBuff As ARRAYREC
    Dim slFile As String
    Dim ilPos As Integer
    Dim ilCodeOffset As Integer

    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slFile = Left$(slFileName, ilPos - 1)
    Else
        slFile = slFileName
    End If
    ilCodeOffset = gFieldOffset(slFile, slFieldName)
    If ilCodeOffset < 0 Then
        gLLCodeRefExist = True
    End If
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gLLCodeRefExistErr
    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrOpen):" & slFileName, frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    If ilRecLen > Len(tlRecBuff) Then
        ilRecLen = Len(tlRecBuff)
    End If
    llNoRec = btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    'Remove getfirst because of variable length records blows up
    'ilRet = btrGetFirst(hlFile, tlRecBuff, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If ilRet = BTRV_ERR_END_OF_FILE Then
    '    ilRet = btrClose(hlFile)
    '    On Error GoTo gLLCodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrReset):" & slFileName, Frm
    '    On Error GoTo 0
    '    btrDestroy hlFile
    '    gLLCodeRefExist = False
    '    Exit Function
    'ElseIf ilRet <> BTRV_ERR_DATABUF_LENGTH Then    'Variable length records- return a 22 error- ignore
    '    On Error GoTo gLLCodeRefExistErr
    '    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrGetFirst):" & slFileName, Frm
    '    On Error GoTo 0
    'End If
    'Offset adjusted for first bytes in the record prior to the array
'    llCode = Asc(tlRecBuff.sChar(ilCodeOffset + 2)) + 256 * Asc(tlRecBuff.sChar(ilCodeOffset + 3))
'    llCode = 32768 * llCode + Asc(tlRecBuff.sChar(ilCodeOffset)) + 256 * Asc(tlRecBuff.sChar(ilCodeOffset + 1))
'    If llCode = llCodeNo Then
'        ilRet = btrClose(hlFile)
'        On Error GoTo gLLCodeRefExistErr
'        gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrReset):" & slFileName, Frm
'        On Error GoTo 0
'        btrDestroy hlFile
'        gLLCodeRefExist = True
'        Exit Function
'    End If
    llNoRec = gExtNoRec(ilRecLen)
    ilPos = InStr(slFileName, ".")
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", Left$(slFileName, ilPos - 1), "") 'Set extract limits (all records)
    tlTypeBuff.lCode = llCodeNo    'Extract all matching records
    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilCodeOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTypeBuff, 4)
    On Error GoTo gLLCodeRefExistErr
    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrExtAddLogicConst):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, 0, ilRecLen)  'Extract lCode field
    On Error GoTo gLLCodeRefExistErr
    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    If ddExtRecordCount(hlFile) > 0 Then
        gLLCodeRefExist = True
    Else
        gLLCodeRefExist = False
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo gLLCodeRefExistErr
    gBtrvErrorMsg ilRet, "gLLCodeRefExist (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
gLLCodeRefExistErr:
    gDbg_HandleError "FileSubs: gLLCodeRefExist"
'    gLLCodeRefExist = True
End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMissedReasonCode         *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the missed reason code   *
'*                     and store into igMnfMissed      *
'*                                                     *
'*******************************************************
Sub gObtainMissedReasonCode()
'
'   gObtainMissedReasonCode
'   Where:
'       igMnfMissed(O)- Missed reason code or zero is not found
'
    Dim hlMnf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlMnf As MNF
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slSyncDate As String
    Dim slSyncTime As String

    If (igMnfMissed > 0) And (igDefaultMnfMissed > 0) Then
        Exit Sub
    End If
    hlMnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Sub
    End If
    ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlMnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        tlMnf.iCode = 0
        tlMnf.sType = "M"
        tlMnf.sName = "Missed"
        tlMnf.sRPU = ""
        tlMnf.sUnitType = ""
        tlMnf.sSSComm = ""
        tlMnf.iMerge = 0
        tlMnf.iGroupNo = 0
        tlMnf.sCodeStn = ""
        tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
        tlMnf.iAutoCode = tlMnf.iCode
        ilRet = btrInsert(hlMnf, tlMnf, ilRecLen, INDEXKEY0)
        Do
            tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tlMnf.iAutoCode = tlMnf.iCode
            gGetSyncDateTime slSyncDate, slSyncTime
            gPackDate slSyncDate, tlMnf.iSyncDate(0), tlMnf.iSyncDate(1)
            gPackTime slSyncTime, tlMnf.iSyncTime(0), tlMnf.iSyncTime(1)
            ilRet = btrUpdate(hlMnf, tlMnf, ilRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        igMnfMissed = tlMnf.iCode
        igDefaultMnfMissed = igMnfMissed
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "M"
    ilOffSet = gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilRet = btrExtAddField(hlMnf, 0, ilRecLen)  'Extract record
    If ilRet <> BTRV_ERR_NONE Then
        tlMnf.iCode = 0
        tlMnf.sType = "M"
        tlMnf.sName = "Missed"
        tlMnf.sRPU = ""
        tlMnf.sUnitType = ""
        tlMnf.sSSComm = ""
        tlMnf.iMerge = 0
        tlMnf.iGroupNo = 0
        tlMnf.sCodeStn = ""
        tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
        tlMnf.iAutoCode = tlMnf.iCode
        ilRet = btrInsert(hlMnf, tlMnf, ilRecLen, INDEXKEY0)
        Do
            tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tlMnf.iAutoCode = tlMnf.iCode
            gGetSyncDateTime slSyncDate, slSyncTime
            gPackDate slSyncDate, tlMnf.iSyncDate(0), tlMnf.iSyncDate(1)
            gPackTime slSyncTime, tlMnf.iSyncTime(0), tlMnf.iSyncTime(1)
            ilRet = btrUpdate(hlMnf, tlMnf, ilRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        igMnfMissed = tlMnf.iCode
        igDefaultMnfMissed = igMnfMissed
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
    ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Sub
        End If
        'ilRet = btrExtGetFirst(hlMnf, tlMnf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim$(tlMnf.sUnitType) = "Y" Then
                igMnfMissed = tlMnf.iCode
                igDefaultMnfMissed = igMnfMissed
                Exit Do
            End If
            'Continue looking for default
            If StrComp("Missed", Trim$(tlMnf.sName), 1) = 0 Then
                igMnfMissed = tlMnf.iCode
                igDefaultMnfMissed = igMnfMissed
                'ilRet = btrClose(hlMnf)
                'btrDestroy hlMnf
                'Exit Sub
            End If
            ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
            Loop
        Loop
    End If
    If igMnfMissed > 0 Then
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Sub
    End If
    'Create missed reason as it is missing
    tlMnf.iCode = 0
    tlMnf.sType = "M"
    tlMnf.sName = "Missed"
    tlMnf.sRPU = ""
    tlMnf.sUnitType = ""
    tlMnf.sSSComm = ""
    tlMnf.iMerge = 0
    tlMnf.iGroupNo = 0
    tlMnf.sCodeStn = ""
    tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tlMnf.iAutoCode = tlMnf.iCode
    ilRet = btrInsert(hlMnf, tlMnf, ilRecLen, INDEXKEY0)
    Do
        tlMnf.iRemoteID = tgUrf(0).iRemoteUserID
        tlMnf.iAutoCode = tlMnf.iCode
        gGetSyncDateTime slSyncDate, slSyncTime
        gPackDate slSyncDate, tlMnf.iSyncDate(0), tlMnf.iSyncDate(1)
        gPackTime slSyncTime, tlMnf.iSyncTime(0), tlMnf.iSyncTime(1)
        ilRet = btrUpdate(hlMnf, tlMnf, ilRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    igMnfMissed = tlMnf.iCode
    igDefaultMnfMissed = igMnfMissed
    ilRet = btrClose(hlMnf)
    btrDestroy hlMnf
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPostLogAvailCode         *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Post Log Avail code  *
'*                                                     *
'*******************************************************
Sub gObtainPostLogAvailCode()
'
'   gObtainPostLogAvailCode
'   Where:
'       igPLAnfCode(O)- Missed reason code or zero is not found
'
    Dim hlAnf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlAnf As ANF
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer

    If igPLAnfCode > 0 Then
        Exit Sub
    End If
    hlAnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        Exit Sub
    End If
    ilRecLen = Len(tlAnf) 'btrRecordLength(hlMnf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlAnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlAnf, tlAnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        tlAnf.iCode = 0
        tlAnf.sName = "Post Log"
        tlAnf.sSustain = ""
        tlAnf.sSponsorship = ""
        tlAnf.iMerge = 0
        tlAnf.iRemoteID = tgUrf(0).iRemoteUserID
        tlAnf.iAutoCode = tlAnf.iCode
        ilRet = btrInsert(hlAnf, tlAnf, ilRecLen, INDEXKEY0)
        Do
            tlAnf.iRemoteID = tgUrf(0).iRemoteUserID
            tlAnf.iAutoCode = tlAnf.iCode
            'gGetSyncDateTime slSyncDate, slSyncTime
            'gPackDate slSyncDate, tlAnf.iSyncDate(0), tlAnf.iSyncDate(1)
            'gPackTime slSyncTime, tlAnf.iSyncTime(0), tlAnf.iSyncTime(1)
            ilRet = btrUpdate(hlAnf, tlAnf, ilRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        igPLAnfCode = tlAnf.iCode
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlAnf)
            btrDestroy hlAnf
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlAnf, llNoRec, -1, "UC", "ANF", "") 'Set extract limits (all records)
    ilRet = btrExtAddField(hlAnf, 0, ilRecLen)  'Extract record
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
    ilRet = btrExtGetNext(hlAnf, tlAnf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlAnf)
            btrDestroy hlAnf
            Exit Sub
        End If
        'ilRet = btrExtGetFirst(hlMnf, tlMnf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlAnf, tlAnf, ilRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If StrComp("Post Log", Trim$(tlAnf.sName), 1) = 0 Then
                igPLAnfCode = tlAnf.iCode
                ilRet = btrClose(hlAnf)
                btrDestroy hlAnf
                Exit Sub
            End If
            ilRet = btrExtGetNext(hlAnf, tlAnf, ilRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlAnf, tlAnf, ilRecLen, llRecPos)
            Loop
        Loop
    End If
    'Create Post Log avail record
    tlAnf.iCode = 0
    tlAnf.sName = "Post Log"
    tlAnf.sSustain = ""
    tlAnf.sSponsorship = ""
    tlAnf.iMerge = 0
    tlAnf.iRemoteID = tgUrf(0).iRemoteUserID
    tlAnf.iAutoCode = tlAnf.iCode
    ilRet = btrInsert(hlAnf, tlAnf, ilRecLen, INDEXKEY0)
    Do
        tlAnf.iRemoteID = tgUrf(0).iRemoteUserID
        tlAnf.iAutoCode = tlAnf.iCode
        'gGetSyncDateTime slSyncDate, slSyncTime
        'gPackDate slSyncDate, tlAnf.iSyncDate(0), tlAnf.iSyncDate(1)
        'gPackTime slSyncTime, tlAnf.iSyncTime(0), tlAnf.iSyncTime(1)
        ilRet = btrUpdate(hlAnf, tlAnf, ilRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    igPLAnfCode = tlAnf.iCode
    ilRet = btrClose(hlAnf)
    btrDestroy hlAnf
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSalespersonName          *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain salesperson name given   *
'*                     its code                        *
'*                                                     *
'*******************************************************
Sub gObtainSalespersonName(ilCode As Integer, slName, Optional blFirstLast = False)
'
'   gObtainSalespersonName ilCode, slName
'   Where:
'       ilCode (I)- Salesperson code #
'       slName (O) - name of the Salesperson (last name, first name) or combo
'

'    Dim tlSlf As SLF
'    'Dim tlVsf as VSF
'    Dim tlSlfSrchKey As INTKEY0    'key record image
'    Dim tlVsfSrchKey As LONGKEY0    'key record image
    Dim ilRet As Integer

    ilRet = gObtainSalesperson()
    If ilCode = 0 Then
        slName = ""
        Exit Sub
    End If
'    hlFile = CBtrvTable(ONEHANDLE)
'    If ilCode >= 0 Then
'        ilRecLen = Len(tlSlf) 'Get and save record length
'        tlSlfSrchKey.iCode = ilCode
'        ilRet = btrOpen(hlFile, "", sgDBPath & "Slf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    Else
'        ilRecLen = Len(tmVsf) 'Get and save record length
'        tlVsfSrchKey.lCode = -ilCode
'        ilRet = btrOpen(hlFile, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    End If
'    If ilRet <> BTRV_ERR_NONE Then
'        slName = ""
'        ilRet = btrClose(hlFile)
'        btrDestroy hlFile
'        Exit Sub
'    End If
'    If ilCode >= 0 Then
'        ilRet = btrGetEqual(hlFile, tlSlf, ilRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            slName = Trim$(tlSlf.sLastName) & ", " & Trim$(tlSlf.sFirstName)
'        Else
'            slName = ""
'        End If
'    Else
'        ilRet = btrGetEqual(hlFile, tmVsf, ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            slName = Trim$(tmVsf.sName)
'        Else
'            slName = ""
'        End If
'    End If
'    ilRet = btrClose(hlFile)
'    btrDestroy hlFile
    ilRet = gBinarySearchSlf(ilCode)
    If ilRet <> -1 Then
        If blFirstLast Then
            slName = Trim$(tgMSlf(ilRet).sFirstName) & " " & Trim$(tgMSlf(ilRet).sLastName)
        Else
            slName = Trim$(tgMSlf(ilRet).sLastName) & ", " & Trim$(tgMSlf(ilRet).sFirstName)
        End If
    Else
        slName = ""
    End If
    Exit Sub
End Sub






'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVehicleName             *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Vehicle name given its   *
'*                     code                            *
'*                                                     *
'*******************************************************
Sub gObtainVehicleName(ilCode As Integer, slName As String, slType As String)
'
'   gObtainVehicleName ilCode, slName, slType
'   Where:
'       ilCode (I)- Vehicle code #
'       slName (O) - name of the Vehicle
'       slType(O)- Vehicle type: N,S or A
'

    'Dim hlFile As Integer     'Vehicle or Combo Vehicle file handle
    'Dim tlVef As VEF
    'Dim tlVsf as VSF
    'Dim tlVefSrchKey As INTKEY0    'key record image
    'Dim tlVsfSrchKey As LONGKEY0    'key record image
    Dim ilRet As Integer
    Dim ilLoop As Integer

    slName = ""
    slType = ""
    ilRet = gObtainVef()
    If (ilRet = False) Or (ilCode = 0) Then
        Exit Sub
    End If
    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
    '    If ilCode = tgMVef(ilLoop).iCode Then
    ilLoop = gBinarySearchVef(ilCode)
    If ilLoop <> -1 Then
        slName = Trim$(tgMVef(ilLoop).sName)
        slType = Trim$(tgMVef(ilLoop).sType)
        Exit Sub
    End If
    '    End If
    'Next ilLoop
    Exit Sub
    'hlFile = CBtrvTable(ONEHANDLE)
    'If ilCode >= 0 Then
    '    ilRecLen = Len(tlVef) 'Get and save record length
    '    tlVefSrchKey.iCode = ilCode
    '    ilRet = btrOpen(hlFile, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'Else
    '    ilRecLen = Len(tmVsf) 'Get and save record length
    '    tlVsfSrchKey.iCode = -ilCode
    '    ilRet = btrOpen(hlFile, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'End If
    'If ilRet <> BTRV_ERR_NONE Then
    '    slName = ""
    '    slType = ""
    '    ilRet = btrClose(hlFile)
    '    btrDestroy hlFile
    '    Exit Sub
    'End If
    'If ilCode >= 0 Then
    '    ilRet = btrGetEqual(hlFile, tlVef, ilRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet = BTRV_ERR_NONE Then
    '        slName = Trim$(tlVef.sName)
    '        slType = Trim$(tlVef.sType)
    '    Else
    '        slName = ""
    '        slType = ""
    '    End If
    'Else
    '    ilRet = btrGetEqual(hlFile, tmVsf, ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet = BTRV_ERR_NONE Then
    '        slName = Trim$(tmVsf.sName)
    '        ilRet = btrClose(hlFile)
    '        ilRecLen = Len(tlVef) 'Get and save record length
    '        tlVefSrchKey.iCode = tmVsf.iFSCode(0)
    '        ilRet = btrOpen(hlFile, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    '        ilRet = btrGetEqual(hlFile, tlVef, ilRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '        If ilRet = BTRV_ERR_NONE Then
    '            slType = Trim$(tlVef.sType)
    '        Else
    '            slType = ""
    '        End If
    '    Else
    '        slName = ""
    '        slType = ""
    '    End If
    'End If
    'ilRet = btrClose(hlFile)
    'btrDestroy hlFile
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVirtVehList              *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Virtual Vehicle List     *
'*                                                     *
'*******************************************************
Sub gObtainVirtVehList()
'
'   gObtainVirtVehList
'   Where:
'       igVirtVefCode(O)- List of virtual vehicle codes
'

    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Vef.btr")

    'On Error GoTo gObtainVirtVehListErr2
    'ilRet = 0
    'imLowLimit = LBound(igVirtVefCode)
    'If ilRet <> 0 Then
    '    sgVVefStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(igVirtVefCode).Ptr <> 0 Then
        imLowLimit = LBound(igVirtVefCode)
    Else
        sgVVefStamp = ""
        imLowLimit = 0
    End If

    If sgVVefStamp <> "" Then
        If StrComp(slStamp, sgVVefStamp, 1) = 0 Then
            Exit Sub
        End If
    End If
    ReDim igVirtVefCode(0 To 0) As Integer
    sgVVefStamp = slStamp
    ilBound = 0
    'hlVef = CBtrvTable(ONEHANDLE)
    'ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    ilRet = btrClose(hlVef)
    '    btrDestroy hlVef
    '    Exit Sub
    'End If
    'ilVefRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
    'ilRet = btrGetFirst(hlVef, tlVef, ilVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Do While ilRet = BTRV_ERR_NONE
    '    If tlVef.sType = "V" Then
    '        igVirtVefCode(ilBound) = tlVef.iCode
    '        ilBound = ilBound + 1
    '        ReDim Preserve igVirtVefCode(0 To ilBound)
    '    End If
    '    ilRet = btrGetNext(hlVef, tlVef, ilVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Loop
    ilRet = gObtainVef()
    If ilRet = False Then
        Exit Sub
    End If
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If tgMVef(ilLoop).sType = "V" Then
            igVirtVefCode(ilBound) = tgMVef(ilLoop).iCode
            ilBound = ilBound + 1
            ReDim Preserve igVirtVefCode(0 To ilBound) As Integer
        End If
    Next ilLoop
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    Exit Sub
gObtainVirtVehListErr2:
    ilRet = 1
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gRCRead                         *
'*                                                     *
'*             Created:8/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the Rate Card file and     *
'*                     programs                        *
'*                                                     *
'*******************************************************
Sub gRCRead(frm As Form, ilCode As Integer, tlRcf As RCF, tlRif() As RIF, tlRdf() As RDF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   gRCRead MainForm, ilCode, tlRcf, tlRif(), tlRdf()
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       ilCode (I)- Rate card code #
'       tlRcf(O)- Rate card header
'       tlRif()(O)- Rate card items
'       tlRdf()(O)- Dayparts
'

    Dim ilRecLen As Integer     'SPF record length
    Dim hlFile As Integer        'site Option file handle
    Dim slStamp As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilUpperBound As Integer
    Dim llUpperBound As Long
    Dim llRif As Long
    Dim tlSrchKey As INTKEY0    'Rcf key record image

    slStamp = gFileDateTime(sgDBPath & "Rcf.btr") & gFileDateTime(sgDBPath & "Rif.btr") & gFileDateTime(sgDBPath & "Rdf.btr") & Trim$(str$(ilCode))

    'On Error GoTo gRCReadErr2
    'ilRet = 0
    'lmLowLimit = LBound(tlRif)
    'If ilRet <> 0 Then
    '    sgRCStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlRif).Ptr <> 0 Then
        imLowLimit = LBound(tlRif)
    Else
        sgRCStamp = ""
        imLowLimit = 0
    End If

    If sgRCStamp <> "" Then
        If slStamp = sgRCStamp Then
            Exit Sub
        End If
    End If
    sgRCStamp = slStamp
    ReDim tlRif(0 To 0) As RIF
    ReDim tlRdf(0 To 0) As RDF
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "Rcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRCReadErr
    gBtrvErrorMsg ilRet, "gRCRead (btrOpen):" & "Rcf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlRcf) 'btrRecordLength(hlSpf)  'Get and save record length
    tlSrchKey.iCode = ilCode
    ilRet = btrGetEqual(hlFile, tlRcf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo gRCReadErr
    gBtrvErrorMsg ilRet, "gRCRec (btrGetEqual)", frm
    On Error GoTo 0
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "Rif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRCReadErr
    gBtrvErrorMsg ilRet, "gRCRead (btrOpen):" & "Rif.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlRif(0))
    llUpperBound = UBound(tlRif)
    tlSrchKey.iCode = tlRcf.iCode
    'tlRpfSrchKey.iVefCode = -32000
    ilRet = btrGetGreaterOrEqual(hlFile, tlRif(llUpperBound), ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (tlRif(llUpperBound).iRcfCode = tlRcf.iCode) And (ilRet <> BTRV_ERR_END_OF_FILE)
        llUpperBound = llUpperBound + 1
        ReDim Preserve tlRif(0 To llUpperBound) As RIF
        ilRet = btrGetNext(hlFile, tlRif(llUpperBound), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo gRCReadErr
        gBtrvErrorMsg ilRet, "gRCRead (btrGetGreaterOrEqual or Next):" & "Rif.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlFile)
    btrDestroy hlFile

    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "Rdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRCReadErr
    gBtrvErrorMsg ilRet, "gRCRead (btrOpen):" & "Rdf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlRdf(0))
    ilUpperBound = UBound(tlRdf)
    For llRif = LBound(tlRif) To UBound(tlRif) - 1 Step 1
        ilFound = False
        For ilIndex = LBound(tlRdf) To UBound(tlRdf) - 1 Step 1
            If tlRif(llRif).iRdfCode = tlRdf(ilIndex).iCode Then
                ilFound = True
                Exit For
            End If
        Next ilIndex
        If Not ilFound Then
            tlSrchKey.iCode = tlRif(llRif).iRdfCode
            ilRet = btrGetEqual(hlFile, tlRdf(ilUpperBound), ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tlRdf(0 To ilUpperBound) As RDF
            End If
        End If
    Next llRif
    ilRet = btrClose(hlFile)
    btrDestroy hlFile

    Exit Sub
gRCReadErr:
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    gDbg_HandleError "FileSubs: gRCRead"
gRCReadErr2:
    ilRet = 1
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gRecExistForFile                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if any records exist within*
'*                     specified file                  *
'*                                                     *
'*******************************************************
Function gRecExistForFile(slFileName As String) As Integer
'
'  ilExist = gRecExistForFile (slFileName)
'   Where:
'        slFileName (I) - Name of file to be checked
'        ilExist (O) - True if any records exist in the specified file
'                           False if no records exist
'
    Dim hlFile As Integer        'File handle
    Dim llNoRec As Long          'Number of records
    Dim ilRet As Integer

    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        gRecExistForFile = False
        Exit Function
    End If
    llNoRec = btrRecords(hlFile) 'Obtain number of records
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    If llNoRec = 0 Then
        gRecExistForFile = False
    Else
        gRecExistForFile = True
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Function gRecLengthOk(slFileName As String, ilLen As Integer) As Integer
'
'   ilRecLen = gChkRecLength(slName, ilLen)
'   Where:
'       slName (I)- Name of the file
'       ilLen (I)- internal record length of file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilSize As Integer

    gRecLengthOk = True
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilCRet = btrClose(hlFile)
        btrDestroy hlFile
'        MsgBox slFileName & " open error: " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
        'If igBkgdProg = 0 Then
        '    MsgBox slFileName & " open error: " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
        'ElseIf igBkgdProg = 1 Then
        '    gLogMsg slFileName & " open error: " & Str$(ilRet), "Bkgd_Schd.Txt", False
        'ElseIf igBkgdProg = 2 Then
        '    gLogMsg slFileName & " open error: " & Str$(ilRet), "Set_Credit.Txt", False
        'Else
        '    gLogMsg slFileName & " open error: " & Str$(ilRet), "TrafficErrors.Txt", False
        'End If
        gMsgBox slFileName & " open error: " & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        gRecLengthOk = False
        Exit Function
    End If
    ilSize = btrRecordLength(hlFile)
    If ilSize <> ilLen Then
        ilCRet = btrClose(hlFile)
        btrDestroy hlFile
'        MsgBox slFileName & " size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(ilLen), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
        'If igBkgdProg = 0 Then
        '    MsgBox slFileName & " size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(ilLen), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
        'ElseIf igBkgdProg = 1 Then
        '    gLogMsg slFileName & " size error: Btrieve Size" & Str$(ilSize), "Bkgd_Schd.Txt", False
        'ElseIf igBkgdProg = 2 Then
        '    gLogMsg slFileName & " size error: Btrieve Size" & Str$(ilSize), "Set_Credit.Txt", False
        'Else
        '    gLogMsg slFileName & " size error: Btrieve Size" & Str$(ilSize), "TrafficErrors.Txt", False
        'End If
        gMsgBox slFileName & " size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(ilLen), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        gRecLengthOk = False
        Exit Function
    End If
    ilCRet = btrClose(hlFile)
    btrDestroy hlFile
    gRecLengthOk = True
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:gSpfRead                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the site Option file       *
'*                                                     *
'*******************************************************
Sub gSpfRead()
'
'  gSpfRead
'  Where:
'

    Dim ilRecLen As Integer     'SPF record length
    Dim hlSpf As Integer        'site Option file handle
    Dim slStamp As String
    Dim llNoRec As Long     'Number of records (must be one)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slGetStamp As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim hlSpfx As Integer        'TTP 10205 - 6/21/21 - JW -  Site Option extension file handle
    
    slStamp = gFileDateTime(sgDBPath & "Spf.btr")
    If sgSpfStamp <> "" Then
        If slStamp = sgSpfStamp Then
            Exit Sub
        End If
        slGetStamp = ""
    Else
        slGetStamp = "" 'gGetCSIStamp("SPF")
    End If
    If slGetStamp <> "" Then
'        sgSpfStamp = slGetStamp
'        ilRet = csiGetRec("SPF", 0, VarPtr(tgSpf), LenB(tgSpf))
        igSiteSet = True
    Else
        sgSpfStamp = slStamp
        hlSpf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gSpfReadErr
        gBtrvErrorMsg ilRet, "gSpfRead (btrOpen):" & "Spf.Btr", Traffic
        On Error GoTo 0
        ilRecLen = Len(tgSpf) 'btrRecordLength(hlSpf)  'Get and save record length
        llNoRec = btrRecords(hlSpf) 'Obtain number of records
        If llNoRec = 0 Then
            btrDestroy hlSpf
            hlSpf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrOpen):" & "Spf.Btr", Traffic
            On Error GoTo 0
            igSiteSet = False
            'General
            tgSpf.sGClient = ""
            For ilLoop = LBound(tgSpf.sGAddr) To UBound(tgSpf.sGAddr) Step 1
                tgSpf.sGAddr(ilLoop) = ""
            Next ilLoop
            tgSpf.sSystemType = "N"
            tgSpf.sGTBar = "U"
            'tgSpf.iGRetainCntr = 0
            tgSpf.sUsingFeatures3 = Chr(0)
            tgSpf.sUsingFeatures4 = Chr(0)
            tgSpf.sUsingFeatures5 = Chr(0)
            tgSpf.sUsingFeatures6 = Chr(0)
            tgSpf.sUsingFeatures7 = Chr(0)
            tgSpf.sUsingFeatures8 = Chr(0)
            tgSpf.sUsingFeatures9 = Chr(0)
            tgSpf.sUsingFeatures10 = Chr(0)
            tgSpf.iRetainAffSpot = 24
            tgSpf.iRetainTrafSpot = 24
            tgSpf.iRetainTrafProp = 12
            'tgSpf.iGRetainCondensed = 21
            'tgSpf.iGRetainPayment = 120
            tgSpf.iMnfClientAbbr = 0
            tgSpf.iMnfInvTerms = 0      'Default terms (if zero use "15 days upon receipt"
            tgSpf.sAutoType = Chr(0)
            tgSpf.sAutoType2 = Chr(0)    '8-10-05 more automation types
            tgSpf.sAutoType3 = Chr(0)    '2/17/11 more automation types
            'Sales
            tgSpf.sSEnterAgeDate = "E"      'for billed & booked, default to use entered date
            tgSpf.sSAudData = "H"
            tgSpf.sSGRPCPPCal = "A"
            tgSpf.sSSellNet = "N"
            tgSpf.sSDelNet = "N"
            'tgSpf.iSDay = 5         'Not Used
            tgSpf.sSBrStdQt = "N"
            'tgSpf.sSUseResv = "N"   'Not Used
            tgSpf.sHideGhostSptScr = "N"
            tgSpf.sUsingFeatures2 = Chr(0)
            tgSpf.sUsingBBs = "N"
            tgSpf.sBBsToAff = "N"
            tgSpf.sTapeShowForm = "A"
            tgSpf.sOptionFields = Chr(&HFF)
            tgSpf.sOverrideOptions = Chr(0)
            tgSpf.sRepRptForm = "M"     'Market
            tgSpf.sAllowDailyBuys = "N" 'Allow daily buys
            tgSpf.sPostCalAff = "S"     'Standard month
            tgSpf.sBActDayCompl = "Y"
            tgSpf.lCxfDemoEst = 0
            'tgSpf.sSUseRem = "N"    'Not Used
            'tgSpf.sSUseDR = "N"     'Not Used
            'tgSpf.sSUsePI = "N"     'Not Used
            'tgSpf.sSUsePSA = "N"    'Not Used
            tgSpf.lBCxfDisclaimer = 0
            tgSpf.lCxfContrComment = 0         '2-12-03
            tgSpf.lCxfInsertComment = 0          '2-12-03
            'tgSpf.sSUsePromo = "N"  'Not Used
            tgSpf.sBMissedDT = "N"
            tgSpf.sGUsePropSys = "N"
            tgSpf.sGUseAffSys = "N"
            tgSpf.sGUseAffFeed = "N"
            tgSpf.iSLen(0) = 5      'Spot lengths
            tgSpf.iSLen(1) = 10
            tgSpf.iSLen(2) = 15
            tgSpf.iSLen(3) = 20
            tgSpf.iSLen(4) = 30
            tgSpf.iSLen(5) = 45
            tgSpf.iSLen(6) = 60
            tgSpf.iSLen(7) = 90
            tgSpf.iSLen(8) = 120
            tgSpf.iSLen(9) = 0
            tgSpf.iNROBlockDate(0) = 0
            tgSpf.iNROBlockDate(1) = 0
            tgSpf.iNROBlockTime(0) = 0
            tgSpf.iNROBlockTime(1) = 0
            tgSpf.sAllowMGs = "Y"    'Allow MG's
            tgSpf.sSSRetainDate = "Y"    'Retain Date
            tgSpf.sSubCompany = "N"    'Allow Sub-Company Vehicle Groups
            tgSpf.sCommByCntr = "N"    'Allow Salesperson commission to vary by Contract
            tgSpf.sMktBase = "N"    'Market base
            tgSpf.sPLMove = "A"    'Post Log Moves (Ask)
            'Agency/Advertiser
            'tgSpf.iAProd = 20   'Product name- set to max
            tgSpf.sAExcl = "N"  'Program exclusions
            tgSpf.iATargets = 2 'Number of demos
            tgSpf.sAISCI = "A"  'Shwo ISCI- Always
            tgSpf.sAEDIC = "N"  'EDI service for contracts
            tgSpf.sAEDII = "N"  'EDI service for invoices
            tgSpf.sAPrtStyle = "W"  'Wide
            tgSpf.sARepCodes = "N"    'Using Rep codes-No
            tgSpf.sAStnCodes = "N"    'Using Station codes-No
            tgSpf.sAAgyCodes = "N"    'Using Agency codes-No
            'Contract/Copy
            tgSpf.sCBumpPast = "N"   'Bump spots in past
            tgSpf.lCLowestNo = 1 'Lowest contract #
            tgSpf.lCHighestNo = 99999999   'Highest contract #- Not using
            tgSpf.lCNextNo = 1 'Next number
            tgSpf.sCEstNo = "N"
            'tgSpf.sCRot = "V"   'Unused-'Copy rotation-Varies with instruction
            tgSpf.sCBlackoutLog = "N"   'Using Blackout on Logs (Y/N)
            tgSpf.sInvAirOrder = "A"
            tgSpf.sCPkOrdered = "N"
            tgSpf.sCPkAired = "N"
            tgSpf.sCPkEqual = "N"
            tgSpf.iPkageGenMeth = 0
            tgSpf.sCWarnMsg = "N"
            tgSpf.iReallMnfDemo = 0
            tgSpf.iReallDate(0) = 0
            tgSpf.iReallDate(1) = 0
            tgSpf.sAllowPrelLog = "N"   'Allow Preliminary Logs
            'tgSpf.sLkLog = ""
            tgSpf.sSportInfo = Chr(0)
            'tgSpf.iLkLogUrfCode = 0
            tgSpf.iReconcGroupNo = 0    '1-20-04 replaced tgspf.ilklogurfcode
            tgSpf.iRptDollarMag = 0
            tgSpf.sUnitOr3060 = "3"
            tgSpf.sCIncludeMissDB = "N"
            tgSpf.sSchdRemnant = "N"    'Schedule Remnant contracts like Standard contracts
            tgSpf.sSchdPromo = "N"    'Schedule Promo contracts like Standard contracts
            tgSpf.sSchdPSA = "N"    'Schedule PSA contracts like Standard contracts
            tgSpf.sCLnStdQt = "N"
            tgSpf.sCUseSegments = "N"
            tgSpf.sUsingNTR = "N"
            tgSpf.sUsingFeatures = Chr(0)
            tgSpf.sAvailEqualize = "N"
            tgSpf.sDemoEstAllowed = "N"
            tgSpf.sUsingTraffic = "Y"
            tgSpf.sImptCntr = "P"
            tgSpf.sInsertAddr = "P"
            'Billing
            tgSpf.sBLCycle = "S"    'Local billing cycle-Standard
            tgSpf.sBRCycle = "S"    'Regional billing cycle-Standard
            tgSpf.sBNCycle = "S"    'National billing cycle-Standard
            tgSpf.lBLowestNo = 1 'Lowest invoice #
            tgSpf.lBHighestNo = 99999999   'Highest invoice #
            tgSpf.lBNextNo = 1 'Next number
            tgSpf.sBCombine = "N"
            tgSpf.sInvVehSel = "N"  '8/31/01 Allow selective vehicle invoicing
            tgSpf.sVirtPkgCompute = "P"   'Recompute Price for Virtual Packages
            tgSpf.sCDefLogCopy = "N"   'Default Log 'Assign Copy' to off
            tgSpf.sBMissedDT = "R"  'Random
            tgSpf.sTax1Text = ""
            '12/17/06-Change to tax by agency or vehicle
            'tgSpf.iBTax(0) = 0
            tgSpf.sTax2Text = ""
            '12/17/06-Change to tax by agency or vehicle
            'tgSpf.iBTax(1) = 0
            tgSpf.iNoMnthNewBus = 13
            tgSpf.iNoMnthNewIsNew = 6
            tgSpf.sNewBusYearType = "R"
            tgSpf.iLastWeekBilled(0) = 0
            tgSpf.iLastWeekBilled(1) = 0
            'tgSpf.sUnused4 = ""
            tgSpf.sDefFillInv = "Y"
            tgSpf.sBOrderDPShow = "T"
            tgSpf.sBLaserForm = "1"
            tgSpf.lBCxfDisclaimer = 0
            tgSpf.lCxfContrComment = 0      '2-12-03
            tgSpf.lCxfInsertComment = 0     '2-12-03
            'Accounting
            'tgSpf.iRCorp(0) = 5 'Corporate calendar # of weeks
            'tgSpf.iRCorp(1) = 4
            'tgSpf.iRCorp(2) = 4
            'tgSpf.sRYEnd = "D"   'Use last day in dec.
            tgSpf.sInvSpotTimeZone = "N"
            tgSpf.sInvExportId = ""
            'tgSpf.iRCorp(0) = 0 'Corporate calendar # of weeks
            'tgSpf.iRCorp(1) = 0
            'tgSpf.iRCorp(2) = 0
            'tgSpf.sRYEnd = "L"   'Use last day in dec.
            tgSpf.sRUseCorpCal = "N"
            tgSpf.iRLastPurgedDate(0) = 0
            tgSpf.iRLastPurgedDate(1) = 0
            tgSpf.iRLastPay(0) = 0
            tgSpf.iRLastPay(1) = 0
            tgSpf.sRCurrAmt = "Y"
            tgSpf.sRUnbilled = "Y"
            'tgSpf.sLkCredit = ""
            tgSpf.sMOFCopyAssign = MGEITHERVEH Or FILLEITHERVEH
            'tgSpf.sRNewCntr = "I"
            tgSpf.iRPctCredit = 0
            tgSpf.sRRP = "S"
            tgSpf.iRPRP(0) = 0
            tgSpf.iRPRP(1) = 0
            tgSpf.iRCRP(0) = 0
            tgSpf.iRCRP(1) = 0
            tgSpf.iRNRP(0) = 0
            tgSpf.iRNRP(1) = 0
            gStrToPDN "0", 2, 6, tgSpf.sRB
            tgSpf.lRPenny = 0
            tgSpf.iBLastStdMnth(0) = 0
            tgSpf.iBLastStdMnth(1) = 0
            tgSpf.iBLastCalMnth(0) = 0
            tgSpf.iBLastCalMnth(1) = 0
            tgSpf.iRepPrintDate(0) = 0
            tgSpf.iRepPrintDate(1) = 0
            'For ilLoop = LBound(tgSpf.sRGLSuffix) To UBound(tgSpf.sRGLSuffix) Step 1
            '    tgSpf.sRGLSuffix(ilLoop) = ""
            'Next ilLoop
            'For ilLoop = LBound(tgSpf.sRName) To UBound(tgSpf.sRName) Step 1
            '    tgSpf.sRName(ilLoop) = ""
            '    tgSpf.sRTsfx(ilLoop) = ""
            '    tgSpf.sRAsfx(ilLoop) = ""
            'Next ilLoop
            tgSpf.sBPayName = ""
            tgSpf.sBPayAddr(0) = ""
            tgSpf.sBPayAddr(1) = ""
            tgSpf.sBPayAddr(2) = ""
            tgSpf.sRCollectPhoneNo = ""
            tgSpf.sRCollectContact = ""
            tgSpf.iRNoWks = 0
            tgSpf.iRCreditDate(0) = 0
            tgSpf.iRCreditDate(1) = 0
            tgSpf.sUseProdSptScr = "A"
            tgSpf.sRemoteUsers = "N"
            'tgSpf.iGRetainCntr = 24
            tgSpf.iRetainAffSpot = 24
            tgSpf.iRetainTrafSpot = 24
            tgSpf.iRetainTrafProp = 12
            'tgSpf.iGRetainCondensed = 24
            'tgSpf.iGRetainPayment = 60
            tgSpf.iMnfClientAbbr = 0
            tgSpf.iMnfInvTerms = 0      'Default terms (if zero use "15 days upon receipt"
            tgSpf.sRUseTrade = "N"
            tgSpf.sRUseMerch = "N"
            tgSpf.sRUsePromo = "N"
            ''Schedule parameters missing
            'tgSpf.sSchRNG = "T"
            'tgSpf.sSchMaL = "H"
            'tgSpf.sSchMdL = "D"
            'tgSpf.sSchMiL = "Q"
            'tgSpf.sSchCycle = "1"
            'tgSpf.sSchMove = "N"
            'tgSpf.sSchCompact = "N"
            'tgSpf.sSchPreempt = "E"
            'tgSpf.sSchHour = "B"
            'tgSpf.sSchDay = "B"
            'tgSpf.sSchQH = "B"
            tgSpf.iVehLen = 20
            tgSpf.sEDICallLetter = ""
            tgSpf.sEDIMediaType = ""
            tgSpf.sEDIBand = ""
            tgSpf.sUseCartNo = "Y"
            tgSpf.sExport = "0" '"N"    '# of spaces on invoice from top to logo (0 = standard)
            tgSpf.sImport = "0" '"N"    '# of spaces on invoice between logo and address (0 = standard)
            tgSpf.lDiscCurrCntrNo = 0
            tgSpf.iDiscDateRun(0) = 0
            tgSpf.iDiscDateRun(1) = 0
            tgSpf.sUseProdSptScr = "A"
            tgSpf.iGAlertInterval = 0
            tgSpf.sBSlspComm = "C"
            'Replaced with Features10
            'tgSpf.sXSDAddAdvtToISCI = "N" 'field was .sBSlspBack = "N"
            tgSpf.iBarterLPD(0) = 0
            tgSpf.iBarterLPD(1) = 0
            tgSpf.iGNoDaysPass = 0
            'Automation
            tgSpf.iPriUrfCode = 0
            tgSpf.iSecUrfCode = 0
            tgSpf.sWegenerGroupChar = "W"
            'Schedule
            tgSpf.sCmmlSchStatus = "R"
            tgSpf.iXDClientID = 0
            tgSpf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hlSpf, tgSpf, ilRecLen, INDEXKEY0)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrGetFirst):" & "Spf.Btr", Traffic
            On Error GoTo 0
        Else
            igSiteSet = True
            ilRet = btrGetFirst(hlSpf, tgSpf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            'ReDim ByteArray(LenB(tgSpf)) As Byte
            'ilRet = btrGetFirst(hlSpf, ByteArray(0), ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            'HMemCpy tgSpf, ByteArray(0), LenB(tgSpf)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrGetFirst):" & "Spf.Btr", Traffic
            On Error GoTo 0
            If tgSpf.sPostCalAff = "D" Then
                tgSpf.sUsingFeatures8 = Chr$(Asc(tgSpf.sUsingFeatures8) Or REPBYDT)
                tgSpf.sPostCalAff = "N"
                ilRet = btrUpdate(hlSpf, tgSpf, ilRecLen)
                On Error GoTo gSpfReadErr
                gBtrvErrorMsg ilRet, "gSpfRead (btrUpdate):" & "Spf.Btr", Traffic
                On Error GoTo 0
            End If
            'Get default terms
            hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
            ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                If tgSpf.iMnfInvTerms <> 0 Then
                    ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
                    tlMnfSrchKey.iCode = tgSpf.iMnfInvTerms
                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) And (Trim$(tlMnf.sType) = "J") And (Trim$(tlMnf.sUnitType) = "D") Then
                        sgDefaultTerms = Trim$(tlMnf.sName)
                    Else
                        tgSpf.iMnfInvTerms = 0
                    End If
                End If
            End If
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
        End If
        ilRet = btrClose(hlSpf)
        btrDestroy hlSpf
'        ilRet = csiSetStamp("SPF", sgSpfStamp)
'        ilRet = csiSetRec("SPF", 0, VarPtr(tgSpf), LenB(tgSpf))
'        'tgSpf.sGClient = ""
'        'tgSpf.iSLen(0) = 0
'        'tgSpf.iGRetainDead = 0
'        'tgSpf.sCommByCntr = ""
        'ilRet = csiGetRec("SPF", 0, VarPtr(tgSpf), LenB(tgSpf))
        
        '---------------------------------------------
        'TTP 10205 - 6/21/21 - JW - SPFX_Site_Options
        hlSpfx = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlSpfx, "", sgDBPath & "SPFX.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gSpfReadErr
        gBtrvErrorMsg ilRet, "gSpfRead (btrOpen):" & "SPFX.btr", Traffic
        On Error GoTo 0
        ilRecLen = Len(tgSpfx) 'Get and save record length
        llNoRec = btrRecords(hlSpfx) 'Obtain number of records
        If llNoRec = 0 Then
            'No SPFX record exists, create one
            btrDestroy hlSpfx
            hlSpfx = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hlSpfx, "", sgDBPath & "SPFX.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrOpen):" & "SPFX.btr", Traffic
            On Error GoTo 0
            '6/21/21 - JW - Fix Site Options Extended Table loading. Dont set igSiteSet, these options are not important to stop system
            'igSiteSet = False
            tgSpfx.iCode = 1  'key
            tgSpfx.sInvExpBillGroup = ""
            tgSpfx.sInvExpPrefix = ""
            tgSpfx.sInvExpProperty = ""
            
            ilRet = btrInsert(hlSpfx, tgSpfx, ilRecLen, INDEXKEY0)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrGetFirst):" & "SPFX.btr", Traffic
            On Error GoTo 0
        Else
            'A SPFX record exists, read it!
            'igSiteSet = True
            ilRet = btrGetFirst(hlSpfx, tgSpfx, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo gSpfReadErr
            gBtrvErrorMsg ilRet, "gSpfRead (btrGetFirst):" & "SPFX.btr", Traffic
            On Error GoTo 0
        End If
        'TTP 10205 - 6/21/21 - JW - Close spfx handles
        ilRet = btrClose(hlSpfx)
        btrDestroy hlSpfx
    End If
    Exit Sub
gSpfReadErr:
    ilRet = btrClose(hlSpf)
    btrDestroy hlSpf
    'TTP 10205 - 6/21/21 - JW - Close spfx handles
    ilRet = btrClose(hlSpfx)
    btrDestroy hlSpfx
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUrfDecrypt                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Decrypt user data               *
'*                                                     *
'*******************************************************
Sub gUrfDecrypt(tlUrf As URF)
'
'  gUrfDecrypt tlUrf
'  Where:
'        tlUrf (I) - user record to be decrypted
'
    Dim ilLoop As Integer
'    slName = tlUrf.sName
'    slDecrypt = ""
'    For ilLoop = 1 To Len(tlUrf.sName) Step 1
'        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
'        slName = Mid$(slName, 2)
'    Next ilLoop
'    tlUrf.sName = slDecrypt
    tlUrf.sName = gDecryptField(tlUrf.sName)
'    slName = tlUrf.sRept
'    slDecrypt = ""
'    For ilLoop = 1 To Len(tlUrf.sRept) Step 1
'        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
'        slName = Mid$(slName, 2)
'    Next ilLoop
'    tlUrf.sRept = slDecrypt
    tlUrf.sRept = gDecryptField(tlUrf.sRept)
'    slName = tlUrf.sPassword
'    slDecrypt = ""
'    For ilLoop = 1 To Len(tlUrf.sPassword) Step 1
'        slDecrypt = slDecrypt & Chr(Asc(slName) - 128)
'        slName = Mid$(slName, 2)
'    Next ilLoop
'    tlUrf.sPassword = slDecrypt
    tlUrf.sPassword = gDecryptField(tlUrf.sPassword)
    For ilLoop = LBound(tlUrf.sWin) To UBound(tlUrf.sWin) Step 1
        If Asc(tlUrf.sWin(ilLoop)) > 127 Then
            tlUrf.sWin(ilLoop) = Chr(Asc(tlUrf.sWin(ilLoop)) - 128)
        End If
    Next ilLoop
    If Asc(tlUrf.sGrid) > 127 Then
        tlUrf.sGrid = Chr(Asc(tlUrf.sGrid) - 128)
    End If
    If Asc(tlUrf.sPrice) > 127 Then
        tlUrf.sPrice = Chr(Asc(tlUrf.sPrice) - 128)
    End If
    If Asc(tlUrf.sCredit) > 127 Then
        tlUrf.sCredit = Chr(Asc(tlUrf.sCredit) - 128)
    End If
    If Asc(tlUrf.sPayRate) > 127 Then
        tlUrf.sPayRate = Chr(Asc(tlUrf.sPayRate) - 128)
    End If
    If Asc(tlUrf.sMerge) > 127 Then
        tlUrf.sMerge = Chr(Asc(tlUrf.sMerge) - 128)
    End If
    If Asc(tlUrf.sHideSpots) > 127 Then
        tlUrf.sHideSpots = Chr(Asc(tlUrf.sHideSpots) - 128)
    End If
    If Asc(tlUrf.sChgBilled) > 127 Then
        tlUrf.sChgBilled = Chr(Asc(tlUrf.sChgBilled) - 128)
    End If
    If Asc(tlUrf.sChgCntr) > 127 Then
        tlUrf.sChgCntr = Chr(Asc(tlUrf.sChgCntr) - 128)
    End If
    If Asc(tlUrf.sChgCrRt) > 127 Then
        tlUrf.sChgCrRt = Chr(Asc(tlUrf.sChgCrRt) - 128)
    End If
    If Asc(tlUrf.sBouChk) > 127 Then
        tlUrf.sBouChk = Chr(Asc(tlUrf.sBouChk) - 128)
    End If
    If Asc(tlUrf.sReprintLogAlert) > 127 Then
        tlUrf.sReprintLogAlert = Chr(Asc(tlUrf.sReprintLogAlert) - 128)
    End If
    If Asc(tlUrf.sIncompAlert) > 127 Then
        tlUrf.sIncompAlert = Chr(Asc(tlUrf.sIncompAlert) - 128)
    End If
    If Asc(tlUrf.sCompAlert) > 127 Then
        tlUrf.sCompAlert = Chr(Asc(tlUrf.sCompAlert) - 128)
    End If
    If Asc(tlUrf.sSchAlert) > 127 Then
        tlUrf.sSchAlert = Chr(Asc(tlUrf.sSchAlert) - 128)
    End If
    If Asc(tlUrf.sHoldAlert) > 127 Then
        tlUrf.sHoldAlert = Chr(Asc(tlUrf.sHoldAlert) - 128)
    End If
    If Asc(tlUrf.sRateCardAlert) > 127 Then
        tlUrf.sRateCardAlert = Chr(Asc(tlUrf.sRateCardAlert) - 128)
    End If
    If Asc(tlUrf.sResearchAlert) > 127 Then
        tlUrf.sResearchAlert = Chr(Asc(tlUrf.sResearchAlert) - 128)
    End If
    If Asc(tlUrf.sAvailAlert) > 127 Then
        tlUrf.sAvailAlert = Chr(Asc(tlUrf.sAvailAlert) - 128)
    End If
    If Asc(tlUrf.sCrdChkAlert) > 127 Then
        tlUrf.sCrdChkAlert = Chr(Asc(tlUrf.sCrdChkAlert) - 128)
    End If
    If Asc(tlUrf.sDeniedAlert) > 127 Then
        tlUrf.sDeniedAlert = Chr(Asc(tlUrf.sDeniedAlert) - 128)
    End If
    If Asc(tlUrf.sCrdLimitAlert) > 127 Then
        tlUrf.sCrdLimitAlert = Chr(Asc(tlUrf.sCrdLimitAlert) - 128)
    End If
    If Asc(tlUrf.sMoveAlert) > 127 Then
        tlUrf.sMoveAlert = Chr(Asc(tlUrf.sMoveAlert) - 128)
    End If
    If Asc(tlUrf.sShowNRMsg) > 127 Then
        tlUrf.sShowNRMsg = Chr(Asc(tlUrf.sShowNRMsg) - 128)
    End If
    If Asc(tlUrf.sWorkToDead) > 127 Then
        tlUrf.sWorkToDead = Chr(Asc(tlUrf.sWorkToDead) - 128)
    End If
    If Asc(tlUrf.sWorkToComp) > 127 Then
        tlUrf.sWorkToComp = Chr(Asc(tlUrf.sWorkToComp) - 128)
    End If
    If Asc(tlUrf.sWorkToHold) > 127 Then
        tlUrf.sWorkToHold = Chr(Asc(tlUrf.sWorkToHold) - 128)
    End If
    If Asc(tlUrf.sWorkToOrder) > 127 Then
        tlUrf.sWorkToOrder = Chr(Asc(tlUrf.sWorkToOrder) - 128)
    End If
    If Asc(tlUrf.sCompToIncomp) > 127 Then
        tlUrf.sCompToIncomp = Chr(Asc(tlUrf.sCompToIncomp) - 128)
    End If
    If Asc(tlUrf.sCompToDead) > 127 Then
        tlUrf.sCompToDead = Chr(Asc(tlUrf.sCompToDead) - 128)
    End If
    If Asc(tlUrf.sCompToHold) > 127 Then
        tlUrf.sCompToHold = Chr(Asc(tlUrf.sCompToHold) - 128)
    End If
    If Asc(tlUrf.sCompToOrder) > 127 Then
        tlUrf.sCompToOrder = Chr(Asc(tlUrf.sCompToOrder) - 128)
    End If
    If Asc(tlUrf.sIncompToDead) > 127 Then
        tlUrf.sIncompToDead = Chr(Asc(tlUrf.sIncompToDead) - 128)
    End If
    If Asc(tlUrf.sIncompToComp) > 127 Then
        tlUrf.sIncompToComp = Chr(Asc(tlUrf.sIncompToComp) - 128)
    End If
    If Asc(tlUrf.sIncompToHold) > 127 Then
        tlUrf.sIncompToHold = Chr(Asc(tlUrf.sIncompToHold) - 128)
    End If
    If Asc(tlUrf.sIncompToOrder) > 127 Then
        tlUrf.sIncompToOrder = Chr(Asc(tlUrf.sIncompToOrder) - 128)
    End If
    If Asc(tlUrf.sDeadToWork) > 127 Then
        tlUrf.sDeadToWork = Chr(Asc(tlUrf.sDeadToWork) - 128)
    End If
    If Asc(tlUrf.sHoldToOrder) > 127 Then
        tlUrf.sHoldToOrder = Chr(Asc(tlUrf.sHoldToOrder) - 128)
    End If
    If Asc(tlUrf.sReviseCntr) > 127 Then
        tlUrf.sReviseCntr = Chr(Asc(tlUrf.sReviseCntr) - 128)
    End If
    If Asc(tlUrf.sResvType) > 127 Then
        tlUrf.sResvType = Chr(Asc(tlUrf.sResvType) - 128)
    End If
    If Asc(tlUrf.sRemType) > 127 Then
        tlUrf.sRemType = Chr(Asc(tlUrf.sRemType) - 128)
    End If
    If Asc(tlUrf.sDRType) > 127 Then
        tlUrf.sDRType = Chr(Asc(tlUrf.sDRType) - 128)
    End If
    If Asc(tlUrf.sPIType) > 127 Then
        tlUrf.sPIType = Chr(Asc(tlUrf.sPIType) - 128)
    End If
    If Asc(tlUrf.sPSAType) > 127 Then
        tlUrf.sPSAType = Chr(Asc(tlUrf.sPSAType) - 128)
    End If
    If Asc(tlUrf.sPromoType) > 127 Then
        tlUrf.sPromoType = Chr(Asc(tlUrf.sPromoType) - 128)
    End If
    If Asc(tlUrf.sRefResvType) > 127 Then
        tlUrf.sRefResvType = Chr(Asc(tlUrf.sRefResvType) - 128)
    End If
    If Asc(tlUrf.sRegionCopy) > 127 Then
        tlUrf.sRegionCopy = Chr(Asc(tlUrf.sRegionCopy) - 128)
    End If
    If Asc(tlUrf.sChgPrices) > 127 Then
        tlUrf.sChgPrices = Chr(Asc(tlUrf.sChgPrices) - 128)
    End If
    If Asc(tlUrf.sChgLnBillPrice) > 127 Then
        tlUrf.sChgLnBillPrice = Chr(Asc(tlUrf.sChgLnBillPrice) - 128)
    End If
    If Asc(tlUrf.sActFlightButton) > 127 Then
        tlUrf.sActFlightButton = Chr(Asc(tlUrf.sActFlightButton) - 128)
    End If
    If Asc(tlUrf.sOldPassword(0)) > 127 Then
        tlUrf.sOldPassword(0) = gDecryptField(tlUrf.sOldPassword(0))
    End If
    If Asc(tlUrf.sOldPassword(1)) > 127 Then
        tlUrf.sOldPassword(1) = gDecryptField(tlUrf.sOldPassword(1))
    End If
    If Asc(tlUrf.sOldPassword(2)) > 127 Then
        tlUrf.sOldPassword(2) = gDecryptField(tlUrf.sOldPassword(2))
    End If
    If Asc(tlUrf.sLiveLogPostOnly) > 127 Then
        tlUrf.sLiveLogPostOnly = Chr(Asc(tlUrf.sLiveLogPostOnly) - 128)
    End If
    If Asc(tlUrf.sSportPropOnly) > 127 Then
        tlUrf.sSportPropOnly = Chr(Asc(tlUrf.sSportPropOnly) - 128)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUrfEncrypt                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Encrypt user data               *
'*                                                     *
'*******************************************************
Sub gUrfEncrypt(tlUrf As URF)
'
'  gUrfEncrypt tlUrf
'  Where:
'        tlUrf (I) - user record to be encrypted
'
    Dim ilLoop As Integer
    Dim slName As String
    Dim slEncrypt As String
    slName = tlUrf.sName
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sName) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sName = slEncrypt
    slName = tlUrf.sRept
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sRept) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sRept = slEncrypt
    slName = tlUrf.sPassword
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sPassword) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sPassword = slEncrypt
    For ilLoop = LBound(tlUrf.sWin) To UBound(tlUrf.sWin) Step 1
        tlUrf.sWin(ilLoop) = Chr(Asc(tlUrf.sWin(ilLoop)) + 128)
    Next ilLoop
    tlUrf.sGrid = Chr(Asc(tlUrf.sGrid) + 128)
    tlUrf.sPrice = Chr(Asc(tlUrf.sPrice) + 128)
    tlUrf.sCredit = Chr(Asc(tlUrf.sCredit) + 128)
    tlUrf.sPayRate = Chr(Asc(tlUrf.sPayRate) + 128)
    tlUrf.sMerge = Chr(Asc(tlUrf.sMerge) + 128)
    tlUrf.sHideSpots = Chr(Asc(tlUrf.sHideSpots) + 128)
    tlUrf.sChgBilled = Chr(Asc(tlUrf.sChgBilled) + 128)
    tlUrf.sChgCntr = Chr(Asc(tlUrf.sChgCntr) + 128)
    tlUrf.sChgCrRt = Chr(Asc(tlUrf.sChgCrRt) + 128)
    tlUrf.sBouChk = Chr(Asc(tlUrf.sBouChk) + 128)
    tlUrf.sReprintLogAlert = Chr(Asc(tlUrf.sReprintLogAlert) + 128)
    tlUrf.sIncompAlert = Chr(Asc(tlUrf.sIncompAlert) + 128)
    tlUrf.sCompAlert = Chr(Asc(tlUrf.sCompAlert) + 128)
    tlUrf.sSchAlert = Chr(Asc(tlUrf.sSchAlert) + 128)
    tlUrf.sHoldAlert = Chr(Asc(tlUrf.sHoldAlert) + 128)
    tlUrf.sRateCardAlert = Chr(Asc(tlUrf.sRateCardAlert) + 128)
    tlUrf.sResearchAlert = Chr(Asc(tlUrf.sResearchAlert) + 128)
    tlUrf.sAvailAlert = Chr(Asc(tlUrf.sAvailAlert) + 128)
    tlUrf.sCrdChkAlert = Chr(Asc(tlUrf.sCrdChkAlert) + 128)
    tlUrf.sDeniedAlert = Chr(Asc(tlUrf.sDeniedAlert) + 128)
    tlUrf.sCrdLimitAlert = Chr(Asc(tlUrf.sCrdLimitAlert) + 128)
    tlUrf.sMoveAlert = Chr(Asc(tlUrf.sMoveAlert) + 128)
    tlUrf.sShowNRMsg = Chr(Asc(tlUrf.sShowNRMsg) + 128)
    tlUrf.sWorkToDead = Chr(Asc(tlUrf.sWorkToDead) + 128)
    tlUrf.sWorkToComp = Chr(Asc(tlUrf.sWorkToComp) + 128)
    tlUrf.sWorkToHold = Chr(Asc(tlUrf.sWorkToHold) + 128)
    tlUrf.sWorkToOrder = Chr(Asc(tlUrf.sWorkToOrder) + 128)
    tlUrf.sCompToIncomp = Chr(Asc(tlUrf.sCompToIncomp) + 128)
    tlUrf.sCompToDead = Chr(Asc(tlUrf.sCompToDead) + 128)
    tlUrf.sCompToHold = Chr(Asc(tlUrf.sCompToHold) + 128)
    tlUrf.sCompToOrder = Chr(Asc(tlUrf.sCompToOrder) + 128)
    tlUrf.sIncompToDead = Chr(Asc(tlUrf.sIncompToDead) + 128)
    tlUrf.sIncompToComp = Chr(Asc(tlUrf.sIncompToComp) + 128)
    tlUrf.sIncompToHold = Chr(Asc(tlUrf.sIncompToHold) + 128)
    tlUrf.sIncompToOrder = Chr(Asc(tlUrf.sIncompToOrder) + 128)
    tlUrf.sDeadToWork = Chr(Asc(tlUrf.sDeadToWork) + 128)
    tlUrf.sHoldToOrder = Chr(Asc(tlUrf.sHoldToOrder) + 128)
    tlUrf.sReviseCntr = Chr(Asc(tlUrf.sReviseCntr) + 128)
    tlUrf.sResvType = Chr(Asc(tlUrf.sResvType) + 128)
    tlUrf.sRemType = Chr(Asc(tlUrf.sRemType) + 128)
    tlUrf.sDRType = Chr(Asc(tlUrf.sDRType) + 128)
    tlUrf.sPIType = Chr(Asc(tlUrf.sPIType) + 128)
    tlUrf.sPSAType = Chr(Asc(tlUrf.sPSAType) + 128)
    tlUrf.sPromoType = Chr(Asc(tlUrf.sPromoType) + 128)
    tlUrf.sRefResvType = Chr(Asc(tlUrf.sRefResvType) + 128)
    tlUrf.sRegionCopy = Chr(Asc(tlUrf.sRegionCopy) + 128)
    tlUrf.sChgPrices = Chr(Asc(tlUrf.sChgPrices) + 128)
    tlUrf.sChgLnBillPrice = Chr(Asc(tlUrf.sChgLnBillPrice) + 128)
    tlUrf.sActFlightButton = Chr(Asc(tlUrf.sActFlightButton) + 128)
    slName = tlUrf.sOldPassword(0)
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sOldPassword(0)) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sOldPassword(0) = slEncrypt
    slName = tlUrf.sOldPassword(1)
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sOldPassword(1)) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sOldPassword(1) = slEncrypt
    slName = tlUrf.sOldPassword(2)
    slEncrypt = ""
    For ilLoop = 1 To Len(tlUrf.sOldPassword(2)) Step 1
        slEncrypt = slEncrypt & Chr(Asc(slName) + 128)
        slName = Mid$(slName, 2)
    Next ilLoop
    tlUrf.sOldPassword(2) = slEncrypt
    tlUrf.sLiveLogPostOnly = Chr(Asc(tlUrf.sLiveLogPostOnly) + 128)
    tlUrf.sSportPropOnly = Chr(Asc(tlUrf.sSportPropOnly) + 128)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUrfRead                        *
'*                                                     *
'*             Created:5/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the user Option file       *
'*                                                     *
'*******************************************************
Sub gUrfRead(frm As Form, slUserName As String, ilSetWin As Integer, tlUrf() As URF, ilIncludeErased As Integer)
'
'   gUrfRead MainForm, slName, ilSetWin, tlUrf()
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       slName (I)- Name of the user
'       ilSetWin (I)- True= set igWinStatus (which jobs/lists allowed to do)
'                           False= don't alter igWinStatus
'       tlUrf (O)- where the user records are to be stored
'                   Note: tlUrf must be defined as Dim tlUrf() as URF
'

    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim slStamp As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilUpperBound As Integer
    Dim hlVef As Integer
    Dim ilVehOk As Integer
    Dim slGetStamp As String
    Dim tlTUrf As URF
    Dim ilMerge As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilUrf As Integer
    Dim blLimitGuide As Boolean
    Dim blInternalGuide As Boolean
    Dim blMatchFound As Boolean

    slStamp = gFileDateTime(sgDBPath & "Urf.btr") & slUserName

    'On Error GoTo gUrfReadErr2
    'ilRet = 0
    'imLowLimit = LBound(tlUrf)
    'If ilRet <> 0 Then
    '    sgUrfStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlUrf).Ptr <> 0 Then
        imLowLimit = LBound(tlUrf)
    Else
        sgUrfStamp = ""
        imLowLimit = 0
    End If

    If (sgUrfStamp <> "") And (sgUrfStamp <> "~") Then
        'Ignore time of URF since it is always changing as each user signon
        'Assume if sgUrfStamp defined, then it is Ok
        'If slStamp = sgUrfStamp Then
            Exit Sub
        'End If
        'slGetStamp = ""
    Else
        slGetStamp = "" 'gGetCSIStamp("URF")
    End If
    blMatchFound = False
    If slGetStamp <> "" Then
'        If ilSetWin Then
'            For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
'                igWinStatus(ilLoop) = 0
'            Next ilLoop
'        End If
'        sgUrfStamp = slGetStamp
'        ilRet = csiGetAlloc("URF", ilStartIndex, ilEndIndex)
'        ReDim tlUrf(ilStartIndex To ilEndIndex) As URF
'        For ilLoop = LBound(tlUrf) To UBound(tlUrf) Step 1
'            ilRet = csiGetRec("URF", ilLoop, VarPtr(tlUrf(ilLoop)), LenB(tlUrf(ilLoop)))
'            If ilSetWin Then
'                If tlUrf(ilLoop).iCode = 1 Then
'                    sgCPName = Trim$(tlUrf(ilLoop).sName)
'                End If
'                If tlUrf(ilLoop).iCode = 2 Then
'                    sgSUName = Trim$(tlUrf(ilLoop).sName)
'                End If
'                If ilSetWin Then
'                    For ilWin = LBound(igWinStatus) To UBound(igWinStatus) Step 1
'                        If tlUrf(ilLoop).sWin(ilWin) = "I" Then
'                            igWinStatus(ilWin) = 2
'                        ElseIf tlUrf(ilLoop).sWin(ilWin) = "V" Then
'                            If igWinStatus(ilWin) = 0 Then
'                                igWinStatus(ilWin) = 1
'                            End If
'                        End If
'                    Next ilWin
'                End If
'            End If
'        Next ilLoop
    Else
        ilRet = gObtainUrf()
        sgUrfStamp = slStamp
        If ilSetWin Then
            For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
                igWinStatus(ilLoop) = 0
            Next ilLoop
        End If
        '7/28/11: Code move to bottom of this routine
        ''Dan M 4//7/09 Limiting Guide means testing current table to see if it conforms to new limits; change as needed.
        'blLimitGuide = gTestLimitGuide(slUserName)
        'If blLimitGuide Then
        '    gLimitGuide tgPopUrf(2)
        '    gRecordLimitGuide tgPopUrf(2)
        'End If
        'blInternalGuide = gTestInternalGuide()
        'If blInternalGuide Then
        '    gInitSuperUser tgPopUrf(2)
        'End If
        ReDim tlUrf(0 To 0) As URF
        igUserByVeh = False
'        hlUrf = CBtrvTable(ONEHANDLE)
'        ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        On Error GoTo gUrfReadErr
'        gBtrvErrorMsg ilRet, "gUrfRead (btrOpen):" & "Urf.Btr", Frm
'        On Error GoTo 0
'        ilRecLen = Len(tlUrf(0))  'btrRecordLength(hlUrf)  'Get and save record length
'        llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
'        btrExtClear hlUrf   'Clear any previous extend operation
'        If btrRecords(hlUrf) <= 0 Then
'            btrDestroy hlUrf
'            hlUrf = CBtrvTable(TWOHANDLES)
'            ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'            On Error GoTo gUrfReadErr
'            gBtrvErrorMsg ilRet, "gUrfRead (btrOpen):" & "Urf.Btr", Frm
'            On Error GoTo 0
'        End If
'        hlVef = CBtrvTable(ONEHANDLE)
'        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        On Error GoTo gUrfReadErr
'        gBtrvErrorMsg ilRet, "gUrfRead (btrOpen):" & "Vef.btr", Frm
'        On Error GoTo 0
'        ilVefRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
'        ilUpperBound = UBound(tlUrf)
'        ilRet = btrGetFirst(hlUrf, tlUrf(ilUpperBound), ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        'Do While ilRet = BTRV_ERR_NONE
'
'        '    gUrfDecrypt tlUrf(ilUpperBound)
'        '    If (tlUrf(ilUpperBound).iVefCode > 0) And (tlUrf(ilUpperBound).sDelete <> "Y") Then
'        '        igUserByVeh = True
'        '    End If
'        '    If (Trim$(tlUrf(ilUpperBound).sName) = Trim$(slUserName)) And (tlUrf(ilUpperBound).sDelete <> "Y") Then
'        '        If (tlUrf(ilUpperBound).iCode <= 2) Or (tlUrf(ilUpperBound).iVefCode = 0) Then
'        '            ilVehOk = True
'        '        Else
'        '            ilVehOk = False
'        '            tlVefSrchKey.iCode = tlUrf(ilUpperBound).iVefCode
'        '            ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        '            If ilRet = BTRV_ERR_NONE Then
'    '                    If tgSpf.sSSellNet = "Y" Then
'    '                        If (tlVef.sType = "S") Or (tlVef.sType = "C") Then
'        '                        ilVehOk = True
'    '                        End If
'    '                    Else
'    '                        If tlVef.sType = "C" Then
'    '                            ilVehOk = True
'    '                        End If
'    '                    End If
'        '            End If
'        '        End If
'        '        If ilVehOk Then
'        '            If tlUrf(ilUpperBound).iCode = 1 Then
'        '                sgCPName = Trim$(tlUrf(ilUpperBound).sName)
'        '            End If
'        '            If tlUrf(ilUpperBound).iCode = 2 Then
'        '                sgSUName = Trim$(tlUrf(ilUpperBound).sName)
'        '            End If
'        '            ilUpperBound = ilUpperBound + 1
'        '            ReDim Preserve tlUrf(0 To ilUpperBound) As URF
'        '            If ilSetWin Then
'        '                For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
'        '                    If tlUrf(ilUpperBound - 1).sWin(ilLoop) = "I" Then
'        '                        igWinStatus(ilLoop) = 2
'        '                    ElseIf tlUrf(ilUpperBound - 1).sWin(ilLoop) = "V" Then
'        '                        If igWinStatus(ilLoop) = 0 Then
'        '                            igWinStatus(ilLoop) = 1
'        '                        End If
'        '                    End If
'        '                Next ilLoop
'        '            End If
'        '        End If
'        '    End If
'            'ilRet = btrGetNext(hlUrf, tlUrf(ilUpperBound), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        'Loop
'        If ilRet = BTRV_ERR_NONE Then
'            Call btrExtSetBounds(hlUrf, llNoRec, -1, "UC", "URF", "") 'Set extract limits (all records)
'            ilRet = btrExtAddField(hlUrf, 0, ilRecLen)  'Extract record
'            If ilRet = BTRV_ERR_NONE Then
'                ilRet = btrExtGetNext(hlUrf, tlUrf(ilUpperBound), ilRecLen, llRecPos)
'                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
'                        ilRecLen = Len(tlUrf(0))  'btrRecordLength(hlUrf)  'Get and save record length
'                        Do While ilRet = BTRV_ERR_REJECT_COUNT
'                            ilRet = btrExtGetNext(hlUrf, tlUrf(ilUpperBound), ilRecLen, llRecPos)
'                        Loop
'                        Do While ilRet = BTRV_ERR_NONE
'                            gUrfDecrypt tlUrf(ilUpperBound)
        'If UBound(tgPopUrf) > 1 Then
        If UBound(tgPopUrf) > 0 Then
            'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
            For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                tlUrf(ilUpperBound) = tgPopUrf(ilUrf)
                            If (tlUrf(ilUpperBound).iVefCode > 0) And ((tlUrf(ilUpperBound).sDelete <> "Y") Or (ilIncludeErased)) Then
                                igUserByVeh = True
                            End If
                            ilMerge = ilUpperBound
                            If (Trim$(tlUrf(ilUpperBound).sName) = Trim$(slUserName)) And ((tlUrf(ilUpperBound).sDelete <> "Y") Or (ilIncludeErased)) Then
                                If (tlUrf(ilUpperBound).iCode <= 2) Or (tlUrf(ilUpperBound).iVefCode = 0) Then
                                    ilVehOk = True
                                    If (tlUrf(ilUpperBound).iCode > 2) And (tlUrf(ilUpperBound).iVefCode = 0) And (ilUpperBound > 0) Then
                                        'Swap records
                                        tlTUrf = tlUrf(0)
                                        tlUrf(0) = tlUrf(ilUpperBound)
                                        tlUrf(ilUpperBound) = tlTUrf
                                        ilMerge = 0
                                    End If
                                Else
                    '                ilVehOk = False
                    '                tlVefSrchKey.iCode = tlUrf(ilUpperBound).iVefCode
                    '                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    '                If ilRet = BTRV_ERR_NONE Then
                    '                    If tgSpf.sSSellNet = "Y" Then
                    '                        If (tlVef.sType = "S") Or (tlVef.sType = "C") Then
                                                ilVehOk = True
                    '                        End If
                    '                    Else
                    '                        If tlVef.sType = "C" Then
                    '                            ilVehOk = True
                    '                        End If
                    '                    End If
                    '                End If
                                End If
                                If ilVehOk Then
                                    blMatchFound = True
                                    If tlUrf(ilMerge).iCode = 1 Then
                                        sgCPName = Trim$(tlUrf(ilMerge).sName)
                                    End If
                                    If tlUrf(ilMerge).iCode = 2 Then
                                        sgSUName = Trim$(tlUrf(ilMerge).sName)
                                    End If
                                    ilUpperBound = ilUpperBound + 1
                                    ReDim Preserve tlUrf(0 To ilUpperBound) As URF
                                    If ilSetWin Then
                                        For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
'                                            If (ilMerge = 0) Or (((ilLoop = RATECARDSJOB) Or (ilLoop = PROPOSALSJOB) Or (ilLoop = BUDGETSJOB) Or (ilLoop > STATIONFEEDJOB)) And (tgSpf.sUsingTraffic = "N")) Or (tgSpf.sUsingTraffic <> "N") Then
                                            If (ilMerge = 0) Or (((ilLoop = RATECARDSJOB) Or (ilLoop = PROPOSALSJOB) Or (ilLoop = BUDGETSJOB) Or (ilLoop > SLSPCOMMSJOB)) And (tgSpf.sUsingTraffic = "N")) Or (tgSpf.sUsingTraffic <> "N") Then
                                                If tlUrf(ilMerge).sWin(ilLoop) = "I" Then
                                                    igWinStatus(ilLoop) = 2
                                                ElseIf tlUrf(ilMerge).sWin(ilLoop) = "V" Then
                                                    If igWinStatus(ilLoop) = 0 Then
                                                        igWinStatus(ilLoop) = 1
                                                    End If
                                                End If
                                            End If
                                        Next ilLoop
                                    End If
                                End If
                            Else
                                If tlUrf(ilUpperBound).iCode = 1 Then
                                    sgCPName = Trim$(tlUrf(ilUpperBound).sName)
                                End If
                                If tlUrf(ilUpperBound).iCode = 2 Then
                                    sgSUName = Trim$(tlUrf(ilUpperBound).sName)
                                End If
                            End If
'                            ilRet = btrExtGetNext(hlUrf, tlUrf(ilUpperBound), ilRecLen, llRecPos)
'                            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                                ilRet = btrExtGetNext(hlUrf, tlUrf(ilUpperBound), ilRecLen, llRecPos)
'                            Loop
'                        Loop
'                    End If
'                End If
'            End If
            Next ilUrf
        End If
'        If ilRet <> BTRV_ERR_END_OF_FILE Then
'            On Error GoTo gUrfReadErr
'            gBtrvErrorMsg ilRet, "gUrfRead (btrGetFirst or Next):" & "Urf.Btr", Frm
'            On Error GoTo 0
'        End If
        If ilUpperBound > 0 Then
            ilUpperBound = ilUpperBound - 1
            ReDim Preserve tlUrf(0 To ilUpperBound) As URF
        ''6/30/11: Handle is not defined any longer, test tgPopUrf instead
        ''ElseIf btrRecords(hlUrf) > 0 Then
        ''Else
        'ElseIf UBound(tgPopUrf) <= 1 Then
        ElseIf UBound(tgPopUrf) <= 0 Then
            blMatchFound = True
            hlUrf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            'On Error GoTo gUrfReadErr
            'gBtrvErrorMsg ilRet, "gUrfRead (btrOpen):" & "Urf.Btr", Frm
            'On Error GoTo 0
            ilRecLen = Len(tlUrf(0))  'btrRecordLength(hlUrf)  'Get and save record length
            tlUrf(0).sName = "Counterpoint"
            tlUrf(0).iVefCode = 0
            tlUrf(0).sPassword = "JFDL"
            tlUrf(0).sRept = "System"
            tlUrf(0).sPhoneNo = ""
            tlUrf(0).sCity = ""
            tlUrf(0).lEMailCefCode = 0
            tlUrf(0).iSlfCode = 0
            tlUrf(0).sClnMoYr = "M"    'Calendar by month
            tlUrf(0).sClnType = "S"    'Calendar type-standard
            tlUrf(0).sClnLayout = "A"  'Calendar- Across
            tlUrf(0).iClnLeft = 100    'Calendar left
            tlUrf(0).iClnTop = 200 'Calendar top
            tlUrf(0).iClcLeft = 2775   'Traffic.Width - 2775 - 100   '2625 width of calc form
            tlUrf(0).iClcTop = 200
            For ilLoop = 0 To 69 Step 1
                tlUrf(0).sWin(ilLoop) = "I"
            Next ilLoop
            tlUrf(0).iDefVeh = 0
            tlUrf(0).sGrid = "I"
            tlUrf(0).sPrice = "I"
            tlUrf(0).sCredit = "I"
            tlUrf(0).sPayRate = "I"
            tlUrf(0).sMerge = "I"
            tlUrf(0).sDelete = "N"
            tlUrf(0).sHideSpots = "I"
            tlUrf(0).sChgBilled = "I"
            tlUrf(0).sChgCntr = "I"
            tlUrf(0).sChgCrRt = "I"
            tlUrf(0).sBouChk = "I"
            tlUrf(0).sReprintLogAlert = "Y"
            tlUrf(0).sIncompAlert = "Y"
            tlUrf(0).sCompAlert = "Y"
            tlUrf(0).sSchAlert = "Y"
            tlUrf(0).sHoldAlert = "Y"
            tlUrf(0).sPrgmmaticAlert = "V"
            tlUrf(0).sAdvanceAvails = "H"
            tlUrf(0).sShowNRMsg = "Y"
            tlUrf(0).sRateCardAlert = "Y"
            tlUrf(0).sResearchAlert = "Y"
            tlUrf(0).sAvailAlert = "Y"
            tlUrf(0).sCrdChkAlert = "Y"
            tlUrf(0).sDeniedAlert = "Y"
            tlUrf(0).sCrdLimitAlert = "Y"
            tlUrf(0).sMoveAlert = "Y"
            tlUrf(0).sWorkToDead = "Y"
            tlUrf(0).sWorkToComp = "Y"
            tlUrf(0).sWorkToHold = "Y"
            tlUrf(0).sWorkToOrder = "Y"
            tlUrf(0).sCompToIncomp = "Y"
            tlUrf(0).sCompToDead = "Y"
            tlUrf(0).sCompToHold = "Y"
            tlUrf(0).sCompToOrder = "Y"
            tlUrf(0).sIncompToDead = "Y"
            tlUrf(0).sIncompToComp = "Y"
            tlUrf(0).sIncompToHold = "Y"
            tlUrf(0).sIncompToOrder = "Y"
            tlUrf(0).sDeadToWork = "Y"
            tlUrf(0).sHoldToOrder = "Y"
            tlUrf(0).sReviseCntr = "Y"
            tlUrf(0).sResvType = "I"
            tlUrf(0).sRemType = "I"
            tlUrf(0).sDRType = "I"
            tlUrf(0).sPIType = "I"
            tlUrf(0).sPSAType = "I"
            tlUrf(0).sPromoType = "I"
            tlUrf(0).sRefResvType = "I"
            tlUrf(0).sUseComputeCMC = "H"
            tlUrf(0).iGroupNo = 0
            tlUrf(0).sBlockRU = "N"
            tlUrf(0).sRCView = "R"
            tlUrf(0).sRegionCopy = "I"
            tlUrf(0).sChgPrices = "I"
            tlUrf(0).sChgLnBillPrice = "I"
            tlUrf(0).sActFlightButton = "I"
            tlUrf(0).iRemoteUserID = 0
            ' Dan M new selected fields 4/10/09
            tlUrf(0).sChangeCSIDate = "I"
            tlUrf(0).sAllowInvDisplay = "I"
            tlUrf(0).sActivityLog = "V"
            'Add new value into gInitSuperUser (initsubs.bas)
            tlUrf(0).iCode = 0  'Autoincrement
            tlUrf(0).iRemoteID = 0
            tlUrf(0).iAutoCode = tlUrf(0).iCode
            tlUrf(0).sPDFDrvChar = ""
            tlUrf(0).iPDFDnArrowCnt = 0
            tlUrf(0).sPrtDrvChar = ""
            tlUrf(0).iPrtDnArrowCnt = 0
            tlUrf(0).sPrtNameAltKey = ""
            tlUrf(0).iPrtNoEnterKeys = 0
            tlUrf(0).sOldPassword(0) = ""
            tlUrf(0).sOldPassword(1) = ""
            tlUrf(0).sOldPassword(2) = ""
            tlUrf(0).sLiveLogPostOnly = "N"
            tlUrf(0).sSportPropOnly = "N"
            gUrfEncrypt tlUrf(0)
            ilRet = btrInsert(hlUrf, tlUrf(0), ilRecLen, INDEXKEY0)
            'On Error GoTo gPopUserListBoxErr
            'gBtrvErrorMsg ilRet, "gPopUserListBox (btrGetFirst):" & "Urf.Btr", Frm
            'On Error GoTo 0
            gUrfDecrypt tlUrf(0)
            gGetSyncDateTime slSyncDate, slSyncTime
            Do
                tlUrf(0).iRemoteID = 0
                tlUrf(0).iAutoCode = tlUrf(0).iCode
                gPackDate slSyncDate, tlUrf(0).iSyncDate(0), tlUrf(0).iSyncDate(1)
                gPackTime slSyncTime, tlUrf(0).iSyncTime(0), tlUrf(0).iSyncTime(1)
                gUrfEncrypt tlUrf(0)
                ilRet = btrUpdate(hlUrf, tlUrf(0), ilRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            gUrfDecrypt tlUrf(0)
            sgCPName = Trim$(tlUrf(0).sName)
            ReDim Preserve tlUrf(0 To 1) As URF
            ' Dan M 4/07/09 'Guide' limited functionality
            gInitGuide tlUrf(1)
'            tlUrf(1) = tlUrf(0)
'            tlUrf(1).sName = "Guide"
'            tlUrf(1).iVefCode = 0
'            tlUrf(1).sPassword = "Guide"
'            tlUrf(1).iCode = 0  'Autoincrement
'            tlUrf(1).iRemoteID = 0
'            tlUrf(1).iAutoCode = tlUrf(1).iCode
            gUrfEncrypt tlUrf(1)
            ilRet = btrInsert(hlUrf, tlUrf(1), ilRecLen, INDEXKEY0)
            'On Error GoTo gPopUserListBoxErr
            'gBtrvErrorMsg ilRet, "gPopUserListBox (btrInsert):" & "Urf.Btr", Frm
            'On Error GoTo 0
            gUrfDecrypt tlUrf(1)
            Do
                tlUrf(1).iRemoteID = 0
                tlUrf(1).iAutoCode = tlUrf(1).iCode
                gPackDate slSyncDate, tlUrf(1).iSyncDate(0), tlUrf(1).iSyncDate(1)
                gPackTime slSyncTime, tlUrf(1).iSyncTime(0), tlUrf(1).iSyncTime(1)
                gUrfEncrypt tlUrf(1)
                ilRet = btrUpdate(hlUrf, tlUrf(1), ilRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            gUrfDecrypt tlUrf(1)
            sgSUName = Trim$(tlUrf(1).sName)
            If ilSetWin Then
                For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
                    'If (((ilLoop = RATECARDSJOB) Or (ilLoop = PROPOSALSJOB) Or (ilLoop = BUDGETSJOB) Or (ilLoop > STATIONFEEDJOB)) And (tgSpf.sUsingTraffic = "N")) Or (tgSpf.sUsingTraffic <> "N") Then
                    If (((ilLoop = RATECARDSJOB) Or (ilLoop = PROPOSALSJOB) Or (ilLoop = BUDGETSJOB) Or (ilLoop > SLSPCOMMSJOB)) And (tgSpf.sUsingTraffic = "N")) Or (tgSpf.sUsingTraffic <> "N") Then
                        If tlUrf(0).sWin(ilLoop) = "I" Then
                            igWinStatus(ilLoop) = 2
                        ElseIf tlUrf(0).sWin(ilLoop) = "V" Then
                            If igWinStatus(ilLoop) = 0 Then
                                igWinStatus(ilLoop) = 1
                            End If
                        End If
                    End If
                Next ilLoop
            End If
            btrClose hlUrf
            sgPopUrfTag = ""
            ilRet = gObtainUrf()
            btrDestroy hlUrf
        End If
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
'        btrDestroy hlVef
'        ilRet = btrClose(hlUrf)
'        btrDestroy hlUrf
'        ilRet = csiSetStamp("URF", sgUrfStamp)
'        ilRet = csiSetAlloc("URF", LBound(tlUrf), UBound(tlUrf))
'        For ilLoop = LBound(tlUrf) To UBound(tlUrf) Step 1
'            ilRet = csiSetRec("URF", ilLoop, VarPtr(tlUrf(ilLoop)), LenB(tlUrf(ilLoop)))
'        Next ilLoop
        '7/28/11: Move from above
        'Dan M 4//7/09 Limiting Guide means testing current table to see if it conforms to new limits; change as needed.
        blLimitGuide = gTestLimitGuide(slUserName)
        If blLimitGuide Then
            'gLimitGuide tgPopUrf(2)
            gLimitGuide tgPopUrf(1)
            'gRecordLimitGuide tgPopUrf(2)
            gRecordLimitGuide tgPopUrf(1)
        End If
        blInternalGuide = gTestInternalGuide()
        If blInternalGuide Then
            'gInitSuperUser tgPopUrf(2)
            gInitSuperUser tgPopUrf(1)
        End If
    End If
    If Not blMatchFound Then
        tlUrf(0).sName = ""
    End If
    Exit Sub

    ilRet = btrClose(hlVef)
    ilRet = btrClose(hlUrf)
    btrDestroy hlVef
    btrDestroy hlUrf
    gDbg_HandleError "FileSubs: gUrfRead"
gUrfReadErr2:
    ilRet = 1
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gVirtVefSpotPrice               *
'*                                                     *
'*             Created:5/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute price of spot if virtual*
'*                     vehicle                         *
'*                                                     *
'*******************************************************
Function gVirtVefSpotPrice(hlVef As Integer, hlVsf As Integer, ilLnVefCode As Integer, ilSpotVefCode As Integer, slFlightPrice As String) As String
'
'   slPrice = gVirtVefSpotPrice(hlVef, hlVsf, ilLnVefCode, ilSpotVefCode, slFlightPrice)
'   Where:
'       hlVef(I)- Open handle to Vehicle file
'       hlVsf(I)- Open handle to Virtual Vehicle List File
'       ilLnVefCode(I)- Line vehicle code
'       ilSpotVefCode(I)- Spot Vehicle Code
'       slFlightPrice(I)- Flight Price
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlVef As VEF                'VEF record image
    Dim tlVefSrchKey As INTKEY0            'VSF record image
    Dim ilVefRecLen As Integer        'VSF record length
    'Dim tlVsf As VSF
    Dim tlVsfSrchKey As LONGKEY0            'VSF record image
    Dim ilVsfReclen As Integer        'VSF record length
    Dim ilVsf As Integer
    Dim slAmount As String
    Dim slPct As String

    'Determine if virtual vehicle
    gVirtVefSpotPrice = slFlightPrice
    If InStr(slFlightPrice, ".") = 0 Then
        Exit Function
    End If
    If ilLnVefCode = ilSpotVefCode Then
        Exit Function
    End If
    For ilLoop = LBound(igVirtVefCode) To UBound(igVirtVefCode) - 1 Step 1
        If igVirtVefCode(ilLoop) = ilLnVefCode Then
            ilVefRecLen = Len(tlVef)
            tlVefSrchKey.iCode = ilLnVefCode
            ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                gVirtVefSpotPrice = ".00"
                Exit Function
            End If
            ilVsfReclen = Len(tmVsf)
            tlVsfSrchKey.lCode = tlVef.lVsfCode
            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                gVirtVefSpotPrice = ".00"
                Exit Function
            End If
            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilVsf) = ilSpotVefCode Then
                    If tmVsf.iNoSpots(ilVsf) > 0 Then
                        'gPDNToStr tmVsf.sFSComm(ilVsf), 4, slPct
                        'slPct = Trim$(Str$(tmVsf.lFSComm(ilVsf)))
                        slPct = gLongToStrDec(tmVsf.lFSComm(ilVsf), 4)
                        slAmount = gDivStr(gDivStr(gMulStr(slFlightPrice, slPct), "100.00"), Trim$(str$(tmVsf.iNoSpots(ilVsf))))
                        gVirtVefSpotPrice = gRoundStr(slAmount, ".01", 2)
                    Else
                        gVirtVefSpotPrice = ".00"
                    End If
                    Exit Function
                End If
            Next ilVsf
            gVirtVefSpotPrice = ".00"
            Exit Function
        End If
    Next ilLoop
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gVpfFind                        *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the Vehicle Option file   *
'*                                                     *
'*******************************************************
Function gVpfFind(frm As Form, ilVefCode As Integer) As Integer
'
'   ilRet = gVpfFind (MainForm, ilCode)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       ilCode (I)- Vehicle code number to be found
'       ilRet (O)- Index into tgVpf for record associated with Vehicle
'                  Note: If Vehicle preference doesn't exist, it will be created
'

    Dim ilLoop As Integer
    Dim ilFound As Integer  'True=record found
    Dim hlVpf As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vpf record length
    Dim ilRet As Integer
    Dim tlVpf As VPF
    Dim slSyncDate As String
    Dim slSyncTime As String

    If ilVefCode <= 0 Then
        gVpfFind = -1
        Exit Function
    End If
    Do
        On Error GoTo gVpfFindErr
        ilRet = gVpfRead()
        On Error GoTo 0
        ilFound = -1
        If ilRet Then
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(ilVefCode)
                If ilLoop <> -1 Then
                    ilFound = ilLoop
            '        Exit For
                End If
            'Next ilLoop
        End If
        If ilFound = -1 Then
            'Create the record
            hlVpf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo gVpfFindErr1
            gBtrvErrorMsg ilRet, "gVpfFind (btrOpen):" & "Vpf.Btr", frm
            On Error GoTo 0
            ilRecLen = Len(tlVpf)  'btrRecordLength(hlVpf)  'Get and save record length
            mVpfInit ilVefCode, tlVpf
            gGetSyncDateTime slSyncDate, slSyncTime
            'gPackDate slSyncDate, tlVpf.iSyncDate(0), tlVpf.iSyncDate(1)
            'gPackTime slSyncTime, tlVpf.iSyncTime(0), tlVpf.iSyncTime(1)
            ilRet = btrInsert(hlVpf, tlVpf, ilRecLen, INDEXKEY0)
            On Error GoTo gVpfFindErr1
            gBtrvErrorMsg ilRet, "gVpfFind (btrInser):" & "Vpf.Btr", frm
            On Error GoTo 0
            ilRet = btrClose(hlVpf)
            btrDestroy hlVpf
            
            'sgVpfStamp = ""
            'ilRet = csiSetStamp("VPF", "")
            ReDim Preserve tgVpf(0 To UBound(tgVpf) + 1) As VPF
            tgVpf(UBound(tgVpf)) = tlVpf
            If UBound(tgVpf) > 0 Then
                ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0
            End If
            
            '11/26/17
            gFileChgdUpdate "vpf.btr", False
            
        End If
    Loop Until ilFound >= 0
    gVpfFind = ilFound
    Exit Function
gVpfFindErr:
    On Error GoTo 0
    gDbg_HandleError "FileSubs: gVpfFind"
'    gVpfFind = -1
'    Exit Function
gVpfFindErr1:
    On Error GoTo 0
    ilRet = btrClose(hlVpf)
    btrDestroy hlVpf
    gDbg_HandleError "FileSubs: gVpfFine"
'    gVpfFind = -1
'    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mVpfFind                        *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the Vehicle Option file    *
'*                     same as gVpfFind except         *
'*                     open/clode removed              *
'*                                                     *
'*******************************************************
Function gVpfFindHd(frm As Form, ilVefCode As Integer, hlVpf As Integer) As Integer 'VBC NR
'
'   ilRet = gVpfFind (MainForm, ilCode)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       ilCode (I)- Vehicle code number to be found
'       ilRet (O)- Index into tgVpf for record associated with Vehicle
'                  Note: If Vehicle preference doesn't exist, it will be created
'

    Dim ilLoop As Integer 'VBC NR
    Dim ilFound As Integer  'True=record found 'VBC NR
    Dim ilRecLen As Integer     'Vpf record length 'VBC NR
    Dim ilRet As Integer 'VBC NR
    Dim tlVpf As VPF 'VBC NR
    Dim slSyncDate As String 'VBC NR
    Dim slSyncTime As String 'VBC NR

    If ilVefCode <= 0 Then
        gVpfFindHd = -1
        Exit Function
    End If
    Do 'VBC NR
        On Error GoTo gVpfFindHdErr 'VBC NR
        ilRet = gVpfRead() 'VBC NR
        On Error GoTo 0 'VBC NR
        ilFound = -1 'VBC NR
        If ilRet Then 'VBC NR
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(ilVefCode) 'VBC NR
                If ilLoop <> -1 Then 'VBC NR
                    ilFound = ilLoop 'VBC NR
            '        Exit For
                End If 'VBC NR
            'Next ilLoop
        End If 'VBC NR
        If ilFound = -1 Then 'VBC NR
            'Create the record
            'hlVpf = CBtrvTable(TWOHANDLES)
            'ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            'On Error GoTo gVpfFindErr1
            'gBtrvErrorMsg ilRet, "gVpfFind (btrOpen):" & "Vpf.Btr", Frm
            'On Error GoTo 0
            ilRecLen = Len(tlVpf)  'btrRecordLength(hlVpf)  'Get and save record length 'VBC NR
            mVpfInit ilVefCode, tlVpf 'VBC NR
            gGetSyncDateTime slSyncDate, slSyncTime 'VBC NR
            'gPackDate slSyncDate, tlVpf.iSyncDate(0), tlVpf.iSyncDate(1)
            'gPackTime slSyncTime, tlVpf.iSyncTime(0), tlVpf.iSyncTime(1)
            ilRet = btrInsert(hlVpf, tlVpf, ilRecLen, INDEXKEY0) 'VBC NR
            On Error GoTo gVpfFindHdErr1 'VBC NR
            gBtrvErrorMsg ilRet, "gVpfFindHd (btrInser):" & "Vpf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            'ilRet = btrClose(hlVpf)
            'btrDestroy hlVpf
            'sgVpfStamp = ""
            'ilRet = csiSetStamp("VPF", "")

            ReDim Preserve tgVpf(0 To UBound(tgVpf) + 1) As VPF 'VBC NR
            tgVpf(UBound(tgVpf)) = tlVpf 'VBC NR
            If UBound(tgVpf) > 0 Then 'VBC NR
                ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0 'VBC NR
            End If 'VBC NR
            
            '11/26/17
            gFileChgdUpdate "vpf.btr", False
            
        End If 'VBC NR
    Loop Until ilFound >= 0 'VBC NR
    gVpfFindHd = ilFound 'VBC NR
    Exit Function 'VBC NR
gVpfFindHdErr: 'VBC NR
    On Error GoTo 0 'VBC NR
    gDbg_HandleError "FileSubs: gVpfFindHd" 'VBC NR
'    gVpfFindHd = -1
'    Exit Function
gVpfFindHdErr1: 'VBC NR
    On Error GoTo 0 'VBC NR
    'ilRet = btrClose(hlVpf)
    'btrDestroy hlVpf
    gDbg_HandleError "FileSubs: gVpfFindHd" 'VBC NR
'    gVpfFindHd = -1
'    Exit Function
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gVpfFindIndex                   *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the Vehicle Option file    *
'*                                                     *
'*******************************************************
Function gVpfFindIndex(ilVefCode As Integer) As Integer
'
'   ilRet = gVpfFindIndex (ilCode)
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       ilCode (I)- Vehicle code number to be found
'       ilRet (O)- Index into tgVpf for record associated with Vehicle
'                   -1 if not found
'

    Dim ilLoop As Integer
    Dim ilFound As Integer  'True=record found

    ilFound = -1
    'For ilLoop = 0 To UBound(tgVpf) Step 1
    '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
        ilLoop = gBinarySearchVpf(ilVefCode)
        If ilLoop <> -1 Then
            ilFound = ilLoop
    '        Exit For
        End If
    'Next ilLoop
    gVpfFindIndex = ilFound
    Exit Function
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:gVpfRead                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the Vehicle Option file   *
'*                                                     *
'*******************************************************
'TTP 11055 - Copy Changed Alert not triggered when copy was added for vehicle/week that the final log had been generated for
'Function gVpfRead() As Integer
Function gVpfRead(Optional blBypassFileChangeCheck As Boolean = False) As Integer
'
'   ilRet = gVpfRead ()
'   Where:
'       ilRet (O)- True if read
'                       False if none read
'

    Dim ilRecLen As Integer     'Vpf record length
    Dim hlVpf As Integer     'Vehicle Option file handle
    Dim hlVef As Integer     'Vehicle Option file handle
'    Dim tlVef As VEF
'    Dim ilVefRecLen As Integer
    'Dim hlVsf As Integer     'Vehicle Option file handle
    'Dim tlVsf As VSF
    'Dim ilVsfRecLen As Integer
    'Dim tlVsfSrchKey As LONGKEY0
    Dim slStamp As String
    Dim llNoRec As Long      'Number of records (must be one)
    Dim llNoVpfRec As Long
    Dim ilRet As Integer
    Dim ilBound As Integer
    Dim ilVpfMax As Integer
    Dim llRecPos As Long
    Dim slGetStamp As String
    Dim ilVef As Integer
    Dim ilAdd As Integer
    slStamp = gFileDateTime(sgDBPath & "Vpf.btr")

    'TTP 11055 - Copy Changed Alert not triggered when copy was added for vehicle/week that the final log had been generated for
    If blBypassFileChangeCheck = True Then
        'we're going to skip checking if the fact table updated
    Else
        '11/26/17: Check Changed date/time
        If Not gFileChgd("vpf.btr") Then
            gVpfRead = True
            Exit Function
        End If
    End If
    'On Error GoTo gVpfReadErr2
    'ilRet = 0
    'imLowLimit = LBound(tgVpf)
    'If ilRet <> 0 Then
    '    sgVpfStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgVpf).Ptr <> 0 Then
        imLowLimit = LBound(tgVpf)
    Else
        sgVpfStamp = ""
        imLowLimit = 0
    End If

    '6/9/18: replaced with gFileChgd
    'If (sgVpfStamp <> "") And (sgVpfStamp <> "~") Then
    '    'If slStamp = sgVpfStamp Then
    '        gVpfRead = True
    '        Exit Function
    '    'End If
    '    'slGetStamp = ""
    'Else
        slGetStamp = "" 'gGetCSIStamp("VPF")
        igVehicleDefined = False    'True
    'End If
    If slGetStamp <> "" Then
'        sgVpfStamp = slGetStamp
'        ilRet = csiGetAlloc("VPF", ilStartIndex, ilEndIndex)
'        ReDim tgVpf(ilStartIndex To ilEndIndex) As VPF
'        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
'            ilRet = csiGetRec("VPF", ilLoop, VarPtr(tgVpf(ilLoop)), LenB(tgVpf(ilLoop)))
'        Next ilLoop
        gVpfRead = True
    Else
        sgVpfStamp = slStamp
        hlVpf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        On Error GoTo gVpfReadErr
'        gBtrvErrorMsg ilRet, "gVpfRead (btrOpen):" & "Vpf.Btr", Frm
'        On Error GoTo 0
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gVpfRead = False
            btrDestroy hlVpf
            Exit Function
        End If
        hlVef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        On Error GoTo gVpfReadErr
'        gBtrvErrorMsg ilRet, "gVpfRead (btrOpen):" & "Vef.Btr", Frm
'        On Error GoTo 0
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gVpfRead = False
            btrDestroy hlVpf
'            btrDestroy hlVef
            Exit Function
        End If
'        ilVefRecLen = Len(tmVef)
        'hlVsf = CBtrvTable(ONEHANDLE)
        'ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo gVpfReadErr
        'gBtrvErrorMsg ilRet, "gVpfRead (btrOpen):" & "Vsf.Btr", Frm
        'On Error GoTo 0
        'ilVsfRecLen = Len(tmVsf)
        llNoVpfRec = btrRecords(hlVpf) 'Obtain number of records
        If llNoVpfRec <= btrRecords(hlVef) Then
            llNoVpfRec = btrRecords(hlVef) 'Obtain number of records
        End If
        btrDestroy hlVef
        If llNoVpfRec <= 0 Then
            ReDim tgVpf(0 To 0) As VPF
            ilRet = btrClose(hlVpf)
            btrDestroy hlVpf
'            ilRet = btrClose(hlVef)
'            btrDestroy hlVef
            'ilRet = btrClose(hlVsf)
            'btrDestroy hlVsf
            gVpfRead = False
            igVehicleDefined = False
            Exit Function
        End If
        igVehicleDefined = True
        ReDim tgVpf(0 To llNoVpfRec - 1) As VPF
        ilRecLen = Len(tgVpf(0)) 'Get and save record length
        llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hlVpf   'Clear any previous extend operation
        ilBound = 0
        ilRet = btrGetFirst(hlVpf, tgVpf(ilBound), ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Do While ilRet = BTRV_ERR_NONE
        '    If ilBound >= llNoRec - 1 Then
        '        Exit Do
        '    End If
        '    ilBound = ilBound + 1
        '    ilRet = btrGetNext(hlVpf, tgVpf(ilBound), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Loop
        If ilRet = BTRV_ERR_NONE Then
            Call btrExtSetBounds(hlVpf, llNoRec, -1, "UC", "VPF", "") 'Set extract limits (all records)
            ilRet = btrExtAddField(hlVpf, 0, ilRecLen)  'Extract record
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrExtGetNext(hlVpf, tgVpf(ilBound), ilRecLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                        ilRecLen = Len(tgVpf(0)) 'Get and save record length
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlVpf, tgVpf(ilBound), ilRecLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            If ilBound >= llNoVpfRec - 1 Then
                                Exit Do
                            End If
                            If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                                ilAdd = True
                            Else
                                ilVef = gBinarySearchVef(tgVpf(ilBound).iVefKCode, True)
                                If ilVef <> -1 Then
                                    ilAdd = True
                                Else
                                    ilAdd = False
                                End If
                            End If
                            If ilAdd Then
                                ilBound = ilBound + 1
                            End If
                            ilRet = btrExtGetNext(hlVpf, tgVpf(ilBound), ilRecLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hlVpf, tgVpf(ilBound), ilRecLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            End If
        End If
        If ilRet <> BTRV_ERR_END_OF_FILE Then
'            On Error GoTo gVpfReadErr
'            gBtrvErrorMsg ilRet, "gVpfRead (btrGetFirst or Next):" & "Vpf.Btr", Frm
'            On Error GoTo 0
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gVpfRead = False
'                btrDestroy hlVef
                btrDestroy hlVpf
                Exit Function
            End If
        End If
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilVpfMax = ilBound - 1
        Else
            ilVpfMax = ilBound
            ilBound = ilBound + 1
        End If
        ReDim Preserve tgVpf(0 To ilVpfMax) As VPF
        'Make Vpf records for virtual vehicles
        'ilRet = btrGetFirst(hlVef, tlVef, ilVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Do While ilRet = BTRV_ERR_NONE
        '    If tlVef.sType = "V" Then
        '        If ilBound > UBound(tgVpf) Then
        '            ReDim Preserve tgVpf(0 To ilBound) As VPF
        '        End If
        '        mVpfInit tlVef.iCode, tgVpf(ilBound)
        '        'Loop thru all vehicle that make up virtual vehicle and combine vpf data
        '        tlVsfSrchKey.lCode = tlVef.lVsfCode
        '        ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        '        If ilRet = BTRV_ERR_NONE Then
        '            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        '                If tmVsf.iFSCode(ilLoop) > 0 Then
        '                    For ilVpf = 0 To ilVpfMax Step 1
        '                        If tmVsf.iFSCode(ilLoop) = tgVpf(ilVpf).iVefKCode Then
        '
        '                            If tgVpf(ilVpf).sGPriceStat = "Y" Then
        '                                tgVpf(ilBound).sGPriceStat = "Y"
        '                            End If
        '                            If tgVpf(ilVpf).sSVarComm = "Y" Then
        '                                tgVpf(ilBound).sSVarComm = "Y"
        '                            End If
        '                            If tgVpf(ilVpf).sSPTA = "A" Then
        '                                tgVpf(ilBound).sSPTA = "A"
        '                            End If
        '                            'Build Libraries into rate card
        '                            tgVpf(ilBound).sSIntoRC = "N"
        '                            'Spot length
        '                            'Last Log Date
        '                            If (tgVpf(ilBound).iLLD(0) = 0) And (tgVpf(ilBound).iLLD(1) = 0) Then
        '                                tgVpf(ilBound).iLLD(0) = tgVpf(ilVpf).iLLD(0)
        '                                tgVpf(ilBound).iLLD(1) = tgVpf(ilVpf).iLLD(1)
        '                            Else
        '                                If (tgVpf(ilVpf).iLLD(0) <> 0) Or (tgVpf(ilVpf).iLLD(1) <> 0) Then
        '                                    gUnpackDateLong tgVpf(ilBound).iLLD(0), tgVpf(ilBound).iLLD(1), llDate1
        '                                    gUnpackDateLong tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), llDate2
        '                                    If llDate2 > llDate1 Then
        '                                        tgVpf(ilBound).iLLD(0) = tgVpf(ilVpf).iLLD(0)
        '                                        tgVpf(ilBound).iLLD(1) = tgVpf(ilVpf).iLLD(1)
        '                                    End If
        '                                End If
        '                            End If
        '                            If tgVpf(ilVpf).sSCommCalc = "C" Then
        '                                tgVpf(ilBound).sSCommCalc = "C"
        '                            End If
        '                            Exit For
        '                        End If
        '                    Next ilVpf
        '                End If
        '            Next ilLoop
        '        End If
        '        ilBound = ilBound + 1
        '    End If
        '    ilRet = btrGetNext(hlVef, tlVef, ilVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Loop
        'If ilRet <> BTRV_ERR_END_OF_FILE Then
        '    On Error GoTo gVpfReadErr
        '    gBtrvErrorMsg ilRet, "gVpfRead (btrGetFirst or Next):" & "Vpf.Btr", Frm
        '    On Error GoTo 0
        'End If

        ilRet = btrClose(hlVpf)
        btrDestroy hlVpf
'        ilRet = btrClose(hlVef)
'        btrDestroy hlVef
        'ilRet = btrClose(hlVsf)
        'btrDestroy hlVsf
        'Sort by vefCode
        If UBound(tgVpf) > 0 Then
            ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0
        End If
'        ilRet = csiSetStamp("VPF", sgVpfStamp)
'        ilRet = csiSetAlloc("VPF", LBound(tgVpf), UBound(tgVpf))
'        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
'            ilRet = csiSetRec("VPF", ilLoop, VarPtr(tgVpf(ilLoop)), LenB(tgVpf(ilLoop)))
'        Next ilLoop
        gVpfRead = True
    End If
    Exit Function

    ilRet = btrClose(hlVpf)
    btrDestroy hlVpf
'    ilRet = btrClose(hlVef)
'    btrDestroy hlVef
    'ilRet = btrClose(hlVsf)
    'btrDestroy hlVsf
    gVpfRead = False
    gDbg_HandleError "FileSubs: gVpfRead"
    Exit Function
gVpfReadErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gVpfInit                        *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize the Vehicle Option   *
'*                                                     *
'*******************************************************
Private Sub mVpfInit(ilVefCode As Integer, tlVpf As VPF)
    Dim ilLoop As Integer
    Dim ilRet As Integer

    tlVpf.iVefKCode = ilVefCode  'Input Vehicle
    tlVpf.iGTime(0) = 0     'Default to 12M
    tlVpf.iGTime(1) = 0
    tlVpf.sGMedium = "N"    'Radio Network
    'tlVpf.sFeedLogOrder = "N"   'Neither
    tlVpf.sUsingFeatures1 = Chr(0)
    tlVpf.iUrfGCode = 1     'Counterpoint
    tlVpf.sAdvtSep = "B"    'Mary request 7/12/98, "T"  'Separate advertiser by time
    tlVpf.sCPLogo = ""
    'tlVpf.iGMoFull = 3      '3 month retain unpacked
    'tlVpf.iGMoPack = 24     '24 months retain packed
    'For ilLoop = LBound(tlVpf.iGLink) To UBound(tlVpf.iGLink) Step 1
    '    tlVpf.iGLink(ilLoop) = 0
    'Next ilLoop
    tlVpf.lLgHd1CefCode = 0
    tlVpf.lLgNmCefCode = 0
    tlVpf.sUnsoldBlank = "Y"
    tlVpf.iSAGroupNo = 0
    'tlVpf.sGPriceStat = "N" 'Don't compute spot price statisitics
    tlVpf.sOwnership = "A" 'A=Owned-Network
    tlVpf.sGGridRes = "F"   'default grid resolution is Full hour
    'tlVpf.sGScript = "N"    'Using scripts
    tlVpf.sExpHiCorp = "Y"  'Export Corporate
    For ilLoop = LBound(tlVpf.sGZone) To UBound(tlVpf.sGZone) Step 1
        tlVpf.sGZone(ilLoop) = ""   'Zone name
        tlVpf.iGLocalAdj(ilLoop) = 0    'Local adjustment
        tlVpf.iGFeedAdj(ilLoop) = 0
        tlVpf.iGV1Z(ilLoop) = 0
        tlVpf.iGV2Z(ilLoop) = 0
        tlVpf.iGV3Z(ilLoop) = 0
        tlVpf.iGV4Z(ilLoop) = 0
        tlVpf.sGCSVer(ilLoop) = ""   'Don't Create log for time zone
        tlVpf.iGMnfNCode(ilLoop) = 0
        tlVpf.sGBus(ilLoop) = ""   'Don't Create log for time zone
        tlVpf.sGSked(ilLoop) = ""   'Don't Create log for time zone
        tlVpf.sGFed(ilLoop) = ""
    Next ilLoop
    tlVpf.sSVarComm = "N"    'Agency commission fixed
    tlVpf.sSCompType = "B"  'Separate competitive by break
    tlVpf.iSCompLen(0) = 0  '0 for competitive time separation
    tlVpf.iSCompLen(1) = 0
    tlVpf.iSBBLen = 5       'Standard BB length
    tlVpf.sSSellOut = "B"   'Sellout by Units/Seconds
    tlVpf.sSOverBook = "N"  'Disallow overbooking
    tlVpf.sSForceMG = "A"   'Set spots moved outside limit as MG's
    '7/15/14
    tlVpf.sEmbeddedOrROS = "R"  'Default Delivery for agreements:E= Embedded or R = ROS.  Test for E. Blank should be treated same as R
    tlVpf.sWegenerExport = "N"   'Wegener Export
    tlVpf.sOLAExport = "N"   'OLA Export (OnLine Affidavit)
    'tlVpf.sSIntoRC = "N"    'Unused-'Automatically insert libraries into Rate Cards
    tlVpf.sAvailNameOnWeb = "N" 'Show Avail Names on Web (Y/N)
    'tlVpf.sSPTA = "A"       'Unused-'Keep spots with program, time, ask
    tlVpf.sUsingFeatures2 = Chr(0)
    tlVpf.sSAvailOrder = "7"    'Avail order= N/A
    For ilLoop = LBound(tlVpf.iSLen) To UBound(tlVpf.iSLen) Step 1
        tlVpf.iSLen(ilLoop) = 0 'tgSpf.iSLen(ilLoop)      'Spot lengths
    Next ilLoop
    tlVpf.iSLen(LBound(tlVpf.iSLen)) = 30
    tlVpf.iSLen(LBound(tlVpf.iSLen) + 1) = 60
    For ilLoop = 0 To 9     'Length group number
        tlVpf.iSLenGroup(ilLoop) = 0
    Next ilLoop
    tlVpf.sSCommCalc = "B"  'Calculate salesperson commission on collected amount
    tlVpf.lMPSA60 = 0       'Last line used to assign 60 PSA
    tlVpf.lMPSA30 = 0
    tlVpf.lMPSA10 = 0
    tlVpf.lMPromo60 = 0     'Last line used to assign 60 Promo
    tlVpf.lMPromo30 = 0
    tlVpf.lMPromo10 = 0
    'For ilLoop = 1 To 24 Step 1
    For ilLoop = 0 To 23 Step 1
        tlVpf.iMMFPSA(ilLoop) = 0
        tlVpf.iMSAPSA(ilLoop) = 0
        tlVpf.iMSUPSA(ilLoop) = 0
        tlVpf.iMMFPromo(ilLoop) = 0
        tlVpf.iMSAPromo(ilLoop) = 0
        tlVpf.iMSUPromo(ilLoop) = 0
    Next ilLoop
    tlVpf.iLLD(0) = 0       'No date- set to zero
    tlVpf.iLLD(1) = 0
    tlVpf.iLPD(0) = 0       'No date- set to zero
    tlVpf.iLPD(1) = 0
    tlVpf.slTimeZone = "E"  'Time zone = pacific
    tlVpf.sLDayLight = "N"  'Daylight saving time
    tlVpf.sLTiming = "N"    'Using log timing
    tlVpf.sLAvailLen = "Y"  'Show length of unsold avails
    If (tlVpf.sGMedium = "R") Or (tlVpf.sGMedium = "N") Then
        tlVpf.iSDLen = 30   '60
    Else
        tlVpf.iSDLen = 30
    End If
    'tlVpf.sLCP = "Y"        'Unused'Create promo avails
    'tlVpf.sLBBs = "Y"       'Unused-'Create BBs avails
    tlVpf.iFTPArfCode = 0   'FTP associated with Vehicle (Audio Stored address so that affiliate can get the Audio)
    'tlVpf.sAffCPs = "N"       'Create Affiliate Custom CPs
    'tlVpf.sAffTimes = "N"        'Post Affiliate Exact Times
    tlVpf.sLShowCut = "N"    'Show cut/instruction # or neither
    tlVpf.sLTimeFormat = "A"    'Time format AM/PM
    tlVpf.slZone = "A"      'All Zones
    tlVpf.sCPTitle = ""
    tlVpf.sPrtCPStation = "N"
    tlVpf.iRnfPlayCode = 0  'default Certification report
    tlVpf.iLastCP(0) = 0       'No date- set to zero
    tlVpf.iLastCP(1) = 0
    tlVpf.sStnFdCart = "N"
    tlVpf.sStnFdXRef = "Y"
    tlVpf.sGenLog = "Y" 'Generate Log
    tlVpf.sCopyOnAir = "N"  'Allow copy on airing vehicle
    tlVpf.sBillSA = "S"        'Bill Selling or Airing spots for Selling vehicle
    tlVpf.sExpVehNo = ""
    tlVpf.sExpBkCpyCart = "N"
    tlVpf.sExpHiCmmlChg = "N"
    tlVpf.lAPenny = 0
    tlVpf.sGGroupNo = ""
    'tlVpf.iACal(0) = 0      'Field moved to Site; No date for last calendar date
    'tlVpf.iACal(1) = 0
    'tlVpf.iAStd(0) = 0      'Field moved to Site; No date for last standard date
    'tlVpf.iAStd(1) = 0
    'tlVpf.sAGL = ""         'Field removed; G/L prefix
    'slStr = ""
    'gStrToPDN slStr, 2, 3, tlVpf.sAPenny
    tlVpf.iLLastDateCpyAsgn(0) = 0
    tlVpf.iLLastDateCpyAsgn(1) = 0
    'For ilLoop = 1 To 5 Step 1
    For ilLoop = 0 To 4 Step 1
        tlVpf.iESTEndTime(ilLoop) = 0
        tlVpf.iCSTEndTime(ilLoop) = 0
        tlVpf.iMSTEndTime(ilLoop) = 0
        tlVpf.iPSTEndTime(ilLoop) = 0
    Next ilLoop
    'For ilLoop = 1 To 4 Step 1
    For ilLoop = 0 To 3 Step 1
        tlVpf.sMapZone(ilLoop) = ""
        tlVpf.sMapProgCode(ilLoop) = ""
        tlVpf.iMapDPNo(ilLoop) = 0
    Next ilLoop
    tlVpf.sExpHiClear = "N"
    tlVpf.sExpHiDallas = "N"
    tlVpf.sExpHiPhoenix = "N"
    tlVpf.sExpHiNY = "N"
    tlVpf.sExpHiNYISCI = "N"
    tlVpf.sBulkXFer = "N"
    tlVpf.sClearAsSell = "N"
    tlVpf.sClearChgTime = "N"
    tlVpf.sMoveLLD = "Y"
    tlVpf.iRnfLogCode = 0   'default Log report
    tlVpf.iRnfCertCode = 0  'default Certification report
    tlVpf.iLNoDaysCycle = 7  'Number of cycle days
    tlVpf.iLLeadTime = 8     'Lead time
    '1/13/16: Default Sports to 1 day
    ilRet = gBinarySearchVef(ilVefCode)
    If ilRet <> -1 Then
        If tgMVef(ilRet).sType = "G" Then
            tlVpf.iLNoDaysCycle = 1  'Number of cycle days
            tlVpf.iLLeadTime = 3     'Lead time
        End If
    End If
    tlVpf.sShowTime = "S"
    tlVpf.sEDICallLetters = ""
    tlVpf.sAccruedRevenue = ""
    tlVpf.sAccruedTrade = ""
    tlVpf.sBilledRevenue = ""
    tlVpf.sBilledTrade = ""
    tlVpf.sLCmmlSmmyAvNm = ""
    tlVpf.lEDASWindow = 400
    tlVpf.sKCGenRot = "Y"
    tlVpf.sExportSQL = "N"
    'tlVpf.sUnused1 = ""
    tlVpf.iLastLog(0) = 0       'No date- set to zero
    tlVpf.iLastLog(1) = 0
    tlVpf.iRnfSvLogCode = 0
    tlVpf.iRnfSvCertCode = 0
    tlVpf.iRnfSvPlayCode = 0
    tlVpf.lLgFt1CefCode = 0
    tlVpf.lLgFt2CefCode = 0
    tlVpf.sStnFdCode = ""
    tlVpf.iProducerArfCode = 0
    tlVpf.iProgProvArfCode = 0
    tlVpf.iCommProvArfCode = 0
    tlVpf.sEmbeddedComm = "N"
    tlVpf.sARBCode = ""
    tlVpf.lEMailCefCode = 0
    tlVpf.sShowRateOnInsert = "N"
    tlVpf.sAllowSplitCopy = "N"
    tlVpf.iAutoExptArfCode = 0
    tlVpf.iAutoImptArfCode = 0
    tlVpf.sWebLogSummary = ""
    tlVpf.sWebLogFeedTime = ""
    tlVpf.sRadarCode = ""
    tlVpf.sEDIBand = ""
    tlVpf.iInterfaceID = 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gVffRead                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Read the Vehicle Option file   *
'*                                                     *
'*******************************************************
Function gVffRead() As Integer
'
'   ilRet = gVffRead ()
'   Where:
'       ilRet (O)- True if read
'                       False if none read
'

    Dim ilRecLen As Integer     'Vff record length
    Dim hlVff As Integer     'Vehicle Option file handle
    Dim hlVef As Integer     'Vehicle Option file handle
    Dim hlVpf As Integer     'Vehicle Option file handle
    Dim tlVpf As VPF                'VPF record image
    Dim tlVpfSrchKey As VPFKEY0     'VPF key 0 image
    Dim ilVpfRecLen As Integer      'VPF record length
    Dim slStamp As String
    Dim llNoRec As Long      'Number of records (must be one)
    Dim llNoVffRec As Long
    Dim ilRet As Integer
    Dim ilBound As Integer
    Dim ilVffMax As Integer
    Dim llRecPos As Long
    Dim slGetStamp As String
    Dim ilVef As Integer
    Dim ilAdd As Integer
    Dim ilLoop As Integer
    Dim ilZone As Integer
    Dim ilVpf As Integer
    Dim ilVff As Integer
    Dim ilCode As Integer
    Dim ilVffRecLen As Integer      'MNF record length
    Dim tlVffSrchKey0 As INTKEY0
    Dim tlVffSrchKey1 As INTKEY0
    Dim tlVff As VFF
    Dim blUpdate As Boolean
    ReDim ilTransferVefCode(0 To 0) As Integer
    

    slStamp = gFileDateTime(sgDBPath & "Vff.btr")

    'On Error GoTo gVffReadErr2
    'ilRet = 0
    'imLowLimit = LBound(tgVff)
    'If ilRet <> 0 Then
    '    sgVffStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgVff).Ptr <> 0 Then
        imLowLimit = LBound(tgVff)
    Else
        sgVffStamp = ""
        imLowLimit = 0
    End If

    If (sgVffStamp <> "") And (sgVffStamp <> "~") Then
        'If slStamp = sgVffStamp Then
            gVffRead = True
            Exit Function
        'End If
        'slGetStamp = ""
    Else
        slGetStamp = "" 'gGetCSIStamp("Vff")
        igVehicleDefined = False    'True
    End If
    If slGetStamp <> "" Then
        gVffRead = True
    Else
        sgVffStamp = slStamp
        hlVff = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlVff, "", sgDBPath & "Vff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gVffRead = False
            btrDestroy hlVff
            Exit Function
        End If
        hlVef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gVffRead = False
            btrDestroy hlVff
            Exit Function
        End If
        llNoVffRec = btrRecords(hlVff) 'Obtain number of records
        If llNoVffRec <= btrRecords(hlVef) Then
            llNoVffRec = btrRecords(hlVef) 'Obtain number of records
        End If
        btrDestroy hlVef
        If llNoVffRec <= 0 Then
            ReDim tgVff(0 To 0) As VFF
            ilRet = btrClose(hlVff)
            btrDestroy hlVff
            gVffRead = False
            igVehicleDefined = False
            Exit Function
        End If
        igVehicleDefined = True
        ReDim tgVff(0 To llNoVffRec - 1) As VFF
        ilRecLen = Len(tgVff(0)) 'Get and save record length
        llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hlVff   'Clear any previous extend operation
        ilBound = 0
        ilRet = btrGetFirst(hlVff, tgVff(ilBound), ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            Call btrExtSetBounds(hlVff, llNoRec, -1, "UC", "Vff", "") 'Set extract limits (all records)
            ilRet = btrExtAddField(hlVff, 0, ilRecLen)  'Extract record
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrExtGetNext(hlVff, tgVff(ilBound), ilRecLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                        ilRecLen = Len(tgVff(0)) 'Get and save record length
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlVff, tgVff(ilBound), ilRecLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            If ilBound >= llNoVffRec - 1 Then
                                Exit Do
                            End If
                            If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                                ilAdd = True
                            Else
                                ilVef = gBinarySearchVef(tgVff(ilBound).iVefCode)
                                If ilVef <> -1 Then
                                    ilAdd = True
                                Else
                                    ilAdd = False
                                End If
                            End If
                            If ilAdd Then
                                ilVpf = gBinarySearchVpf(tgVff(ilBound).iVefCode)
                                If ilVpf <> -1 Then
                                    For ilZone = LBound(tgVpf(ilVpf).sGFed) To UBound(tgVpf(ilVpf).sGFed) Step 1
                                        If (tgVpf(ilVpf).sGFed(ilZone) = "N") Or (tgVpf(ilVpf).sGFed(ilZone) = "Y") Then
                                            ilTransferVefCode(UBound(ilTransferVefCode)) = tgVff(ilBound).iVefCode
                                            ReDim Preserve ilTransferVefCode(0 To UBound(ilTransferVefCode) + 1) As Integer
                                            Exit For
                                        End If
                                    Next ilZone
                                End If
                                ilBound = ilBound + 1
                                ilRet = btrExtGetNext(hlVff, tgVff(ilBound), ilRecLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hlVff, tgVff(ilBound), ilRecLen, llRecPos)
                                Loop
                            End If
                        Loop
                    End If
                End If
            End If
        End If
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gVffRead = False
                btrDestroy hlVff
                Exit Function
            End If
        End If
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilVffMax = ilBound - 1
        Else
            ilVffMax = ilBound
            ilBound = ilBound + 1
        End If
        ReDim Preserve tgVff(0 To ilVffMax) As VFF

        '5/11/12: Move Delivery fed from vpf to vff
        If UBound(ilTransferVefCode) > LBound(ilTransferVefCode) Then
            ilVffRecLen = Len(tlVff)
            ilVpfRecLen = Len(tlVpf)
            hlVpf = CBtrvTable(ONEHANDLE)
            ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                For ilLoop = LBound(ilTransferVefCode) To UBound(ilTransferVefCode) - 1 Step 1
                    tlVpfSrchKey.iVefKCode = ilTransferVefCode(ilLoop)
                    ilRet = btrGetEqual(hlVpf, tlVpf, ilVpfRecLen, tlVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tlVffSrchKey1.iCode = ilTransferVefCode(ilLoop)
                        ilRet = btrGetEqual(hlVff, tlVff, ilVffRecLen, tlVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            For ilZone = LBound(tlVpf.sGFed) To UBound(tlVpf.sGFed) Step 1
                                tlVff.sFedDelivery(ilZone) = tlVpf.sGFed(ilZone)
                                tlVpf.sGFed(ilZone) = ""
                            Next ilZone
                            ilRet = btrUpdate(hlVpf, tlVpf, ilVpfRecLen)
                            ilRet = btrUpdate(hlVff, tlVff, ilVffRecLen)
                            ilVpf = gBinarySearchVpf(tlVpf.iVefKCode)
                            If ilVpf <> -1 Then
                                tgVpf(ilVpf) = tlVpf
                            End If
                            '11/26/17
                            gFileChgdUpdate "vpf.btr", False

                            'For ilVff = LBound(tgVff) To UBound(tgVff) - 1 Step 1
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If tgVff(ilVff).iVefCode = ilTransferVefCode(ilLoop) Then
                                    tgVff(ilVff) = tlVff
                                    Exit For
                                End If
                            Next ilVff
                        End If
                    End If
                Next ilLoop
            End If
            ilRet = btrClose(hlVpf)
            btrDestroy hlVpf
        End If
        'If UBound(tgVff) > 0 Then
        '    ArraySortTyp fnAV(tgVff(), 0), UBound(tgVff) + 1, 0, LenB(tgVff(0)), 0, -1, 0
        'End If
        gVffRead = True
    End If
    If UBound(tgVff) > 0 Then
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            ilCode = tgVff(ilVff).iCode
            tgVff(ilVff).iCode = tgVff(ilVff).iVefCode
            tgVff(ilVff).iVefCode = ilCode
        Next ilVff
        ArraySortTyp fnAV(tgVff(), 0), UBound(tgVff) + 1, 0, LenB(tgVff(0)), 0, -1, 0
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            ilCode = tgVff(ilVff).iCode
            tgVff(ilVff).iCode = tgVff(ilVff).iVefCode
            tgVff(ilVff).iVefCode = ilCode
        Next ilVff
        ilVffRecLen = Len(tlVff)
        If ((Trim$(tgSite.sNoMissedReason) = "") Or (Trim$(tgSite.sNoMissedReason) = "N")) And ((Trim$(tgSite.sAllowMGSpots) = "") Or (Trim$(tgSite.sAllowMGSpots) = "Y")) Then
            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                blUpdate = False
                If tgVff(ilVff).sMGsOnWeb <> "Y" And tgVff(ilVff).sMGsOnWeb <> "N" Then
                    tgVff(ilVff).sMGsOnWeb = "N"    'Y"
                    blUpdate = True
                End If
                If tgVff(ilVff).sReplacementOnWeb <> "Y" And tgVff(ilVff).sReplacementOnWeb <> "N" Then
                    tgVff(ilVff).sReplacementOnWeb = "N"    '"Y"
                    blUpdate = True
                End If
                If blUpdate = True Then
                    tlVffSrchKey0.iCode = tgVff(ilVff).iCode
                    ilRet = btrGetEqual(hlVff, tlVff, ilVffRecLen, tlVffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrUpdate(hlVff, tgVff(ilVff), ilVffRecLen)
                    End If
                End If
            Next ilVff
        End If
    End If
    On Error Resume Next
    ilRet = btrClose(hlVff)
    btrDestroy hlVff
    Exit Function
gVffReadErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertAdd                       *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Alert if not previously     *
'*                     added                           *
'*                                                     *
'*******************************************************
Public Function gAlertAdd(slType As String, slSubType As String, llChfCode As Long, ilVefCode As Integer, slInWeekDate As String) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim ilVpf As Integer
    Dim slTestType As String
    Dim slUlfCode As String
    Dim slCefCode As String
    Dim slCountdown As String

    If (tgSpf.sGUseAffSys <> "Y") And ((slType = "F") Or (slType = "R") Or (slType = "P")) Then
        gAlertAdd = BTRV_ERR_NONE
        Exit Function
    End If
    
    '4/12/14: Bypass adding Log Alerts if vehicle is not generating Logs
    If (slType = "L") Then 'See if Log generated
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            If (tgVpf(ilVpf).sGenLog = "N") Then
                gAlertAdd = BTRV_ERR_NONE
                Exit Function
            End If
        End If
    End If
    
    If (slType <> "C") And (slType <> "B") And (slType <> "N") And (slType <> "M") Then
        llMoWeekDate = gDateValue(slInWeekDate)
        Do While gWeekDayLong(llMoWeekDate) <> 0
            llMoWeekDate = llMoWeekDate - 1
        Loop
    Else
        llMoWeekDate = 0
    End If

    ilFound = False
    imAufRecLen = Len(tmAuf)
    If (slType <> "N") And (slType <> "M") Then
        '7/15/14
        If slType = "L" Or slType = "F" Or slType = "R" Or slType = "P" Then
            tmAufSrchKey2.sType = slType
            tmAufSrchKey2.sStatus = "R"
            tmAufSrchKey2.sSubType = slSubType
            tmAufSrchKey2.iVefCode = ilVefCode
            gPackDateLong llMoWeekDate, tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
            If slType <> "P" Then
                gPackDateLong llMoWeekDate, tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            Else
                gPackDate "1/1/1970", tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
                ilRet = btrGetGreaterOrEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            End If
        ElseIf slType = "C" Then
            tmAufSrchKey3.sType = slType
            tmAufSrchKey3.sStatus = "R"
            tmAufSrchKey3.sSubType = slSubType
            tmAufSrchKey3.lChfCode = llChfCode
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        Else
            tmAufSrchKey1.sType = slType
            tmAufSrchKey1.sStatus = "R"
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (slType = tmAuf.sType) And (tmAuf.sStatus = "R")
            If slType = "C" Then    'Contract
                If (llChfCode = tmAuf.lChfCode) And (slSubType = tmAuf.sSubType) Then
                    ilFound = True
                End If
                Exit Do
            ElseIf slType = "L" Then    'Log
                If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate = llMoWeekDate Then
                        ilFound = True
                    End If
                End If
                Exit Do
            ElseIf slType = "F" Then    'Affiliate Export-Final
                If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate = llMoWeekDate Then
                        ilFound = True
                    End If
                End If
                Exit Do
            ElseIf slType = "R" Then    'Affiliate Export-Reprint
                If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate = llMoWeekDate Then
                        ilFound = True
                    End If
                End If
                Exit Do
            ElseIf slType = "P" Then    'Program
                If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                    ilFound = True
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate > llMoWeekDate Then
                        gPackDateLong llMoWeekDate, tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1)
                        ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
                    End If
                    Exit Do
                End If
            ElseIf slType = "B" Then    'Block
                ilFound = True
                ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
                ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
                ilRet = gParseItem(slInWeekDate, 3, "|", slCountdown)
                tmAuf.lUlfCode = Val(slUlfCode)
                tmAuf.lCefCode = Val(slCefCode)
                tmAuf.iCountdown = Val(slCountdown)
                If tmAuf.iCountdown <= 0 Then
                    tmAuf.sStatus = "C"
                    tmAuf.sClearMethod = "M"
                    slDate = Format$(gNow(), "m/d/yy")
                    slTime = Format$(gNow(), "h:mm:ssAM/PM")
                    gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
                    gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
                    tmAuf.iClearUrfCode = tgUrf(0).iCode
                End If
                slDate = Format$(gNow(), "m/d/yy")
                slTime = Format$(gNow(), "h:mm:ssAM/PM")
                gPackDate slDate, tmAuf.iEnteredDate(0), tmAuf.iEnteredDate(1)
                gPackTime slTime, tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1)
                ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
                Exit Do
            End If
            ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
        Loop
    End If
    If Not ilFound Then
        If ((slType = "F") Or (slType = "R")) And (slSubType = "I") Then 'Check if provider defined, if not then don't add
            ilVpf = gBinarySearchVpf(ilVefCode)
            If ilVpf <> -1 Then
                If (tgVpf(ilVpf).iCommProvArfCode <= 0) Then
                    gAlertAdd = BTRV_ERR_NONE
                    Exit Function
                End If
            Else
                gAlertAdd = BTRV_ERR_NONE
                Exit Function
            End If
        ElseIf (slType = "L") Then 'See if Log generated
            ilVpf = gBinarySearchVpf(ilVefCode)
            If ilVpf <> -1 Then
                If (tgVpf(ilVpf).sGenLog = "N") Then
                    gAlertAdd = BTRV_ERR_NONE
                    Exit Function
                End If
            Else
                gAlertAdd = BTRV_ERR_NONE
                Exit Function
            End If
        End If
        '6/12/06: TTP 1576
        'Check if adding Final and Reprint alert already exist, if so, remove Reprint
        'Check if adding Reprint and Final alert already exist, if so, ignore the add
        If (slType = "F") Or (slType = "R") Then
            If slType = "F" Then
                slTestType = "R"
            Else
                slTestType = "F"
            End If
            'tmAufSrchKey1.sType = slTestType
            'tmAufSrchKey1.sStatus = "R"
            'ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            '7/15/14
            tmAufSrchKey2.sType = slTestType
            tmAufSrchKey2.sStatus = "R"
            tmAufSrchKey2.sSubType = slSubType
            tmAufSrchKey2.iVefCode = ilVefCode
            gPackDateLong llMoWeekDate, tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = slTestType) And (tmAuf.sStatus = "R")
                If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                    gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                    If llDate = llMoWeekDate Then
                        ilFound = True
                        Exit Do
                    End If
                End If
                ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
            Loop
            If ilFound Then
                If slTestType = "R" Then
                    ilRet = btrDelete(hgAuf)
                Else
                    gAlertAdd = BTRV_ERR_NONE
                    Exit Function
                End If
            End If
        End If
        tmAuf.lCode = 0
        slDate = Format$(gNow(), "m/d/yy")
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackDate slDate, tmAuf.iEnteredDate(0), tmAuf.iEnteredDate(1)
        gPackTime slTime, tmAuf.iEnteredTime(0), tmAuf.iEnteredTime(1)
        tmAuf.sStatus = "R"
        tmAuf.sType = slType
        tmAuf.sSubType = slSubType
        tmAuf.lChfCode = llChfCode
        tmAuf.iVefCode = ilVefCode
        gPackDateLong llMoWeekDate, tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1)
        tmAuf.iCreateUrfCode = tgUrf(0).iCode
        tmAuf.sClearMethod = ""
        tmAuf.iClearUrfCode = 0
        tmAuf.iClearUstCode = 0
        gPackDate "", tmAuf.iClearDate(0), tmAuf.iClearDate(1)
        gPackTime "", tmAuf.iClearTime(0), tmAuf.iClearTime(1)
        tmAuf.lUlfCode = 0
        tmAuf.lCefCode = 0
        tmAuf.iCountdown = 0
        If slType = "B" Then    'Block
            ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
            ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
            ilRet = gParseItem(slInWeekDate, 3, "|", slCountdown)
            tmAuf.lUlfCode = Val(slUlfCode)
            tmAuf.lCefCode = Val(slCefCode)
            tmAuf.iCountdown = Val(slCountdown)
        ElseIf slType = "N" Then    'Notification
            ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
            ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
            tmAuf.lUlfCode = Val(slUlfCode)
            tmAuf.lCefCode = Val(slCefCode)
        ElseIf slType = "M" Then    'Message
            slCefCode = Trim$(slInWeekDate)
            tmAuf.lUlfCode = 0
            tmAuf.lCefCode = Val(slCefCode)
        End If
        tmAuf.sUnused = ""
        ilRet = btrInsert(hgAuf, tmAuf, imAufRecLen, INDEXKEY0)
        gAlertAdd = ilRet
        ilRet = gAlertForceCheck()
    Else
        gAlertAdd = BTRV_ERR_NONE
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertClear                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Alert previously added    *
'*                                                     *
'*******************************************************
Public Function gAlertClear(slMethod As String, slType As String, slSubType As String, llChfCode As Long, ilVefCode As Integer, slInWeekDate As String)
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim llAufCode As Long

    If (slType <> "C") And (slType <> "B") And (slType <> "N") Then
        llMoWeekDate = gDateValue(slInWeekDate)
        Do While gWeekDayLong(llMoWeekDate) <> 0
            llMoWeekDate = llMoWeekDate - 1
        Loop
    Else
        llMoWeekDate = 0
    End If
    If slType = "N" Then
        llAufCode = Val(slInWeekDate)
    Else
        llAufCode = 0
    End If

    ilFound = False
    imAufRecLen = Len(tmAuf)
    '7/15/14
    If slType = "L" Or slType = "P" Then
        tmAufSrchKey2.sType = slType
        tmAufSrchKey2.sStatus = "R"
        tmAufSrchKey2.sSubType = slSubType
        tmAufSrchKey2.iVefCode = ilVefCode
        gPackDateLong llMoWeekDate, tmAufSrchKey2.iMoWeekDate(0), tmAufSrchKey2.iMoWeekDate(1)
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    ElseIf slType = "C" Then
        tmAufSrchKey3.sType = slType
        tmAufSrchKey3.sStatus = "R"
        tmAufSrchKey3.sSubType = slSubType
        tmAufSrchKey3.lChfCode = llChfCode
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    Else
        tmAufSrchKey1.sType = slType
        tmAufSrchKey1.sStatus = "R"
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (slType = tmAuf.sType) And (tmAuf.sStatus = "R")
        If slType = "C" Then
            If (llChfCode = tmAuf.lChfCode) And (slSubType = tmAuf.sSubType) Then
                ilFound = True
            End If
            Exit Do
        ElseIf slType = "L" Then
            If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                If llDate = llMoWeekDate Then
                    ilFound = True
                End If
            End If
            Exit Do
        ElseIf slType = "P" Then
            If (ilVefCode = tmAuf.iVefCode) And (slSubType = tmAuf.sSubType) Then
                gUnpackDateLong tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), llDate
                If llDate = llMoWeekDate Then
                    ilFound = True
                End If
            End If
            Exit Do
        ElseIf slType = "B" Then
            ilFound = True
            Exit Do
        ElseIf slType = "N" Then
            If tmAuf.lCode = llAufCode Then
                ilFound = True
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get next record
    Loop
    If ilFound Then
        tmAuf.sStatus = "C"
        tmAuf.sClearMethod = slMethod
        slDate = Format$(gNow(), "m/d/yy")
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
        gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
        tmAuf.iClearUrfCode = tgUrf(0).iCode
        ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
        gAlertClear = True
    Else
        gAlertClear = False
    End If
    ilRet = gAlertForceCheck()
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertContractErase             *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Erase Alerts for Contract       *
'*                                                     *
'*******************************************************
Public Function gAlertContractErase(llChfCode As Long) As Integer
    Dim ilRet As Integer
    Dim slType As String
    Dim ilLoop As Integer
    ReDim llAufCode(0 To 0) As Long

    slType = "C"
    imAufRecLen = Len(tmAuf)
    tmAufSrchKey1.sType = slType
    tmAufSrchKey1.sStatus = ""
    ilRet = btrGetGreaterOrEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (slType = tmAuf.sType)
        If (llChfCode = tmAuf.lChfCode) Then
            llAufCode(UBound(llAufCode)) = tmAuf.lCode
            ReDim Preserve llAufCode(0 To UBound(llAufCode) + 1) As Long
        End If
        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get next record
    Loop
    For ilLoop = 0 To UBound(llAufCode) - 1 Step 1
        tmAufSrchKey0.lCode = llAufCode(ilLoop)
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgAuf)
        End If
    Next ilLoop
    ilRet = gAlertForceCheck()
    gAlertContractErase = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertVehicleReplace            *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Replace selling vehicles with   *
'*                     airing vehicles                 *
'*                                                     *
'*******************************************************
Public Sub gAlertVehicleReplace(hlVlf As Integer)
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilAir As Integer
    Dim slDate As String
    Dim slTime As String
    Dim tlVef As VEF
    ReDim llAufCode(0 To 0) As Long
    ReDim ilVefCode(0 To 0) As Integer

    imAufRecLen = Len(tmAuf)
    tmAufSrchKey1.sType = "L"
    tmAufSrchKey1.sStatus = "R"
    ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "L") And (tmAuf.sStatus = "R")
        ilVef = gBinarySearchVef(tmAuf.iVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType = "S" Then
                llAufCode(UBound(llAufCode)) = tmAuf.lCode
                ReDim Preserve llAufCode(0 To UBound(llAufCode) + 1) As Long
            End If
        End If
        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get next record
    Loop
    For ilLoop = 0 To UBound(llAufCode) - 1 Step 1
        tmAufSrchKey0.lCode = llAufCode(ilLoop)
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            'Replace selling by airing
            ilVef = gBinarySearchVef(tmAuf.iVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "S" Then
                    gUnpackDate tmAuf.iMoWeekDate(0), tmAuf.iMoWeekDate(1), slDate
                    tlVef = tgMVef(ilVef)
                    gBuildLinkArray hlVlf, tlVef, slDate, ilVefCode()
                    For ilAir = 0 To UBound(ilVefCode) - 1 Step 1
                        tmAuf.lCode = 0
                        tmAuf.iVefCode = ilVefCode(ilAir)
                        ilRet = gAlertFound(tmAuf.sType, tmAuf.sSubType, tmAuf.lChfCode, tmAuf.iVefCode, slDate)
                        If Not ilRet Then
                            ilRet = btrInsert(hgAuf, tmAuf, imAufRecLen, INDEXKEY0)
                        End If
                    Next ilAir
                End If
            End If
        End If
        'Taken from gAlertClear
        tmAufSrchKey0.lCode = llAufCode(ilLoop)
        ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmAuf.sStatus = "C"
            tmAuf.sClearMethod = "A"
            slDate = Format$(gNow(), "m/d/yy")
            slTime = Format$(gNow(), "h:mm:ssAM/PM")
            gPackDate slDate, tmAuf.iClearDate(0), tmAuf.iClearDate(1)
            gPackTime slTime, tmAuf.iClearTime(0), tmAuf.iClearTime(1)
            tmAuf.iClearUrfCode = tgUrf(0).iCode
            ilRet = btrUpdate(hgAuf, tmAuf, imAufRecLen)
        End If
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertCheck                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if Alerts should be   *
'*                     turn on or off                  *
'*                                                     *
'*******************************************************
Public Function gAlertCheck() As Integer
    Dim ilRet As Integer
    Dim ilCheckCntr As Integer
    Dim tlChf As CHF
    Dim ilChfRecLen As Integer
    Dim tlChfSrchKey As LONGKEY0

    If igBkgdProg <> 0 Then
        gAlertCheck = False
        Exit Function
    End If
    If (tgSpf.iGAlertInterval <> 0) And (igSportsSystem <> 2) Then
        If ((tgSpf.iGAlertInterval <> 0) And (tgSpf.iGAlertInterval <= igAlertTimer)) Then
            'Check if Alert exist
            imAufRecLen = Len(tmAuf)
            '8/15/18: added Unapproved alert
            If (tgUrf(0).sCompAlert <> "N") Or (tgUrf(0).sSchAlert <> "N") Or (tgUrf(0).sIncompAlert <> "N") Then
                ilChfRecLen = Len(tlChf)
                tmAufSrchKey1.sType = "C"
                tmAufSrchKey1.sStatus = "R"
                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "C") And (tmAuf.sStatus = "R")
                    tlChfSrchKey.lCode = tmAuf.lChfCode
                    ilRet = btrGetEqual(hgChf, tlChf, ilChfRecLen, tlChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If ((tlChf.sStatus = tmAuf.sSubType) And (tgUrf(0).sCompAlert <> "N")) Then
                            ilCheckCntr = True
                            If (tmAuf.sSubType = "C") And (tgUrf(0).sReviseCntr = "N") And (tlChf.iCntRevNo > 0) Then
                                ilCheckCntr = False
                            End If
                            If ilCheckCntr Then
                                If gCntrOkForUser(hgVsf, tgUrf(0).iSlfCode, tlChf.lVefCode, tlChf.iSlfCode(), tlChf.sSource) Then
                                    Traffic!tmcCheckAlert.Enabled = False
                                    igAlertFlash = -1
                                    igAlertTimer = 0
                                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                                    Traffic!tmcFlashAlert.Enabled = True
                                    DoEvents
                                    gAlertCheck = True
                                    Exit Function
                                End If
                            End If
                        '8/15/18: added Unapproved alert
                        ElseIf ((tlChf.sStatus = tmAuf.sSubType) And (tgUrf(0).sIncompAlert <> "N")) Then
                            ilCheckCntr = True
                            'If (tmAuf.sSubType = "C") And (tgUrf(0).sReviseCntr = "N") And (tlChf.iCntRevNo > 0) Then
                            '    ilCheckCntr = False
                            'End If
                            If ilCheckCntr Then
                                If gCntrOkForUser(hgVsf, tgUrf(0).iSlfCode, tlChf.lVefCode, tlChf.iSlfCode(), tlChf.sSource) Then
                                    Traffic!tmcCheckAlert.Enabled = False
                                    igAlertFlash = -1
                                    igAlertTimer = 0
                                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                                    Traffic!tmcFlashAlert.Enabled = True
                                    DoEvents
                                    gAlertCheck = True
                                    Exit Function
                                End If
                            End If
                        '8/23/18: added Programmatic Buy alert
                        ElseIf (tmAuf.sSubType = "P") And ((tgUrf(0).sCompAlert <> "N") Or (tgUrf(0).sIncompAlert <> "N")) Then
                            ilCheckCntr = True
                            'If (tmAuf.sSubType = "C") And (tgUrf(0).sReviseCntr = "N") And (tlChf.iCntRevNo > 0) Then
                            '    ilCheckCntr = False
                            'End If
                            If ilCheckCntr Then
                                If gCntrOkForUser(hgVsf, tgUrf(0).iSlfCode, tlChf.lVefCode, tlChf.iSlfCode(), tlChf.sSource) Then
                                    Traffic!tmcCheckAlert.Enabled = False
                                    igAlertFlash = -1
                                    igAlertTimer = 0
                                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                                    Traffic!tmcFlashAlert.Enabled = True
                                    DoEvents
                                    gAlertCheck = True
                                    Exit Function
                                End If
                            End If
                        ElseIf ((tlChf.sSchStatus = "A") Or (tlChf.sSchStatus = "N")) And (tgUrf(0).sSchAlert <> "N") Then
                            ilCheckCntr = True
                            If (tmAuf.sSubType <> "S") Then
                                ilCheckCntr = False
                            End If
                            If ilCheckCntr Then
                                If gCntrOkForUser(hgVsf, tgUrf(0).iSlfCode, tlChf.lVefCode, tlChf.iSlfCode(), tlChf.sSource) Then
                                    Traffic!tmcCheckAlert.Enabled = False
                                    igAlertFlash = -1
                                    igAlertTimer = 0
                                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                                    Traffic!tmcFlashAlert.Enabled = True
                                    DoEvents
                                    gAlertCheck = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
                Loop
            End If
            If (tgUrf(0).sReprintLogAlert <> "N") Then
                tmAufSrchKey1.sType = "L"
                tmAufSrchKey1.sStatus = "R"
                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    Traffic!tmcCheckAlert.Enabled = False
                    igAlertFlash = -1
                    igAlertTimer = 0
                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                    Traffic!tmcFlashAlert.Enabled = True
                    DoEvents
                    gAlertCheck = True
                    Exit Function
                End If
            End If
            If (tgUrf(0).sShowNRMsg <> "N") Then
                tmAufSrchKey1.sType = "M"
                tmAufSrchKey1.sStatus = "R"
                ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    Traffic!tmcCheckAlert.Enabled = False
                    igAlertFlash = -1
                    igAlertTimer = 0
                    Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                    Traffic!tmcFlashAlert.Enabled = True
                    DoEvents
                    gAlertCheck = True
                    Exit Function
                End If
            End If
            tmAufSrchKey1.sType = "U"
            tmAufSrchKey1.sStatus = "R"
            ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                Traffic!tmcCheckAlert.Enabled = False
                igAlertFlash = -1
                igAlertTimer = 0
                Traffic!tmcFlashAlert.Interval = 2000   'Every 2 second
                Traffic!tmcFlashAlert.Enabled = True
                DoEvents
                gAlertCheck = True
                Exit Function
            End If
            Traffic!tmcFlashAlert.Enabled = False
            Traffic!tmcCheckAlert.Enabled = True
            Traffic!mnuAlert.Visible = False
            igAlertTimer = 1
        Else
            igAlertTimer = igAlertTimer + 1
        End If
    Else
        Traffic!tmcCheckAlert.Enabled = False
        Traffic!mnuAlert.Visible = False
    End If
    gAlertCheck = False
End Function
Public Function gAlertCheckBlock(tlAuf As AUF) As Integer
    Dim ilRet As Integer

    imAufRecLen = Len(tmAuf)
    tmAufSrchKey1.sType = "B"
    tmAufSrchKey1.sStatus = "R"
    ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "B") And (tmAuf.sStatus = "R") Then
        If (tmAuf.sSubType = "T") Or (tmAuf.sSubType = "B") Then
            tlAuf = tmAuf
            If (tmAuf.iCreateUrfCode <> tgUrf(0).iCode) Then
                gAlertCheckBlock = 1    'Block user
                Exit Function
            Else
                gAlertCheckBlock = 2    'User that initiate block
                Exit Function
            End If
        End If
    End If
    gAlertCheckBlock = 0    'No blocks
End Function
Public Function gAlertCheckNotice(llUlfCode As Long, llCefCode As Long) As Long
    Dim ilRet As Integer

    imAufRecLen = Len(tmAuf)
    tmAufSrchKey1.sType = "N"
    tmAufSrchKey1.sStatus = "R"
    ilRet = btrGetEqual(hgAuf, tmAuf, imAufRecLen, tmAufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmAuf.sType = "N") And (tmAuf.sStatus = "R")
        llCefCode = tmAuf.lCefCode
        If tmAuf.lUlfCode = llUlfCode Then
            gAlertCheckNotice = tmAuf.lCode
            Exit Function
        End If
        imUlfRecLen = Len(tmUlf)
        tmUlfSrchKey0.lCode = tmAuf.lUlfCode
        ilRet = btrGetEqual(hgUlf, tmUlf, imUlfRecLen, tmUlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If (ilRet = BTRV_ERR_NONE) Then
            If tgUrf(0).iCode = tmUlf.iUrfCode Then
                gAlertCheckNotice = tmAuf.lCode
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hgAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    gAlertCheckNotice = 0
End Function

Public Function gAlertForceCheck() As Integer
    If igBkgdProg <> 0 Then
        gAlertForceCheck = False
        Exit Function
    End If
    If tgSpf.iGAlertInterval <> 0 Then
        '8/15/18: Added unapproved
        If (tgUrf(0).sCompAlert <> "N") Or (tgUrf(0).sSchAlert <> "N") Or (tgUrf(0).sReprintLogAlert <> "N") Or (tgUrf(0).sShowNRMsg <> "N") Or (tgUrf(0).sIncompAlert <> "N") Then
            Traffic!tmcFlashAlert.Enabled = False
            Traffic!tmcCheckAlert.Enabled = False
            If tgSpf.iGAlertInterval > 0 Then
                igAlertTimer = tgSpf.iGAlertInterval
            End If
            gAlertForceCheck = gAlertCheck()
            igAlertFlash = -1
            Exit Function
        End If
    End If
    Traffic!mnuAlert.Visible = False
    gAlertForceCheck = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetFlightsFromLine             *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get flights for Line            *
'*                                                     *
'*******************************************************
Public Function gGetFlightsFromLine(hlCff As Integer, tlClf As CLF, tlCff() As CFF) As Integer
    Dim tlCffSrchKey As CFFKEY0 'CFF key record image
    Dim ilCffRecLen As Integer     'CFF record length
    Dim ilRet As Integer
    Dim ilUpper As Integer

    ReDim tlCff(0 To 0) As CFF
    ilUpper = 0
    tlCffSrchKey.lChfCode = tlClf.lChfCode
    tlCffSrchKey.iClfLine = tlClf.iLine
    tlCffSrchKey.iCntRevNo = tlClf.iCntRevNo
    tlCffSrchKey.iPropVer = tlClf.iPropVer
    tlCffSrchKey.iStartDate(0) = 0
    tlCffSrchKey.iStartDate(1) = 0
    ilCffRecLen = Len(tlCff(0))
    ilRet = btrGetGreaterOrEqual(hlCff, tlCff(ilUpper), ilCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlCff(ilUpper).lChfCode = tlClf.lChfCode) And (tlCff(ilUpper).iClfLine = tlClf.iLine)
        If (tlCff(ilUpper).iCntRevNo = tlClf.iCntRevNo) And (tlCff(ilUpper).iPropVer = tlClf.iPropVer) Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlCff(0 To ilUpper) As CFF
        End If
        ilRet = btrGetNext(hlCff, tlCff(ilUpper), ilCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    gGetFlightsFromLine = True
    Exit Function

End Function




Public Function gGetEarliestTrafSpotDate(hlSdf As Integer, ilVefCode As Integer) As Long
    Dim ilRet As Integer
    Dim llDate As Long
    Dim tlSdf As SDF
    Dim ilSdfRecLen As Integer
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim tlSdfSrchKey4 As SDFKEY4

    ilSdfRecLen = Len(tlSdf)
    If ilVefCode = -1 Then
        tlSdfSrchKey4.iDate(0) = 257
        tlSdfSrchKey4.iDate(1) = 1900
        tlSdfSrchKey4.lChfCode = 0
        ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) Then
            gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
            gGetEarliestTrafSpotDate = llDate
        Else
            gGetEarliestTrafSpotDate = -1
        End If
    Else
        tlSdfSrchKey1.iVefCode = ilVefCode
        tlSdfSrchKey1.iDate(0) = 257
        tlSdfSrchKey1.iDate(1) = 1900
        tlSdfSrchKey1.iTime(0) = 0
        tlSdfSrchKey1.iTime(1) = 0
        tlSdfSrchKey1.sSchStatus = ""
        ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tlSdf.iVefCode = ilVefCode) Then
            gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
            gGetEarliestTrafSpotDate = llDate
        Else
            gGetEarliestTrafSpotDate = -1
        End If
    End If
End Function

Public Function gGetEarliestAffSpotDate(hlAst As Integer, hlAtt As Integer, ilVefCode As Integer) As Long 'VBC NR
    Dim ilRet As Integer 'VBC NR
    Dim slOnAirDate As String 'VBC NR
    Dim llDate As Long 'VBC NR
    Dim llFeedDate As Long 'VBC NR
    Dim ilExtLen As Integer 'VBC NR
    Dim llNoRec As Long 'VBC NR
    Dim llRecPos As Long 'VBC NR
    Dim ilOffSet As Integer 'VBC NR
    Dim llMaxAtt As Long 'VBC NR
    Dim llAtt As Long 'VBC NR
    Dim tlAtt As ATT 'VBC NR
    Dim ilAttRecLen As Integer 'VBC NR
    Dim tlATTSrchKey1 As ATTKEY1 'VBC NR
    'Dim tlATTDateInfo As ATTDATEINFO
    Dim tlAst As AST 'VBC NR
    Dim ilAstRecLen As Integer 'VBC NR
    Dim tlAstSrchKey1 As ASTKEY1 'VBC NR
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record 'VBC NR

    ilAttRecLen = Len(tlAtt) 'VBC NR
    ilAstRecLen = Len(tlAst) 'VBC NR
    llDate = 0 'VBC NR
    If ilVefCode = -1 Then 'VBC NR
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slOnAirDate 'VBC NR
        If slOnAirDate = "" Then 'VBC NR
            gGetEarliestAffSpotDate = -1 'VBC NR
            Exit Function 'VBC NR
        End If 'VBC NR
        slOnAirDate = gObtainEndStd(DateAdd("d", -31 * tgSpf.iRetainAffSpot - 7, slOnAirDate)) 'VBC NR
        'btrExtClear hlAtt   'Clear any previous extend operation
        'ilExtLen = Len(tlATTDateInfo)  'Extract operation record size
        'ilRet = btrGetFirst(hlAtt, tlAtt, ilAttRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'If ilRet <> BTRV_ERR_NONE Then
        '    gGetEarliestAffSpotDate = -1
        '    Exit Function
        'End If
        'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        'Call btrExtSetBounds(hlAtt, llNoRec, -1, "UC", "ATTDATEEXTPK", ATTDATEEXTPK) 'Set extract limits (all records)
        'gPackDate slOnAirDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Att", "attOnAir")
        'ilRet = btrExtAddLogicConst(hlAtt, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        'ilOffset = gFieldOffset("Att", "attCode")
        'ilRet = btrExtAddField(hlAtt, ilOffset, 4) 'Extract start/end time, and days
        'If ilRet <> BTRV_ERR_NONE Then
        '    gGetEarliestAffSpotDate = -1
        '    Exit Function
        'End If
        'ilOffset = gFieldOffset("Att", "attOnAir")
        'ilRet = btrExtAddField(hlAtt, ilOffset, 4) 'Extract start/end time, and days
        'If ilRet <> BTRV_ERR_NONE Then
         '   gGetEarliestAffSpotDate = -1
        '    Exit Function
        'End If
        'ilOffset = gFieldOffset("Att", "attOffAir")
        'ilRet = btrExtAddField(hlAtt, ilOffset, 4) 'Extract start/end time, and days
        'If ilRet <> BTRV_ERR_NONE Then
        '    gGetEarliestAffSpotDate = -1
        '    Exit Function
        'End If
        'ilOffset = gFieldOffset("Att", "attDropDate")
        'ilRet = btrExtAddField(hlAtt, ilOffset, 4) 'Extract start/end time, and days
        'If ilRet <> BTRV_ERR_NONE Then
        '    gGetEarliestAffSpotDate = -1
        '    Exit Function
        'End If
        'ilRet = btrExtGetNext(hlAtt, tlATTDateInfo, ilExtLen, llRecPos)
        'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        '    'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
        '    ilExtLen = Len(tlATTDateInfo)  'Extract operation record size
        '    If ilRet = BTRV_ERR_REJECT_COUNT Then
        '        ilRet = btrExtGetNext(hlAtt, tlATTDateInfo, ilExtLen, llRecPos)
        '    End If
        '    Do While ilRet = BTRV_ERR_NONE
        '        tlAstSrchKey1.lAtfCode = tlATTDateInfo.lCode
        '        tlAstSrchKey1.iFeedDate(0) = 257
        '        tlAstSrchKey1.iFeedDate(1) = 1990
        '        ilRet = btrGetGreaterOrEqual(hlAst, tlAst, ilAstRecLen, tlAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        '        If (ilRet = BTRV_ERR_NONE) And (tlAst.lAtfCode = tlATTDateInfo.lCode) Then
        '            gUnpackDateLong tlAst.iFeedDate(0), tlAst.iFeedDate(1), llFeedDate
        '            If llDate = 0 Then
        '                llDate = llFeedDate
        '            Else
        '                If llFeedDate < llDate Then
        '                    llDate = llFeedDate
        '                End If
        '            End If
        '        End If
        '        ilRet = btrExtGetNext(hlAtt, tlATTDateInfo, ilExtLen, llRecPos)
        '        Do While ilRet = BTRV_ERR_REJECT_COUNT
        '            ilRet = btrExtGetNext(hlAtt, tlATTDateInfo, ilExtLen, llRecPos)
        '        Loop
        '    Loop
        'End If
        ilRet = btrGetLast(hlAtt, tlAtt, ilAttRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
        If ilRet <> BTRV_ERR_NONE Then 'VBC NR
            gGetEarliestAffSpotDate = -1 'VBC NR
            Exit Function 'VBC NR
        End If 'VBC NR
        llMaxAtt = tlAtt.lCode 'VBC NR
        For llAtt = 1 To llMaxAtt Step 1 'VBC NR
            tlAstSrchKey1.lAtfCode = llAtt 'VBC NR
            tlAstSrchKey1.iFeedDate(0) = 257 'VBC NR
            tlAstSrchKey1.iFeedDate(1) = 1990 'VBC NR
            ilRet = btrGetGreaterOrEqual(hlAst, tlAst, ilAstRecLen, tlAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date 'VBC NR
            If (ilRet = BTRV_ERR_NONE) And (tlAst.lAtfCode = llAtt) Then 'VBC NR
                gUnpackDateLong tlAst.iFeedDate(0), tlAst.iFeedDate(1), llFeedDate 'VBC NR
                If llDate = 0 Then 'VBC NR
                    llDate = llFeedDate 'VBC NR
                Else 'VBC NR
                    If llFeedDate < llDate Then 'VBC NR
                        llDate = llFeedDate 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
            End If 'VBC NR
        Next llAtt 'VBC NR
        llAtt = llMaxAtt + 1 'VBC NR
        Do 'VBC NR
            tlAstSrchKey1.lAtfCode = llAtt 'VBC NR
            tlAstSrchKey1.iFeedDate(0) = 257 'VBC NR
            tlAstSrchKey1.iFeedDate(1) = 1990 'VBC NR
            ilRet = btrGetGreaterOrEqual(hlAst, tlAst, ilAstRecLen, tlAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date 'VBC NR
            If (ilRet = BTRV_ERR_NONE) Then 'VBC NR
                gUnpackDateLong tlAst.iFeedDate(0), tlAst.iFeedDate(1), llFeedDate 'VBC NR
                If llDate = 0 Then 'VBC NR
                    llDate = llFeedDate 'VBC NR
                Else 'VBC NR
                    If llFeedDate < llDate Then 'VBC NR
                        llDate = llFeedDate 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
            Else 'VBC NR
                Exit Do 'VBC NR
            End If 'VBC NR
            llAtt = tlAst.lAtfCode + 1 'VBC NR
        Loop 'VBC NR
    Else 'VBC NR
        tlATTSrchKey1.iVefCode = ilVefCode 'VBC NR
        ilRet = btrGetEqual(hlAtt, tlAtt, ilAttRecLen, tlATTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
        Do While (ilRet = BTRV_ERR_NONE) And (tlAtt.iVefCode = ilVefCode) 'VBC NR
            tlAstSrchKey1.lAtfCode = tlAtt.lCode 'VBC NR
            tlAstSrchKey1.iFeedDate(0) = 257 'VBC NR
            tlAstSrchKey1.iFeedDate(1) = 1990 'VBC NR
            ilRet = btrGetGreaterOrEqual(hlAst, tlAst, ilAstRecLen, tlAstSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date 'VBC NR
            If (ilRet = BTRV_ERR_NONE) And (tlAst.iVefCode = ilVefCode) Then 'VBC NR
                gUnpackDateLong tlAst.iFeedDate(0), tlAst.iFeedDate(1), llFeedDate 'VBC NR
                If llDate = 0 Then 'VBC NR
                    llDate = llFeedDate 'VBC NR
                Else 'VBC NR
                    If llFeedDate < llDate Then 'VBC NR
                        llDate = llFeedDate 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
            End If 'VBC NR
            ilRet = btrGetNext(hlAtt, tlAtt, ilAttRecLen, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
        Loop 'VBC NR
    End If 'VBC NR
    gGetEarliestAffSpotDate = llDate 'VBC NR
End Function 'VBC NR
Private Sub gInitGuide(ByRef tlUrf As URF)
'set values that may be changed by user later
    tlUrf.sName = "Guide"
    tlUrf.iVefCode = 0
    tlUrf.sPassword = "Guide"
    tlUrf.iCode = 0  'Autoincrement
    tlUrf.iRemoteID = 0
    tlUrf.iAutoCode = tlUrf.iCode
    tlUrf.sRept = "System"
    tlUrf.sPhoneNo = ""
    tlUrf.sCity = ""
    tlUrf.lEMailCefCode = 0
    tlUrf.iSlfCode = 0
    tlUrf.sClnMoYr = "M"    'Calendar by month
    tlUrf.sClnType = "S"    'Calendar type-standard
    tlUrf.sClnLayout = "A"  'Calendar- Across
    tlUrf.iClnLeft = 100    'Calendar left
    tlUrf.iClnTop = 200 'Calendar top
    tlUrf.iClcLeft = 2775   'Traffic.Width - 2775 - 100   '2625 width of calc form
    tlUrf.iClcTop = 200

    tlUrf.iDefVeh = 0
    tlUrf.iGroupNo = 0
    tlUrf.iRemoteUserID = 0
    tlUrf.sPDFDrvChar = ""
    tlUrf.iPDFDnArrowCnt = 0
    tlUrf.sPrtDrvChar = ""
    tlUrf.iPrtDnArrowCnt = 0
    tlUrf.sPrtNameAltKey = ""
    tlUrf.iPrtNoEnterKeys = 0
    tlUrf.sOldPassword(0) = ""
    tlUrf.sOldPassword(1) = ""
    tlUrf.sOldPassword(2) = ""

    tlUrf.sRCView = "R"

    gLimitGuide tlUrf
End Sub
Private Sub gLimitGuide(ByRef tlUrf As URF)
Dim ilLoop As Integer
    For ilLoop = 0 To 69 Step 1
        tlUrf.sWin(ilLoop) = "H"
    Next ilLoop
    tlUrf.sWin(USERLIST) = "I"
    tlUrf.sWin(SITELIST) = "I"
    tlUrf.sGrid = "H"
    tlUrf.sPrice = "H"
    tlUrf.sCredit = "H"
    tlUrf.sPayRate = "H"
    tlUrf.sMerge = "H"
    tlUrf.sDelete = "N"
    tlUrf.sHideSpots = "H"
    tlUrf.sChgBilled = "H"
    tlUrf.sChgCntr = "H"
    tlUrf.sChgCrRt = "H"
    tlUrf.sBouChk = "H"
    tlUrf.sReprintLogAlert = "N"
    tlUrf.sIncompAlert = "N"
    tlUrf.sCompAlert = "N"
    tlUrf.sSchAlert = "N"
    tlUrf.sHoldAlert = "N"
    tlUrf.sPrgmmaticAlert = "H"
    tlUrf.sAdvanceAvails = "H"
    tlUrf.sShowNRMsg = "N"
    tlUrf.sRateCardAlert = "N"
    tlUrf.sResearchAlert = "N"
    tlUrf.sAvailAlert = "N"
    tlUrf.sCrdChkAlert = "N"
    tlUrf.sDeniedAlert = "N"
    tlUrf.sCrdLimitAlert = "N"
    tlUrf.sMoveAlert = "N"
    tlUrf.sWorkToDead = "N"
    tlUrf.sWorkToComp = "N"
    tlUrf.sWorkToHold = "N"
    tlUrf.sWorkToOrder = "N"
    tlUrf.sCompToIncomp = "N"
    tlUrf.sCompToDead = "N"
    tlUrf.sCompToHold = "N"
    tlUrf.sCompToOrder = "N"
    tlUrf.sIncompToDead = "N"
    tlUrf.sIncompToComp = "N"
    tlUrf.sIncompToHold = "N"
    tlUrf.sIncompToOrder = "N"
    tlUrf.sDeadToWork = "N"
    tlUrf.sHoldToOrder = "N"
    tlUrf.sReviseCntr = "N"
    tlUrf.sResvType = "H"
    tlUrf.sRemType = "H"
    tlUrf.sDRType = "H"
    tlUrf.sPIType = "H"
    tlUrf.sPSAType = "H"
    tlUrf.sPromoType = "H"
    tlUrf.sRefResvType = "H"
    tlUrf.sUseComputeCMC = "H"
    tlUrf.sBlockRU = "N"
    tlUrf.sRegionCopy = "H"
    tlUrf.sChgPrices = "H"
    tlUrf.sChgLnBillPrice = "H"
    tlUrf.sActFlightButton = "H"
    tlUrf.sLiveLogPostOnly = "N"
    tlUrf.sSportPropOnly = "N"
    tlUrf.iSnfCode = 0
        ' Dan M new selected fields 4/10/09
    tlUrf.sChangeCSIDate = "H"
    tlUrf.sAllowInvDisplay = "H"
    tlUrf.sActivityLog = "H"

End Sub
Private Function gTestLimitGuide(slUserName As String) As Boolean
'are tables wrong after change to limited guide?
Dim ilLoop As Integer
gTestLimitGuide = False
    'If StrComp(Trim(tgPopUrf(2).sName), slUserName, vbTextCompare) = 0 Then 'or compare slusername to sgsuname?
    If StrComp(Trim(tgPopUrf(1).sName), slUserName, vbTextCompare) = 0 Then 'or compare slusername to sgsuname?
' are report buttons available when they shouldn't be?
        For ilLoop = 0 To 10 Step 1
            'If tgPopUrf(2).sWin(ilLoop) = "I" Then
            If tgPopUrf(1).sWin(ilLoop) = "I" Then
                gTestLimitGuide = True
            End If
        Next ilLoop
    End If
End Function
Private Function gTestInternalGuide() As Boolean
gTestInternalGuide = False
If sgSpecialPassword <> "" Then
    gTestInternalGuide = True
End If
End Function
Public Sub gExpandGuideAsUser(ByRef tlUrf As URF)
Dim ilLoop As Integer
    bgInternalGuide = True
    For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
            igWinStatus(ilLoop) = 2
    Next
    gInitSuperUser tlUrf
End Sub
Private Sub gRecordLimitGuide(tlUrf As URF)
Dim hlUrf As Integer
Dim ilRet As Integer
Dim ilRecLen As Integer
Dim tlGuideUrf As URF
Dim ilGuideFinder As Integer

    hlUrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilGuideFinder = tlUrf.iAutoCode
    ilRecLen = Len(tlGuideUrf)
    ilRet = btrGetEqual(hlUrf, tlGuideUrf, ilRecLen, ilGuideFinder, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    tlGuideUrf = tlUrf
    gUrfEncrypt tlGuideUrf
    ilRet = btrUpdate(hlUrf, tlGuideUrf, ilRecLen)
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
End Sub

Public Sub gSetWinStatusFromSPF(ilWinStatus() As Integer)
    If (tgUrf(0).iCode <= 2) Then
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic = "Y") Then
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys <> "Y") And (tgSpf.sUsingTraffic = "Y") Then
        ilWinStatus(BUDGETSJOB) = 0
        ilWinStatus(PROPOSALSJOB) = 0
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys <> "Y") Then
        ilWinStatus(CONTRACTSJOB) = 0
        ilWinStatus(PROGRAMMINGJOB) = 0
        ilWinStatus(SPOTSJOB) = 0
        ilWinStatus(COPYJOB) = 0
        ilWinStatus(LOGSJOB) = 0
'        ilWinStatus(STATIONFEEDJOB) = 0
        ilWinStatus(POSTLOGSJOB) = 0
        ilWinStatus(INVOICESJOB) = 0
        ilWinStatus(COLLECTIONSJOB) = 0
        ilWinStatus(SLSPCOMMSJOB) = 0
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys = "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys = "Y") Then
        ilWinStatus(CONTRACTSJOB) = 0
        ilWinStatus(SPOTSJOB) = 0
        ilWinStatus(COPYJOB) = 0
        ilWinStatus(LOGSJOB) = 0
'        ilWinStatus(STATIONFEEDJOB) = 0
        ilWinStatus(POSTLOGSJOB) = 0
        ilWinStatus(INVOICESJOB) = 0
        ilWinStatus(COLLECTIONSJOB) = 0
        ilWinStatus(SLSPCOMMSJOB) = 0
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys <> "Y") And (tgSpf.sUsingTraffic <> "Y") And (tgSpf.sGUseAffSys = "Y") Then
        ilWinStatus(BUDGETSJOB) = 0
        ilWinStatus(RATECARDSJOB) = 0
        ilWinStatus(PROPOSALSJOB) = 0
        ilWinStatus(CONTRACTSJOB) = 0
        ilWinStatus(SPOTSJOB) = 0
        ilWinStatus(COPYJOB) = 0
        ilWinStatus(LOGSJOB) = 0
'        ilWinStatus(STATIONFEEDJOB) = 0
        ilWinStatus(POSTLOGSJOB) = 0
        ilWinStatus(INVOICESJOB) = 0
        ilWinStatus(COLLECTIONSJOB) = 0
        ilWinStatus(SLSPCOMMSJOB) = 0
        Exit Sub
    End If
    
End Sub

Public Sub gGetEventTitles(ilVefCode As Integer, slEventTitle1 As String, slEventTitle2 As String)
    Dim hlVtf As Integer
    Dim ilVtfRecLen As Integer        'Record length
    Dim ilRet As Integer 'btrieve status
    Dim tlVtf As VTF   'file record
    Dim tlVtfSrchKey1 As VTFKEY1

    On Error Resume Next
    slEventTitle1 = ""
    slEventTitle2 = ""
    ilVtfRecLen = Len(tlVtf)
    hlVtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlVtf, "", sgDBPath & "Vtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        tlVtfSrchKey1.iVefCode = ilVefCode
        tlVtfSrchKey1.sType = "1"
        ilRet = btrGetEqual(hlVtf, tlVtf, ilVtfRecLen, tlVtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slEventTitle1 = gStripChr0(tlVtf.sText)
        End If
        tlVtfSrchKey1.iVefCode = ilVefCode
        tlVtfSrchKey1.sType = "2"
        ilRet = btrGetEqual(hlVtf, tlVtf, ilVtfRecLen, tlVtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slEventTitle2 = gStripChr0(tlVtf.sText)
        End If
    End If
    If slEventTitle1 = "" Then
        If Trim$(tgSaf(0).sEventTitle1) <> "" Then
            slEventTitle1 = Trim$(tgSaf(0).sEventTitle1)
        Else
            slEventTitle1 = "Visiting Team"
        End If
    End If
    If slEventTitle2 = "" Then
        If Trim$(tgSaf(0).sEventTitle2) <> "" Then
            slEventTitle2 = Trim$(tgSaf(0).sEventTitle2)
        Else
            slEventTitle2 = "Home Team"
        End If
    End If
    ilRet = btrClose(hlVtf)
    btrDestroy hlVtf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetFlightWhichPrice            *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Duplicate of gGetFlightPrice    *
'           and changed to add option field of which
'           rate to return: either Actual or Proposal
'*                                                     *
'*******************************************************
Function gGetFlightWhichPrice(tlSdf As SDF, tlClf As CLF, hlCff As Integer, hlSmf As Integer, slPrice As String, Optional slWhichPrice As String = "A") As Integer
'
'   ilRet = gGetFlightWhichPrice(tlSdf, hlCff, hlSmf, slPrice)
'   Where:
'       tlSdf(I)- Spot image to get flight for
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       hlSmf(I)- Open handle to SMF
'       slPrice(O)- Spot price
'       slWHichPrice (I) : "A" = Actual, "P" proposal price
'       tgPriceCff(O)- Flight associated with sdf
'       ilRet = True if found; False if error
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slBonusOnInv As String
    'Dim tlCff As CFF

    If tlSdf.sSpotType = "X" Then
        slBonusOnInv = "Y"
        'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
            'If tlSdf.iAdfCode = tgCommAdf(ilLoop).iCode Then
            ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
            If ilLoop <> -1 Then
                slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
            '    Exit For
            End If
        'Next ilLoop
        'If tlSdf.sPriceType <> "N" Then
        If tlSdf.sPriceType = "+" Then
            slPrice = "+ Fill"  '"> Fill"
        ElseIf tlSdf.sPriceType = "-" Then
            slPrice = "- Fill"  '"< Fill"
        Else
            If slBonusOnInv <> "N" Then
                slPrice = "+ Fill"
            Else
                slPrice = "- Fill"
            End If
        End If
        gGetFlightWhichPrice = True
        Exit Function
    ElseIf tlSdf.sSpotType = "O" Or tlSdf.sSpotType = "C" Then      '12-11-12
        slPrice = ".00"
        gGetFlightWhichPrice = True
        Exit Function
    End If
    ilRet = gGetSpotFlight(tlSdf, tlClf, hlCff, hlSmf, tgPriceCff)  'tlCff)
    If ilRet = True Then
        Select Case tgPriceCff.sPriceType   'tlCff.sPriceType
       '1-9-19 Each case needs to test if using Prop Price to pick up the prop price to accumulate vs the actual spot price.
       'Previously , line types defined as N/C, Bonus etc were not gettin gthe Prop Price if that was the requested option
            Case "T"    'True
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                Else
                    slPrice = gLongToStrDec(tgPriceCff.lActPrice, 2)    'tlCff.lActPrice, 2)
                End If
            Case "N"    'No Charge
                slPrice = "N/C"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "M"    'MG Line
                slPrice = "MG"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "B"    'Bonus
                slPrice = "Bonus"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "S"    'Spinoff
                slPrice = "Spinoff"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "P"    'Package
                slPrice = ".00"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "R"    'Recapturable
                slPrice = "Recapturable"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
            Case "A"    'ADU
                slPrice = "ADU"
                If slWhichPrice = "P" Then          'use proposal price
                    slPrice = gLongToStrDec(tgPriceCff.lPropPrice * 100, 2)
                End If
        End Select
        gGetFlightWhichPrice = True
    Else
        slPrice = ".00"
        gGetFlightWhichPrice = False
    End If
    Exit Function
End Function

Public Function gPoolExist() As Boolean
    Dim slSQLQuery As String
    Dim rst_adf As ADODB.Recordset
    gPoolExist = False
    slSQLQuery = "Select Count(1) as BkoutCount from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where "
    slSQLQuery = slSQLQuery & " (adfBkoutPoolStatus = 'A'"
    slSQLQuery = slSQLQuery & " Or adfBkoutPoolStatus = 'U'" & ")"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    If Not rst_adf.EOF Then
        If rst_adf!BkoutCount > 0 Then
            gPoolExist = True
        End If
    End If
    rst_adf.Close
End Function
