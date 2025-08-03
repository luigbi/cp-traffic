Attribute VB_Name = "BLACKOUTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Blackout.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Budget.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
'Dim tmSRec As LPOPREC

Type SPOTSUM
    'ExpNY and Logs
    iVefCode As Integer
    lDate As Long
    lChfCode As Long
    iMnfComp(0 To 1) As Integer 'Product Protection
    iAdfCode As Integer
    lTime As Long
    iLen As Integer
    sProduct As String * 35
    'ExpNY
    iNewIndex As Integer    'Index into smNewLines
    sShortTitle As String * 15
    'Logs (part of Odf key)
    sZone As String * 3     'Time zone
    iSeqNo As Integer   'Seq number to kept same time events in order
    'Logs (Lst key)
    lLstCode As Long
    'Copy
    iLnVefCode As Integer
    imnfSeg As Integer      '6-19-01 mnf segment code from chf
    'Blackout replacement
    lSdfCode As Long
    sLogType As String * 1  'Log: F=Final; R=Reprint
    lCrfCode As Long
    sDays As String * 7
    iOrigAirDate(0 To 1) As Integer
End Type
Public tgSpotSum() As SPOTSUM
Public ig30StartBofIndex As Integer
Public ig60StartBofIndex As Integer
Public igStartBofIndex As Integer
Public lgStartIndex As Long
Public lgEndIndex As Long
'Blackout
Dim tmBofSrchKey0 As LONGKEY0
Dim tmBofSrchKey1 As BOFKEY1    'Bof key record image
Dim imBofRecLen As Integer        'Bof record length
Dim tmBof As BOF
'Copy Inventory
Dim tmCifSrchKey As LONGKEY0    'Bof key record image
Dim imCifRecLen As Integer        'Bof record length
Dim tmCif As CIF
'Product/ISCI record
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0
Dim imCpfRecLen As Integer

'Copy Vehicles
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0  'CVF key record image
Dim tmCvfSrchKey1 As LONGKEY0  'CVF key record image
Dim imCvfRecLen As Integer      'CVF record length

Public tgAsgnVehCvf As CVF
Public igAsgnVehRowNo As Integer

'Media code record information
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
'Product
Dim tmPrfSrchKey As LONGKEY0    'Bof key record image
Dim imPrfRecLen As Integer        'Bof record length
Dim tmPrf As PRF
'Short Title
Dim tmSifSrchKey As LONGKEY0    'Bof key record image
Dim imSifRecLen As Integer        'Bof record length
Dim tmSif As SIF
'Contract
Dim tmChfSrchKey As LONGKEY0    'Bof key record image
Dim imCHFRecLen As Integer        'Bof record length
Dim tmChf As CHF
'Log Spot record
Dim tmLst As LST
Dim imLstRecLen As Integer
Dim tmLstSrchKey As LONGKEY0
'One day file (ODF)
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdfSrchKey0 As ODFKEY0 'ODF key record image
Dim tmOdf As ODF            'ODF record image
'Regional or Blackout Copy
Dim tmRsf As RSF
Dim tmRsfSrchKey As LONGKEY0
Dim tmRsfSrchKey1 As LONGKEY0
Dim tmRsfSrchKey4 As RSFKEY4

Dim imRsfRecLen As Integer
Type BOFREC
    sKey As String * 130 'End Date, Advertiser
    tBof As BOF
    sAdfName As String * 46
    sVefName As String * 40
    sShtTitle As String * 35    'Note: Short Title is 15 characters, Product is 35
    iLen As Integer
    lSCntrNo As Long
    sRAdfName As String * 46
    lRCntrNo As Long
    iStatus As Integer
    lRecPos As Long
End Type
'Type USERVEH
'    iCode As Integer
'    sName As String * 40
'End Type
Public igView As Integer
Public tgSBofRec() As BOFREC  'Suppression
Public tgRBofRec() As BOFREC  'Replacement
Public tgBofDel() As BOFREC
'Public igShowHelpMsg As Integer
Dim tmAvail As AVAILSS
Dim tmSdf As SDF
'4/30/11:  Add Region copy by Airing vehicle
Dim tmCrf As CRF    'Used to pass back crf from gObtainAirCopy to calling routine

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetAirCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Public Sub gGetAirCopy(slVefType As String, ilVefCode As Integer, ilVpfIndex As Integer, tlSdf As SDF, hlCrf As Integer, hlRsf As Integer, hlCvf As Integer, slZone As String)
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slType As String
    Dim ilARotNo As Integer
    Dim slAZone As String
    'Copy rotation record information
    'Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim tlCrfSrchKey4 As CRFKEY4 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim blVefFound As Boolean
    Dim tlCrf As CRF            'CRF record image


    If (slVefType <> "A") Or (tgVpf(ilVpfIndex).sCopyOnAir <> "Y") Then
        Exit Sub
    End If
    ilFound = False
    ilARotNo = -1
    imRsfRecLen = Len(tmRsf)
    ilCrfRecLen = Len(tlCrf)
    '7/15/14
    tmRsfSrchKey4.lSdfCode = tlSdf.lCode
    tmRsfSrchKey4.sType = "A"
    tmRsfSrchKey4.iBVefCode = ilVefCode
    ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tlSdf.lCode) And (tmRsf.sType = "A")
        If (tmRsf.iBVefCode = ilVefCode) Then
            slType = "A"
            'tlCrfSrchKey1.sRotType = slType
            'tlCrfSrchKey1.iEtfCode = 0
            'tlCrfSrchKey1.iEnfCode = 0
            'tlCrfSrchKey1.iadfCode = tlSdf.iadfCode
            'tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
            'tlCrfSrchKey1.iVefCode = ilVefCode  'tmVef.iCode
            'tlCrfSrchKey1.iRotNo = tmRsf.iRotNo
            'ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            tlCrfSrchKey4.sRotType = slType
            tlCrfSrchKey4.iEtfCode = 0
            tlCrfSrchKey4.iEnfCode = 0
            tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
            tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
            tlCrfSrchKey4.lFsfCode = 0
            tlCrfSrchKey4.iRotNo = tmRsf.iRotNo
            ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            blVefFound = False
            Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iRotNo = tmRsf.iRotNo)
                blVefFound = gCheckCrfVehicle(ilVefCode, tlCrf, hlCvf)
                If blVefFound Then
                    Exit Do
                End If
                ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            
            'If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iadfCode = tlSdf.iadfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode) And (tlCrf.iRotNo = tmRsf.iRotNo) Then      'tmVef.iCode)    'ilVefCode)
            If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (blVefFound) And (tlCrf.iRotNo = tmRsf.iRotNo) Then      'tmVef.iCode)    'ilVefCode)
                If (Trim$(tlCrf.sZone) = "") Or (StrComp(Trim$(tlCrf.sZone), Trim$(slZone), vbTextCompare) = 0) Then
                    If (ilARotNo = -1) Or (tlCrf.iRotNo > ilARotNo) Then
                        ilARotNo = tlCrf.iRotNo
                        tlSdf.sPtType = "1"
                        tlSdf.lCopyCode = tmRsf.lCopyCode
                        slAZone = tlCrf.sZone
                    End If
                ElseIf (StrComp(Trim$(tlCrf.sZone), "Oth", vbTextCompare) = 0) Then
                    If (ilARotNo = -1) Or ((tlCrf.iRotNo > ilARotNo) And ((Trim$(slAZone) = "") Or (StrComp(Trim$(slAZone), "Oth", vbTextCompare) = 0))) Then
                        ilARotNo = tlCrf.iRotNo
                        tlSdf.sPtType = "1"
                        tlSdf.lCopyCode = tmRsf.lCopyCode
                        slAZone = tlCrf.sZone
                    End If
                End If
            End If
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hlRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainAirCopy                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Sub gObtainAirCopy(ilType As Integer, slVefType As String, ilVefCode As Integer, ilVpfIndex As Integer, tlSdf As SDF, tlAvail As AVAILSS, hlCrf As Integer, hlCnf As Integer, hlCif As Integer, hlCvf As Integer, hlClf As Integer, slZone As String, ilZoneFd As Integer, ilCopyReplaced As Integer, ilRotNo As Integer)
' 10451 added hlClf
'
'   ilType(I)- 0=Airing only (Airing spots copy); 1=Specified vehicle (call by blackout replacement without copy)
'              2=Airing Only (Airing spot copy but zone must match); 3=Specified vehicle (like 1 except test only, return ilRotNo)
'              4=Airing Only (like 0 except test only, return ilRotNo);
'              6=Airing Only (Airing spot copy but zone must match)(like 2 except test only, return ilRotNo)
'
'
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llSpotTime As Long
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim ilAsgnDate0 As Integer
    Dim ilAsgnDate1 As Integer
    Dim slType As String
    'Dim llCrfRecPos As Long
    Dim llCrfCode As Long
    'Dim llCifRecPos As Long
    Dim llCifCode As Long
    Dim ilDay As Integer
    Dim ilSpotAsgn As Integer
    Dim ilAvailOk As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    'Copy rotation record information
    'Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim tlCrfSrchKey As LONGKEY0 'CIF key record image
    Dim tlCrfSrchKey4 As CRFKEY4 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    'Copy instruction record information
    Dim tlCnfSrchKey As CNFKEY0 'CNF key record image
    Dim ilCnfRecLen As Integer  'CNF record length
    Dim tlCnf As CNF            'CNF record image
    'Copy inventory record information
    Dim tlCifSrchKey As LONGKEY0 'CIF key record image
    Dim ilCifRecLen As Integer  'CIF record length
    Dim tlCif As CIF            'CIF record image
    Dim ilBypassCrf As Integer
    
    Dim blVefFound As Boolean
    Dim blCrfFound As Boolean
    Dim ilCvf As Integer
    Dim ilCrf As Integer
    Dim ilSvRotNo As Integer
    '10451
    Dim slLive As String
    
    ReDim llProcessedCrfCode(0 To 0) As Long
    
    ilZoneFd = False
    ilCopyReplaced = False
    ilRotNo = 0
    If (ilType = 0) Or (ilType = 2) Or (ilType = 4) Or (ilType = 6) Then
        If (slVefType <> "A") Or (tgVpf(ilVpfIndex).sCopyOnAir <> "Y") Then
            Exit Sub
        End If
    End If
    'Test if copy has been superseded and if so, replace copy definition within tlSdf
    'Find rotation to assign
    'Code later- test spot type to determine which rotation type
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    ilCrfRecLen = Len(tlCrf)
    ilCnfRecLen = Len(tlCnf)
    ilCifRecLen = Len(tlCif)
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
    ilDay = gWeekDayLong(llDate)
    gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llSpotTime
    ilSpotAsgn = False
    ilAvailOk = True
    slType = "A"
    
    '10451
    slLive = gGetClfLive(hlClf, tlSdf)
    If slLive = "L" Or slLive = "M" Then
        Exit Sub
    End If
    'tlCrfSrchKey1.sRotType = slType
    'tlCrfSrchKey1.iEtfCode = 0
    'tlCrfSrchKey1.iEnfCode = 0
    'tlCrfSrchKey1.iadfCode = tlSdf.iadfCode
    'tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
    'tlCrfSrchKey1.lFsfCode = 0
    'tlCrfSrchKey1.iVefCode = ilVefCode  'tmVef.iCode
    'tlCrfSrchKey1.iRotNo = 32000
    'ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
    'Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iadfCode = tlSdf.iadfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode)   'tmVef.iCode)    'ilVefCode)
    tlCrfSrchKey4.sRotType = slType
    tlCrfSrchKey4.iEtfCode = 0
    tlCrfSrchKey4.iEnfCode = 0
    tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
    tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
    tlCrfSrchKey4.lFsfCode = 0
    tlCrfSrchKey4.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) 'And (tlCrf.iVefCode = ilVefCode)   'tmVef.iCode)    'ilVefCode)
        'ilRet = btrGetPosition(hlCrf, llCrfRecPos)
        'Test date, time, day and zone
        llCrfCode = tlCrf.lCode
        ilSvRotNo = tlCrf.iRotNo
        llProcessedCrfCode(UBound(llProcessedCrfCode)) = tlCrf.lCode
        ReDim Preserve llProcessedCrfCode(0 To UBound(llProcessedCrfCode) + 1) As Long
        ilSpotAsgn = False
        ilBypassCrf = False
        If tlCrf.sState <> "D" Then
            'Note:  Only Recorded copy allowed with airing vehicles
            'Test if looking for Live or Recorded rotations
            'If slLive = "L" Then
            '    If tmCrf.sLiveCopy <> "L" Then
            '        ilBypassCrf = True
            '    End If
            'Else
                If (tlCrf.sLiveCopy = "L") Or (tlCrf.sLiveCopy = "M") Then
                    ilBypassCrf = True
                End If
            'End If
             '10451 added
             'Test Recorded rotations
            If ilBypassCrf = False Then
                If slLive = "S" Then
                    If tlCrf.sLiveCopy <> "S" Then
                        ilBypassCrf = True
                    End If
                ElseIf slLive = "P" Then
                    If tlCrf.sLiveCopy <> "P" Then
                        ilBypassCrf = True
                    End If
                ElseIf slLive = "Q" Then
                    If tlCrf.sLiveCopy <> "Q" Then
                        ilBypassCrf = True
                    End If
                'R
                Else
                    If (tlCrf.sLiveCopy = "S") Or (tlCrf.sLiveCopy = "P") Or (tlCrf.sLiveCopy = "Q") Then
                        ilBypassCrf = True
                    End If
                End If
            End If
        Else
            ilBypassCrf = True
        End If
        If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo > tlCrf.iRotNo)) Then
            '1/4/15: Since CRF read with RotNo in desceasing order, once it is equal or low exit
            'ilBypassCrf = True
            Exit Sub
        End If
        If Not ilBypassCrf Then
            If Not gCheckCrfVehicle(ilVefCode, tlCrf, hlCvf) Then
                ilBypassCrf = True
            End If
        End If
        If (tlCrf.sDay(ilDay) = "Y") And (tlSdf.iLen = tlCrf.iLen) And (Not ilBypassCrf) And ((Trim$(tlCrf.sZone) = "") Or (Trim$(tlCrf.sZone) = Trim$(slZone))) Then
            If (ilType = 0) Or (ilType = 1) Or (ilType = 3) Or (ilType = 4) Or (((ilType = 2) Or (ilType = 6)) And (Trim$(tlCrf.sZone) = Trim$(slZone))) Then
                gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slDate
                llSAsgnDate = gDateValue(slDate)
                gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slDate
                llEAsgnDate = gDateValue(slDate)
                If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                    gUnpackTime tlCrf.iStartTime(0), tlCrf.iStartTime(1), "A", "1", slTime
                    llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                    gUnpackTime tlCrf.iEndTime(0), tlCrf.iEndTime(1), "A", "1", slTime
                    llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                    If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                        ilAvailOk = True    'Ok to book into
                        If (tlCrf.sInOut = "I") Or (tlCrf.sInOut = "O") Then
                            'This line was required for assign copy to missed spots
                            'it is left in even though assigning to missed has been removed
                            If ((tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O")) And (tlSdf.iRotNo <> -1) Then    'Add spot
                                If tlCrf.sInOut = "I" Then
                                    If tlCrf.ianfCode <> tlAvail.ianfCode Then
                                        ilAvailOk = False   'No
                                    End If
                                Else
                                    If tlCrf.ianfCode = tlAvail.ianfCode Then
                                        ilAvailOk = False   'No
                                    End If
                                End If
                            End If
                        End If
                        If ilAvailOk Then
                            If Not ilSpotAsgn Then
                                If (ilType = 3) Or (ilType = 4) Or (ilType = 6) Then
                                    ilRotNo = tlCrf.iRotNo
                                    '4/30/11:  Add Region copy by Airing vehicle
                                    tmCrf = tlCrf
                                    Exit Sub
                                End If
                                ilRet = btrBeginTrans(hlCrf, 1000)
                                'Assign copy, find next instruction to be assigned
                                '9/30/16: Handle sync copy
                                tlCnfSrchKey.lCrfCode = tlCrf.lCode
                                If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
                                    'If ilPFAsgn = 1 Then    'Final
                                        tlCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextFinal(igAsgnVehRowNo)
                                    'Else
                                    '    tmCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextPrelim(igAsgnVehRowNo)
                                    'End If
                                Else
                                    'If ilPFAsgn = 1 Then    'Final
                                        tlCnfSrchKey.iInstrNo = tlCrf.iNextFinal
                                    'Else
                                    '    tlCnfSrchKey.iInstrNo = tlCrf.iNextPrelim
                                    'End If
                                End If
                                ilRet = btrGetEqual(hlCnf, tlCnf, ilCnfRecLen, tlCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet <> BTRV_ERR_NONE Then
                                    tlCnfSrchKey.lCrfCode = tlCrf.lCode
                                    tlCnfSrchKey.iInstrNo = 1
                                    ilRet = btrGetEqual(hlCnf, tlCnf, ilCnfRecLen, tlCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                End If
                                If ilRet = BTRV_ERR_NONE Then
                                    'One inventory per instruction
                                    If tlCnf.lCifCode > 0 Then
                                        tlCifSrchKey.lCode = tlCnf.lCifCode
                                        ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        llCifCode = tlCif.lCode
                                    Else
                                        'Not coded
                                        ilRet = Not BTRV_ERR_NONE
                                    End If
                                    If ilRet = BTRV_ERR_NONE Then
                                        'ilRet = btrGetPosition(hlCrf, llCrfRecPos)
                                        Do
                                            'ilRet = btrGetDirect(hlCrf, tlCrf, ilCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                                            tlCrfSrchKey.lCode = llCrfCode
                                            ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrAbortTrans(hlCrf)
                                                'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetDirect Crf(1), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                Exit Sub
                                            End If
                                            'tmSRec = tlCrf
                                            'ilRet = gGetByKeyForUpdate("Crf", hlCrf, tmSRec)
                                            'tlCrf = tmSRec
                                            'If ilRet <> BTRV_ERR_NONE Then
                                            '    ilCRet = btrAbortTrans(hlCrf)
                                            '    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Crf(2), Try Later", vbOkOnly + vbExclamation, "Erase")
                                            '    Exit Sub
                                            'End If
                                            'If ilPFAsgn = 1 Then    'Final
                                                tlCrf.iNextFinal = tlCnf.iInstrNo + 1
                                                tlCrf.iNextPrelim = tlCrf.iNextFinal
                                            'Else
                                            '    tlCrf.iNextPrelim = tlCnf.iInstrNo + 1
                                            'End If
                                            If (tlCrf.iEarliestDateAssg(0) = 0) And (tlCrf.iEarliestDateAssg(1) = 0) Then
                                                tlCrf.iEarliestDateAssg(0) = ilAsgnDate0
                                                tlCrf.iEarliestDateAssg(1) = ilAsgnDate1
                                                tlCrf.iLatestDateAssg(0) = ilAsgnDate0
                                                tlCrf.iLatestDateAssg(1) = ilAsgnDate1
                                            Else
                                                If (ilAsgnDate1 < tlCrf.iEarliestDateAssg(1)) Or ((ilAsgnDate1 = tlCrf.iEarliestDateAssg(1)) And (ilAsgnDate0 < tlCrf.iEarliestDateAssg(0))) Then
                                                    tlCrf.iEarliestDateAssg(0) = ilAsgnDate0
                                                    tlCrf.iEarliestDateAssg(1) = ilAsgnDate1
                                                End If
                                                If (ilAsgnDate1 > tlCrf.iLatestDateAssg(1)) Or ((ilAsgnDate1 = tlCrf.iLatestDateAssg(1)) And (ilAsgnDate0 > tlCrf.iLatestDateAssg(0))) Then
                                                    tlCrf.iLatestDateAssg(0) = ilAsgnDate0
                                                    tlCrf.iLatestDateAssg(1) = ilAsgnDate1
                                                End If
                                            End If
                                            tlCrf.iLastDateAssg(0) = ilAsgnDate0
                                            tlCrf.iLastDateAssg(1) = ilAsgnDate1
                                            gPackDate slNowDate, tlCrf.iDateAssgDone(0), tlCrf.iDateAssgDone(1)
                                            gPackTime slNowTime, tlCrf.iTimeAssgDone(0), tlCrf.iTimeAssgDone(1)
                                            ilRet = btrUpdate(hlCrf, tlCrf, ilCrfRecLen)
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilCRet = btrAbortTrans(hlCrf)
                                            'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/Update Crf(3), Try Later", vbOkOnly + vbExclamation, "Erase")
                                            Exit Sub
                                        End If
                                        '9/30/16: Update pointers
                                        If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
                                            imCvfRecLen = Len(tmCvf)
                                            tmCvfSrchKey.lCode = tgAsgnVehCvf.lCode
                                            ilRet = btrGetEqual(hlCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                'If ilPFAsgn = 1 Then    'Final
                                                    tmCvf.iNextFinal(igAsgnVehRowNo) = tlCnf.iInstrNo + 1
                                                    tmCvf.iNextPrelim(igAsgnVehRowNo) = tlCnf.iInstrNo + 1
                                                'Else
                                                '    tmCvf.iNextPrelim(igAsgnVehRowNo) = tmCnf.iInstrNo + 1
                                                'End If
                                                ilRet = btrUpdate(hlCvf, tmCvf, imCvfRecLen)
                                                tgAsgnVehCvf = tmCvf
                                            End If
                                        End If

                                        'If ilPFAsgn = 1 Then    'Final
                                            'ilRet = btrGetPosition(hlCif, llCifRecPos)
                                            Do
                                                'ilRet = btrGetDirect(hlCif, tlCif, ilCifRecLen, llCifRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                tlCifSrchKey.lCode = llCifCode
                                                ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    ilCRet = btrAbortTrans(hlCrf)
                                                    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetDirect Cif(4), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                    Exit Sub
                                                End If
                                                'tmSRec = tlCif
                                                'ilRet = gGetByKeyForUpdate("Cif", hlCif, tmSRec)
                                                'tlCif = tmSRec
                                                'If ilRet <> BTRV_ERR_NONE Then
                                                '    ilCRet = btrAbortTrans(hlCrf)
                                                '    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Cif(5), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                '    Exit Sub
                                                'End If
                                                If (tlCif.iUsedDate(0) = 0) And (tlCif.iUsedDate(1) = 0) Then
                                                    tlCif.iUsedDate(0) = ilAsgnDate0
                                                    tlCif.iUsedDate(1) = ilAsgnDate1
                                                Else
                                                    gUnpackDate tlCif.iUsedDate(0), tlCif.iUsedDate(1), slDate
                                                    If llDate > gDateValue(slDate) Then
                                                        tlCif.iUsedDate(0) = ilAsgnDate0
                                                        tlCif.iUsedDate(1) = ilAsgnDate1
                                                    'Else
                                                    '    Exit Do 'Don't update date
                                                    End If
                                                End If
                                                tlCif.iNoTimesAir = tlCif.iNoTimesAir + 1
                                                'DL:7/1/03, Wrap value around to aviod overflow error
                                                If tlCif.iNoTimesAir > 32000 Then
                                                    tlCif.iNoTimesAir = 0
                                                End If
                                                ilRet = btrUpdate(hlCif, tlCif, ilCifRecLen)
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrAbortTrans(hlCrf)
                                                'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/Update Cif(6), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                Exit Sub
                                            End If
                                        'End If
                                        tlSdf.sPtType = "1"
                                        tlSdf.lCopyCode = tlCnf.lCifCode
                                        ilRotNo = tlCrf.iRotNo
                                        ilCopyReplaced = True
                                        ilSpotAsgn = True
                                        If Trim$(tlCrf.sZone) = Trim$(slZone) Then
                                            ilZoneFd = True
                                        End If
                                        '4/30/11:  Add Region copy by Airing vehicle
                                        tmCrf = tlCrf
                                    End If
                                End If
                                ilRet = btrEndTrans(hlCrf)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If ilSpotAsgn Then
            Exit Do
        End If
        ''Reposition to Crf so GetNext is correct
        'ilRet = btrGetDirect(hlCrf, tlCrf, ilCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
        If ilBypassCrf Then
            ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            tlCrfSrchKey4.sRotType = slType
            tlCrfSrchKey4.iEtfCode = 0
            tlCrfSrchKey4.iEnfCode = 0
            tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
            tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
            tlCrfSrchKey4.lFsfCode = 0
            tlCrfSrchKey4.iRotNo = ilSvRotNo
            ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
        End If
        Do While ilRet = BTRV_ERR_NONE
            blCrfFound = False
            For ilCrf = 0 To UBound(llProcessedCrfCode) - 1 Step 1
                If llProcessedCrfCode(ilCrf) = tlCrf.lCode Then
                    blCrfFound = True
                    Exit For
                End If
            Next ilCrf
            If Not blCrfFound Then
                Exit Do
            End If
            ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainAirCopy                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Function mObtainCrfAirCopy(blFirstTime As Boolean, ilType As Integer, slVefType As String, ilVefCode As Integer, ilVpfIndex As Integer, tlSdf As SDF, tlAvail As AVAILSS, hlCrf As Integer, hlCnf As Integer, hlCif As Integer, hlCvf As Integer, slZone As String) As Boolean
'
'
'   ilType(I)- 0=Airing only (Airing spots copy); 1=Specified vehicle (call by blackout replacement without copy)
'              2=Airing Only (Airing spot copy but zone must match); 3=Specified vehicle (like 1 except test only, return ilRotNo)
'              4=Airing Only (like 0 except test only, return ilRotNo);
'              6=Airing Only (Airing spot copy but zone must match)(like 2 except test only, return ilRotNo)
'
'
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llSpotTime As Long
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim ilAsgnDate0 As Integer
    Dim ilAsgnDate1 As Integer
    Dim slType As String
    'Dim llCrfRecPos As Long
    Dim llCrfCode As Long
    'Dim llCifRecPos As Long
    Dim llCifCode As Long
    Dim ilDay As Integer
    Dim ilSpotAsgn As Integer
    Dim ilAvailOk As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    'Copy rotation record information
    'Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim tlCrfSrchKey As LONGKEY0 'CIF key record image
    Dim tlCrfSrchKey4 As CRFKEY4 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    'Copy instruction record information
    Dim tlCnfSrchKey As CNFKEY0 'CNF key record image
    Dim ilCnfRecLen As Integer  'CNF record length
    Dim tlCnf As CNF            'CNF record image
    'Copy inventory record information
    Dim tlCifSrchKey As LONGKEY0 'CIF key record image
    Dim ilCifRecLen As Integer  'CIF record length
    Dim tlCif As CIF            'CIF record image
    Dim ilBypassCrf As Integer
    
    Dim slSvType As String
    Dim ilSvEtfCode As Integer
    Dim ilSvEnfCode As Integer
    Dim ilSvAdfCode As Integer
    Dim llSvChfCode As Long
    Dim llSvFsfCode As Long
    Dim blVefFound As Boolean
    Dim blCrfFound As Boolean
    Dim ilCvf As Integer
    Dim ilCrf As Integer
    Dim ilSvRotNo As Integer
    Dim ilSvVefCode As Integer
    ReDim llProcessedCrfCode(0 To 0) As Long
    
    mObtainCrfAirCopy = False
    If (ilType = 0) Or (ilType = 2) Or (ilType = 4) Or (ilType = 6) Then
        If (slVefType <> "A") Or (tgVpf(ilVpfIndex).sCopyOnAir <> "Y") Then
            Exit Function
        End If
    End If
    'Test if copy has been superseded and if so, replace copy definition within tlSdf
    'Find rotation to assign
    'Code later- test spot type to determine which rotation type
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    ilCrfRecLen = Len(tlCrf)
    ilCnfRecLen = Len(tlCnf)
    ilCifRecLen = Len(tlCif)
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
    ilDay = gWeekDayLong(llDate)
    gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llSpotTime
    ilSpotAsgn = False
    ilAvailOk = True
    slType = "A"
    If blFirstTime Then
        'tlCrfSrchKey1.sRotType = slType
        'tlCrfSrchKey1.iEtfCode = 0
        'tlCrfSrchKey1.iEnfCode = 0
        'tlCrfSrchKey1.iadfCode = tlSdf.iadfCode
        'tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
        'tlCrfSrchKey1.lFsfCode = 0
        'tlCrfSrchKey1.iVefCode = ilVefCode  'tmVef.iCode
        'tlCrfSrchKey1.iRotNo = 32000
        'ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        ilSvVefCode = ilVefCode
        slSvType = slType
        ilSvEtfCode = 0
        ilSvEnfCode = 0
        ilSvAdfCode = tlSdf.iAdfCode
        llSvChfCode = tlSdf.lChfCode
        llSvFsfCode = 0
        tlCrfSrchKey4.sRotType = slType
        tlCrfSrchKey4.iEtfCode = 0
        tlCrfSrchKey4.iEnfCode = 0
        tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
        tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
        tlCrfSrchKey4.lFsfCode = 0
        tlCrfSrchKey4.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Else
        'tlCrfSrchKey1.sRotType = tmCrf.sRotType
        'tlCrfSrchKey1.iEtfCode = tmCrf.iEtfCode
        'tlCrfSrchKey1.iEnfCode = tmCrf.iEnfCode
        'tlCrfSrchKey1.iadfCode = tmCrf.iadfCode
        'tlCrfSrchKey1.lChfCode = tmCrf.lChfCode
        'tlCrfSrchKey1.lFsfCode = tmCrf.lFsfCode
        'tlCrfSrchKey1.iVefCode = tmCrf.iVefCode
        'tlCrfSrchKey1.iRotNo = tmCrf.iRotNo
        'ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        ilSvVefCode = ilVefCode 'tmCrf.iVefCode
        slSvType = tmCrf.sRotType
        ilSvEtfCode = tmCrf.iEtfCode
        ilSvEnfCode = tmCrf.iEnfCode
        ilSvAdfCode = tmCrf.iAdfCode
        llSvChfCode = tmCrf.lChfCode
        llSvFsfCode = tmCrf.lFsfCode
        tlCrfSrchKey4.sRotType = tmCrf.sRotType
        tlCrfSrchKey4.iEtfCode = tmCrf.iEtfCode
        tlCrfSrchKey4.iEnfCode = tmCrf.iEnfCode
        tlCrfSrchKey4.iAdfCode = tmCrf.iAdfCode
        tlCrfSrchKey4.lChfCode = tmCrf.lChfCode
        tlCrfSrchKey4.lFsfCode = tmCrf.lFsfCode
        tlCrfSrchKey4.iRotNo = tmCrf.iRotNo
        ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = tmCrf.sRotType) And (tlCrf.iEtfCode = tmCrf.iEtfCode) And (tlCrf.iEnfCode = tmCrf.iEnfCode) And (tlCrf.iAdfCode = tmCrf.iAdfCode) And (tlCrf.lChfCode = tmCrf.lChfCode) 'And (tlCrf.iVefCode = tmCrf.iVefCode)   'tmVef.iCode)    'ilVefCode)
            If tlCrf.lCode = tmCrf.lCode Then
                ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Exit Do
            End If
            llProcessedCrfCode(UBound(llProcessedCrfCode)) = tlCrf.lCode
            ReDim Preserve llProcessedCrfCode(0 To UBound(llProcessedCrfCode) + 1) As Long
            ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) 'And (tlCrf.iVefCode = ilVefCode)   'tmVef.iCode)    'ilVefCode)
        'ilRet = btrGetPosition(hlCrf, llCrfRecPos)
        'Test date, time, day and zone
        llCrfCode = tlCrf.lCode
        ilSvRotNo = tlCrf.iRotNo
        llProcessedCrfCode(UBound(llProcessedCrfCode)) = tlCrf.lCode
        ReDim Preserve llProcessedCrfCode(0 To UBound(llProcessedCrfCode) + 1) As Long
        ilSpotAsgn = False
        ilBypassCrf = False
        If tlCrf.sState <> "D" Then
            'Note:  Only Recorded copy allowed with airing vehicles
            'Test if looking for Live or Recorded rotations
            'If slLive = "L" Then
            '    If tmCrf.sLiveCopy <> "L" Then
            '        ilBypassCrf = True
            '    End If
            'Else
                If (tlCrf.sLiveCopy = "L") Or (tlCrf.sLiveCopy = "M") Then
                    ilBypassCrf = True
                End If
            'End If
        Else
            ilBypassCrf = True
        End If
        If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo > tlCrf.iRotNo)) Then
            '1/4/15: Since CRF read with RotNo in desceasing order, once it is equal or low exit
            'ilBypassCrf = True
            Exit Function
        End If
        If Not ilBypassCrf Then
            If Not gCheckCrfVehicle(ilSvVefCode, tlCrf, hlCvf) Then
                ilBypassCrf = True
            End If
        End If
        If (tlCrf.sDay(ilDay) = "Y") And (tlSdf.iLen = tlCrf.iLen) And (Not ilBypassCrf) And ((Trim$(tlCrf.sZone) = "") Or (Trim$(tlCrf.sZone) = Trim$(slZone))) Then
            If (ilType = 0) Or (ilType = 1) Or (ilType = 3) Or (ilType = 4) Or (((ilType = 2) Or (ilType = 6)) And (Trim$(tlCrf.sZone) = Trim$(slZone))) Then
                gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slDate
                llSAsgnDate = gDateValue(slDate)
                gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slDate
                llEAsgnDate = gDateValue(slDate)
                If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                    gUnpackTime tlCrf.iStartTime(0), tlCrf.iStartTime(1), "A", "1", slTime
                    llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                    gUnpackTime tlCrf.iEndTime(0), tlCrf.iEndTime(1), "A", "1", slTime
                    llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                    If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                        ilAvailOk = True    'Ok to book into
                        If (tlCrf.sInOut = "I") Or (tlCrf.sInOut = "O") Then
                            'This line was required for assign copy to missed spots
                            'it is left in even though assigning to missed has been removed
                            If ((tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O")) And (tlSdf.iRotNo <> -1) Then    'Add spot
                                If tlCrf.sInOut = "I" Then
                                    If tlCrf.ianfCode <> tlAvail.ianfCode Then
                                        ilAvailOk = False   'No
                                    End If
                                Else
                                    If tlCrf.ianfCode = tlAvail.ianfCode Then
                                        ilAvailOk = False   'No
                                    End If
                                End If
                            End If
                        End If
                        If ilAvailOk Then
                            If Not ilSpotAsgn Then
                                If (ilType = 3) Or (ilType = 4) Or (ilType = 6) Then
                                    '4/30/11:  Add Region copy by Airing vehicle
                                    tmCrf = tlCrf
                                    mObtainCrfAirCopy = True
                                    Exit Function
                                End If
                                ilRet = btrBeginTrans(hlCrf, 1000)
                                'Assign copy, find next instruction to be assigned
                                '9/30/16: Handle sync copy
                                tlCnfSrchKey.lCrfCode = tlCrf.lCode
                                If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
                                    'If ilPFAsgn = 1 Then    'Final
                                        tlCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextFinal(igAsgnVehRowNo)
                                    'Else
                                    '    tmCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextPrelim(igAsgnVehRowNo)
                                    'End If
                                Else
                                    'If ilPFAsgn = 1 Then    'Final
                                        tlCnfSrchKey.iInstrNo = tlCrf.iNextFinal
                                    'Else
                                    '    tlCnfSrchKey.iInstrNo = tlCrf.iNextPrelim
                                    'End If
                                End If
                                ilRet = btrGetEqual(hlCnf, tlCnf, ilCnfRecLen, tlCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet <> BTRV_ERR_NONE Then
                                    tlCnfSrchKey.lCrfCode = tlCrf.lCode
                                    tlCnfSrchKey.iInstrNo = 1
                                    ilRet = btrGetEqual(hlCnf, tlCnf, ilCnfRecLen, tlCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                End If
                                If ilRet = BTRV_ERR_NONE Then
                                    'One inventory per instruction
                                    If tlCnf.lCifCode > 0 Then
                                        tlCifSrchKey.lCode = tlCnf.lCifCode
                                        ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        llCifCode = tlCif.lCode
                                    Else
                                        'Not coded
                                        ilRet = Not BTRV_ERR_NONE
                                    End If
                                    If ilRet = BTRV_ERR_NONE Then
                                        'ilRet = btrGetPosition(hlCrf, llCrfRecPos)
                                        Do
                                            'ilRet = btrGetDirect(hlCrf, tlCrf, ilCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                                            tlCrfSrchKey.lCode = llCrfCode
                                            ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrAbortTrans(hlCrf)
                                                'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetDirect Crf(1), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                Exit Function
                                            End If
                                            'tmSRec = tlCrf
                                            'ilRet = gGetByKeyForUpdate("Crf", hlCrf, tmSRec)
                                            'tlCrf = tmSRec
                                            'If ilRet <> BTRV_ERR_NONE Then
                                            '    ilCRet = btrAbortTrans(hlCrf)
                                            '    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Crf(2), Try Later", vbOkOnly + vbExclamation, "Erase")
                                            '    Exit Sub
                                            'End If
                                            'If ilPFAsgn = 1 Then    'Final
                                                tlCrf.iNextFinal = tlCnf.iInstrNo + 1
                                                tlCrf.iNextPrelim = tlCrf.iNextFinal
                                            'Else
                                            '    tlCrf.iNextPrelim = tlCnf.iInstrNo + 1
                                            'End If
                                            If (tlCrf.iEarliestDateAssg(0) = 0) And (tlCrf.iEarliestDateAssg(1) = 0) Then
                                                tlCrf.iEarliestDateAssg(0) = ilAsgnDate0
                                                tlCrf.iEarliestDateAssg(1) = ilAsgnDate1
                                                tlCrf.iLatestDateAssg(0) = ilAsgnDate0
                                                tlCrf.iLatestDateAssg(1) = ilAsgnDate1
                                            Else
                                                If (ilAsgnDate1 < tlCrf.iEarliestDateAssg(1)) Or ((ilAsgnDate1 = tlCrf.iEarliestDateAssg(1)) And (ilAsgnDate0 < tlCrf.iEarliestDateAssg(0))) Then
                                                    tlCrf.iEarliestDateAssg(0) = ilAsgnDate0
                                                    tlCrf.iEarliestDateAssg(1) = ilAsgnDate1
                                                End If
                                                If (ilAsgnDate1 > tlCrf.iLatestDateAssg(1)) Or ((ilAsgnDate1 = tlCrf.iLatestDateAssg(1)) And (ilAsgnDate0 > tlCrf.iLatestDateAssg(0))) Then
                                                    tlCrf.iLatestDateAssg(0) = ilAsgnDate0
                                                    tlCrf.iLatestDateAssg(1) = ilAsgnDate1
                                                End If
                                            End If
                                            tlCrf.iLastDateAssg(0) = ilAsgnDate0
                                            tlCrf.iLastDateAssg(1) = ilAsgnDate1
                                            gPackDate slNowDate, tlCrf.iDateAssgDone(0), tlCrf.iDateAssgDone(1)
                                            gPackTime slNowTime, tlCrf.iTimeAssgDone(0), tlCrf.iTimeAssgDone(1)
                                            ilRet = btrUpdate(hlCrf, tlCrf, ilCrfRecLen)
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilCRet = btrAbortTrans(hlCrf)
                                            'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/Update Crf(3), Try Later", vbOkOnly + vbExclamation, "Erase")
                                            Exit Function
                                        End If
                                        '9/30/16: Update pointers
                                        If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
                                            imCvfRecLen = Len(tmCvf)
                                            tmCvfSrchKey.lCode = tgAsgnVehCvf.lCode
                                            ilRet = btrGetEqual(hlCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                'If ilPFAsgn = 1 Then    'Final
                                                    tmCvf.iNextFinal(igAsgnVehRowNo) = tlCnf.iInstrNo + 1
                                                    tmCvf.iNextPrelim(igAsgnVehRowNo) = tlCnf.iInstrNo + 1
                                                'Else
                                                '    tmCvf.iNextPrelim(igAsgnVehRowNo) = tmCnf.iInstrNo + 1
                                                'End If
                                                ilRet = btrUpdate(hlCvf, tmCvf, imCvfRecLen)
                                                tgAsgnVehCvf = tmCvf
                                            End If
                                        End If

                                        'If ilPFAsgn = 1 Then    'Final
                                            'ilRet = btrGetPosition(hlCif, llCifRecPos)
                                            Do
                                                'ilRet = btrGetDirect(hlCif, tlCif, ilCifRecLen, llCifRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                tlCifSrchKey.lCode = llCifCode
                                                ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    ilCRet = btrAbortTrans(hlCrf)
                                                    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetDirect Cif(4), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                    Exit Function
                                                End If
                                                'tmSRec = tlCif
                                                'ilRet = gGetByKeyForUpdate("Cif", hlCif, tmSRec)
                                                'tlCif = tmSRec
                                                'If ilRet <> BTRV_ERR_NONE Then
                                                '    ilCRet = btrAbortTrans(hlCrf)
                                                '    'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Cif(5), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                '    Exit Sub
                                                'End If
                                                If (tlCif.iUsedDate(0) = 0) And (tlCif.iUsedDate(1) = 0) Then
                                                    tlCif.iUsedDate(0) = ilAsgnDate0
                                                    tlCif.iUsedDate(1) = ilAsgnDate1
                                                Else
                                                    gUnpackDate tlCif.iUsedDate(0), tlCif.iUsedDate(1), slDate
                                                    If llDate > gDateValue(slDate) Then
                                                        tlCif.iUsedDate(0) = ilAsgnDate0
                                                        tlCif.iUsedDate(1) = ilAsgnDate1
                                                    'Else
                                                    '    Exit Do 'Don't update date
                                                    End If
                                                End If
                                                tlCif.iNoTimesAir = tlCif.iNoTimesAir + 1
                                                'DL:7/1/03, Wrap value around to aviod overflow error
                                                If tlCif.iNoTimesAir > 32000 Then
                                                    tlCif.iNoTimesAir = 0
                                                End If
                                                ilRet = btrUpdate(hlCif, tlCif, ilCifRecLen)
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrAbortTrans(hlCrf)
                                                'ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/Update Cif(6), Try Later", vbOkOnly + vbExclamation, "Erase")
                                                Exit Function
                                            End If
                                        'End If
                                        tlSdf.sPtType = "1"
                                        tlSdf.lCopyCode = tlCnf.lCifCode
                                        tlSdf.iRotNo = tlCrf.iRotNo
                                        ilSpotAsgn = True
                                        '4/30/11:  Add Region copy by Airing vehicle
                                        tmCrf = tlCrf
                                    End If
                                End If
                                ilRet = btrEndTrans(hlCrf)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If ilSpotAsgn Then
            mObtainCrfAirCopy = True
            Exit Do
        End If
        ''Reposition to Crf so GetNext is correct
        'ilRet = btrGetDirect(hlCrf, tlCrf, ilCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
        'ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilBypassCrf Then
            ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            tlCrfSrchKey4.sRotType = slSvType
            tlCrfSrchKey4.iEtfCode = ilSvEtfCode
            tlCrfSrchKey4.iEnfCode = ilSvEnfCode
            tlCrfSrchKey4.iAdfCode = ilSvAdfCode
            tlCrfSrchKey4.lChfCode = llSvChfCode
            tlCrfSrchKey4.lFsfCode = llSvFsfCode
            tlCrfSrchKey4.iRotNo = ilSvRotNo
            ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
        End If
        Do While ilRet = BTRV_ERR_NONE
            blCrfFound = False
            For ilCrf = 0 To UBound(llProcessedCrfCode) - 1 Step 1
                If llProcessedCrfCode(ilCrf) = tlCrf.lCode Then
                    blCrfFound = True
                    Exit For
                End If
            Next ilCrf
            If Not blCrfFound Then
                Exit Do
            End If
            ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Loop
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBlackoutTest                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test and replaces spots if     *
'*                      required                       *
'*                                                     *
'*******************************************************
Public Sub gBlackoutTest(ilSource As Integer, hlCif As Integer, hlMcf As Integer, hlODF As Integer, hlRsf As Integer, hlCpf As Integer, hlCrf As Integer, hlCnf As Integer, hlClf As Integer, hlLst As Integer, hlCvf As Integer, slNewLines() As String * 72, hlMsg As Integer, lbcMsg As control)

'
'   ilSource(I)- 0=ExpNY; 1=Logs (ODF)
'
'   ilSource = 0 requires hlCif, hlMcf, slNewLines, hlMsg, lbcMsg
'   ilSource = 1 requires hlCif, hlMcf, hlOdf, hlLst, hlCpf, hlCrf
'
    Dim llLoop As Long
    Dim ilSBof As Integer
    Dim ilRBof As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilDay As Integer
    Dim ilFound As Integer
    Dim slCart As String
    Dim ilRet As Integer
    Dim ilStartBofIndex As Integer
    Dim slMsg As String
    Dim ilCompOk As Integer
    Dim llCompTime As Long
    Dim slLength As String
    Dim llStartAvailTime As Long
    Dim llEndAvailTime As Long
    Dim llRunTime As Long
    Dim llIndex As Long
    Dim ilVpfIndex As Integer
    Dim ilLnVpfIndex As Integer
    Dim ilVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim ilSuppress As Integer
    Dim slZone As String
    Dim ilZoneFd As Integer
    Dim ilCopyReplaced As Integer
    Dim lSvCifCode As Long
    Dim ilLen As Integer
    Dim ilSeqNo As Integer
    Dim llTime As Long
    Dim ilMatch As Integer
    Dim ilCrfVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilCLnVefCode As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim slTime As String
    Dim slVehName As String
    Dim ilVef As Integer
    Dim ilRotNo As Integer
    Dim ilRotNoVefCode As Integer
    Dim ilTestRotNo As Integer
    Dim llChfCode As Long
    Dim ilRsfExist As Integer
    Dim ilSBofExist As Integer
    Dim ilRBofExist As Integer
    Dim llRsfCode As Long
    Dim llSBofCode As Long
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim slLogType As String
    Dim ilAlert As Integer
    Dim ilEPass As Integer
    Dim tlBofRec As BOFREC
    imCifRecLen = Len(tmCif)
    imMcfRecLen = Len(tmMcf)
    imOdfRecLen = Len(tmOdf)
    imLstRecLen = Len(tmLst)
    imCpfRecLen = Len(tmCpf)
    imRsfRecLen = Len(tmRsf)
    'llDate = gDateValue(slDate)
    'ilDay = gWeekDayStr(slDate)
    'If tgVpf(ilVpfIndex).sSCompType = "T" Then
    '    gUnpackLength tgVpf(ilVpfIndex).iSCompLen(0), tgVpf(ilVpfIndex).iSCompLen(1), "3", False, slLength
    '    llCompTime = CLng(gLengthToCurrency(slLength))
    'Else
    '    llCompTime = 0&
    'End If
    ilVpfIndex = -1
    'Check Replacements
    For llLoop = LBound(tgSpotSum) To UBound(tgSpotSum) - 1 Step 1
        ilVefCode = tgSpotSum(llLoop).iVefCode
        llDate = tgSpotSum(llLoop).lDate
        ilDay = gWeekDayLong(llDate)
        If ilVpfIndex <> -1 Then
            If ilVefCode <> tgVpf(ilVpfIndex).iVefKCode Then
                ilVpfIndex = -1
            End If
        End If
        If ilVpfIndex = -1 Then
            'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
            '    If ilVefCode = tgVpf(llIndex).iVefKCode Then
                llIndex = gBinarySearchVpf(ilVefCode)
                If llIndex <> -1 Then
                    ilVpfIndex = llIndex
            '        Exit For
                End If
            'Next llIndex
        End If
        If ilVpfIndex = -1 Then
            Exit Sub
        End If
        'Test if Blackout previously defined- if so use it again
        ilRsfExist = False
        ilSBofExist = False
        ilRBofExist = False
        If ilSource = 1 Then
            '7/15/14
            tmRsfSrchKey4.lSdfCode = tgSpotSum(llLoop).lSdfCode
            tmRsfSrchKey4.sType = "B"
            tmRsfSrchKey4.iBVefCode = ilVefCode
            ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tgSpotSum(llLoop).lSdfCode) And (tmRsf.sType = "B")
                If (tmRsf.iBVefCode = ilVefCode) Then
                    If tmRsf.lRBofCode = tmRsf.lSBofCode Then
                    Else
                    End If
                    ilRsfExist = True
                    llRsfCode = tmRsf.lCode
                    For ilSBof = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
                        If tmRsf.lSBofCode = tgSBofRec(ilSBof).tBof.lCode Then
                            ilSBofExist = True
                            llSBofCode = tmRsf.lSBofCode
                            tlBofRec = tgSBofRec(ilSBof)
                            If tmRsf.sPtType = "1" Then
                                tlBofRec.tBof.lCifCode = tmRsf.lCopyCode
                            Else
                                tlBofRec.tBof.lCifCode = 0
                            End If
                            Exit For
                        End If
                    Next ilSBof
                    If tmRsf.lSBofCode <> tmRsf.lRBofCode Then
                        For ilRBof = LBound(tgRBofRec) To UBound(tgRBofRec) - 1 Step 1
                            If tmRsf.lRBofCode = tgRBofRec(ilRBof).tBof.lCode Then
                                ilRBofExist = True
                                tlBofRec = tgRBofRec(ilRBof)
                                tlBofRec.tBof.iRAdfCode = tlBofRec.tBof.iAdfCode
                                If tmRsf.sPtType = "1" Then
                                    tlBofRec.tBof.lCifCode = tmRsf.lCopyCode
                                Else
                                    tlBofRec.tBof.lCifCode = 0
                                End If
                                Exit For
                            End If
                        Next ilRBof
                    Else
                        'Replace is part of suppress, now check that replace contract not changed
                        If ilSBofExist Then
                            If tlBofRec.tBof.lRChfCode <> 0 Then
                                If tmRsf.lRChfCode = tlBofRec.tBof.lRChfCode Then
                                    ilRBofExist = ilSBofExist
                                End If
                            Else
                                ilRBofExist = ilSBofExist
                            End If
                        Else
                            ilRBofExist = ilSBofExist
                        End If
                    End If
                    Exit Do
                Else
                    Exit Do
                End If
                ilRet = btrGetNext(hlRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            ''If (Not ilSBofExist) Or (Not ilRBofExist) Then
            'If (ilSBofExist = False) And (ilRBofExist = False) Then
            If ((ilSBofExist = False) And (ilRBofExist = False)) Or ((tmRsf.lSBofCode <> tmRsf.lRBofCode) And (ilRBofExist = False) And (ilSBofExist)) Then
                If ilRsfExist Then
                    'Remove rsf
                    tmRsfSrchKey.lCode = llRsfCode
                    ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrDelete(hlRsf)
                        ilRsfExist = False
                    End If
                End If
            End If
        End If
        slZone = tgSpotSum(llLoop).sZone
        ilLen = tgSpotSum(llLoop).iLen
        ilSeqNo = tgSpotSum(llLoop).iSeqNo
        llTime = tgSpotSum(llLoop).lTime
        ilLnVefCode = tgSpotSum(llLoop).iLnVefCode
        slLogType = tgSpotSum(llLoop).sLogType
        If (Not ilRsfExist) Then
            If tgVpf(ilVpfIndex).sSCompType = "T" Then
                gUnpackLength tgVpf(ilVpfIndex).iSCompLen(0), tgVpf(ilVpfIndex).iSCompLen(1), "3", False, slLength
                llCompTime = CLng(gLengthToCurrency(slLength))
            Else
                llCompTime = 0&
            End If
            If ilSBofExist Then
                ilSPass = 1
                ilEPass = 2
            Else
                ilSPass = 1
                ilEPass = 1
            End If
            For ilPass = ilSPass To ilEPass Step 1
                'Is this spot suppressed
                For ilSBof = LBound(tgSBofRec) To UBound(tgSBofRec) - 1 Step 1
                    ilMatch = False
                    If (ilSBofExist) Then
                        If llSBofCode = tgSBofRec(ilSBof).tBof.lCode Then
                            ilMatch = True
                        End If
                    Else
                    If ilSource = 1 Then
                        If (tgSBofRec(ilSBof).tBof.iVefCode = ilVefCode) And (tgSBofRec(ilSBof).tBof.iAdfCode = tgSpotSum(llLoop).iAdfCode) And (tgSBofRec(ilSBof).tBof.sDays(ilDay) = "Y") Then
                            ilMatch = True
                        End If
                        If (tgSBofRec(ilSBof).tBof.iVefCode = ilVefCode) And (tgSBofRec(ilSBof).tBof.iAdfCode = 0) And (tgSBofRec(ilSBof).tBof.sDays(ilDay) = "Y") Then
                            ilMatch = True
                        End If
                        If (tgSBofRec(ilSBof).tBof.iVefCode = 0) And (tgSBofRec(ilSBof).tBof.iAdfCode = tgSpotSum(llLoop).iAdfCode) And (tgSBofRec(ilSBof).tBof.sDays(ilDay) = "Y") Then
                            ilMatch = True
                        End If
                        If (tgSBofRec(ilSBof).tBof.iVefCode = 0) And (tgSBofRec(ilSBof).tBof.iAdfCode = 0) And (tgSBofRec(ilSBof).tBof.sDays(ilDay) = "Y") Then
                            ilMatch = True
                        End If
                    Else
                        If (tgSBofRec(ilSBof).tBof.iVefCode = ilVefCode) And (tgSBofRec(ilSBof).tBof.iAdfCode = tgSpotSum(llLoop).iAdfCode) And (tgSBofRec(ilSBof).tBof.sDays(ilDay) = "Y") Then
                            ilMatch = True
                        End If
                    End If
                    End If
                    If ilMatch Then
                        gUnpackDateLong tgSBofRec(ilSBof).tBof.iStartDate(0), tgSBofRec(ilSBof).tBof.iStartDate(1), llSDate
                        gUnpackDateLong tgSBofRec(ilSBof).tBof.iEndDate(0), tgSBofRec(ilSBof).tBof.iEndDate(1), llEDate
                        If llEDate = 0 Then
                            llEDate = 999999999
                        End If
                        gUnpackTimeLong tgSBofRec(ilSBof).tBof.iStartTime(0), tgSBofRec(ilSBof).tBof.iStartTime(1), False, llSTime
                        gUnpackTimeLong tgSBofRec(ilSBof).tBof.iEndTime(0), tgSBofRec(ilSBof).tBof.iEndTime(1), True, llETime
                        ilSuppress = False
                        If (llDate >= llSDate) And (llDate <= llEDate) And (tgSpotSum(llLoop).lTime >= llSTime) And (tgSpotSum(llLoop).lTime <= llETime) Then
                            'If (tgSBofRec(ilSBof).tBof.lSifCode = 0) Or ((StrComp(Trim$(tgSBofRec(ilSBof).sShtTitle), Trim$(tgSpotSum(llLoop).sShortTitle), 1) = 0) And (tgSpf.sUseProdSptScr = "P")) Or ((StrComp(Trim$(tgSBofRec(ilSBof).sShtTitle), Trim$(tgSpotSum(llLoop).sProduct), 1) = 0) And (tgSpf.sUseProdSptScr <> "P")) Then
                            If (tgSBofRec(ilSBof).tBof.lSifCode = 0) Or ((StrComp(Trim$(tgSBofRec(ilSBof).sShtTitle), Trim$(tgSpotSum(llLoop).sShortTitle), 1) = 0) And (tgSpf.sUseProdSptScr = "P") And (ilSource = 0)) Or ((StrComp(Trim$(tgSBofRec(ilSBof).sShtTitle), Trim$(tgSpotSum(llLoop).sProduct), 1) = 0) And (tgSpf.sUseProdSptScr <> "P") And (ilSource = 0)) Or ((tgSBofRec(ilSBof).tBof.lSifCode <> 0) And (ilSource = 1)) Then
                                If (tgSBofRec(ilSBof).tBof.iLen = 0) Or (ilLen = tgSBofRec(ilSBof).tBof.iLen) Then
                                    If (tgSBofRec(ilSBof).tBof.lSChfCode = 0) Or (tgSpotSum(llLoop).lChfCode = tgSBofRec(ilSBof).tBof.lSChfCode) Then
                                        ilSuppress = True
                                    End If
                                End If
                            End If
                        End If
                        If ilSuppress Then
                            If (tgSBofRec(ilSBof).tBof.lRChfCode <> 0) And (ilSource = 1) Then
                                tlBofRec = tgSBofRec(ilSBof)
                                '8/24/06:  Added test to determine if copy is Ok
                                'tlBofRec.tBof.lCifCode = 0
                                If tlBofRec.tBof.lCifCode > 0 Then
                                    tmCifSrchKey.lCode = tlBofRec.tBof.lCifCode
                                    ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        If (tmCif.iAdfCode <> tlBofRec.tBof.iRAdfCode) Or (tmCif.sPurged <> "A") Then
                                            tlBofRec.tBof.lCifCode = 0
                                        End If
                                    Else
                                        tlBofRec.tBof.lCifCode = 0
                                    End If
                                End If
                                '6/4/16: Replaced GoSub
                                'GoSub ReplaceSpot
                                mReplaceSpot ilSource, tlBofRec, ilVefCode, ilLnVefCode, ilVpfIndex, ilLnVpfIndex, llDate, slDate, llTime, slTime, ilSeqNo, ilLen, hlODF, hlMcf, hlClf, hlCrf, hlCnf, hlCif, hlCvf, hlCpf, hlLst, hlRsf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo, ilTestRotNo, ilRotNoVefCode, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode, slVehName, llLoop, ilSBof, slLogType, slNewLines(), hlMsg, lbcMsg, ilRsfExist, ilAlert, ilFound
                            Else
                                'Find replacement
                                ilFound = False
                                If UBound(tgRBofRec) > LBound(tgRBofRec) Then
                                    If ilLen = 30 Then
                                        ilStartBofIndex = ig30StartBofIndex + 1
                                    ElseIf ilLen = 60 Then
                                        ilStartBofIndex = ig60StartBofIndex + 1
                                    Else
                                        ilStartBofIndex = igStartBofIndex + 1
                                    End If
                                    If ilStartBofIndex >= UBound(tgRBofRec) Then
                                        ilStartBofIndex = LBound(tgRBofRec)
                                    End If
                                    ilRBof = ilStartBofIndex
                                    Do
                                        If ((ilSource <> 1) Or (tgRBofRec(ilRBof).tBof.iVefCode = 0) Or (tgSpotSum(llLoop).iVefCode = tgRBofRec(ilRBof).tBof.iVefCode)) And ((ilLen = tgRBofRec(ilRBof).iLen) Or ((tgRBofRec(ilRBof).iLen = 0) And (ilSource = 1))) And (tgRBofRec(ilRBof).tBof.sDays(ilDay) = "Y") Then
                                            gUnpackDateLong tgRBofRec(ilRBof).tBof.iStartDate(0), tgRBofRec(ilRBof).tBof.iStartDate(1), llSDate
                                            gUnpackDateLong tgRBofRec(ilRBof).tBof.iEndDate(0), tgRBofRec(ilRBof).tBof.iEndDate(1), llEDate
                                            If llEDate = 0 Then
                                                llEDate = 999999999
                                            End If
                                            gUnpackTimeLong tgRBofRec(ilRBof).tBof.iStartTime(0), tgRBofRec(ilRBof).tBof.iStartTime(1), False, llSTime
                                            gUnpackTimeLong tgRBofRec(ilRBof).tBof.iEndTime(0), tgRBofRec(ilRBof).tBof.iEndTime(1), True, llETime
                                            If (llDate >= llSDate) And (llDate <= llEDate) And (tgSpotSum(llLoop).lTime >= llSTime) And (tgSpotSum(llLoop).lTime <= llETime) Then
                                                'Product protection test
                                                ilCompOk = False
                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(0) = 0) And (tgRBofRec(ilRBof).tBof.iMnfComp(1) = 0) Then
                                                    ilCompOk = True
                                                Else
                                                    If (tgVpf(ilVpfIndex).sSCompType = "T") And (llCompTime <= 0) Then
                                                        ilCompOk = True
                                                    Else
                                                        If (tgVpf(ilVpfIndex).sSCompType = "N") Then    'N="Not Back to Back"
                                                            'Look one spot up and one spot down- ignore time
                                                            llIndex = llLoop - 1
                                                            ilCompOk = True
                                                            If llIndex >= lgStartIndex Then
                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                    ilCompOk = False
                                                                End If
                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                    ilCompOk = False
                                                                End If
                                                            End If
                                                            llIndex = llLoop + 1
                                                            If llIndex < lgEndIndex Then
                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                    ilCompOk = False
                                                                End If
                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                    ilCompOk = False
                                                                End If
                                                            End If
                                                        Else
                                                            'Check within same avail
                                                            ilCompOk = True
                                                            llRunTime = tgSpotSum(llLoop).lTime
                                                            llIndex = llLoop - 1
                                                            Do While llIndex >= lgStartIndex
                                                                If tgSpotSum(llIndex).lTime + tgSpotSum(llIndex).iLen >= llRunTime Then
                                                                    If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                        ilCompOk = False
                                                                        Exit Do
                                                                    End If
                                                                    If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                        ilCompOk = False
                                                                        Exit Do
                                                                    End If
                                                                Else
                                                                    Exit Do
                                                                End If
                                                                llRunTime = tgSpotSum(llIndex).lTime
                                                                llIndex = llIndex - 1
                                                            Loop
                                                            If ilCompOk Then
                                                                llRunTime = tgSpotSum(llLoop).lTime + ilLen
                                                                llIndex = llLoop + 1
                                                                Do While llIndex < lgEndIndex
                                                                    If tgSpotSum(llIndex).lTime <= llRunTime Then
                                                                        If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                            ilCompOk = False
                                                                            Exit Do
                                                                        End If
                                                                        If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                            ilCompOk = False
                                                                            Exit Do
                                                                        End If
                                                                    Else
                                                                        Exit Do
                                                                    End If
                                                                    llRunTime = tgSpotSum(llIndex).lTime + tgSpotSum(llIndex).iLen
                                                                    llIndex = llIndex + 1
                                                                Loop
                                                                If ilCompOk Then
                                                                    'Check From llTime-llCompTime to llTime+llCompTime
                                                                    If llCompTime > 0 Then
                                                                        llStartAvailTime = tgSpotSum(llLoop).lTime - llCompTime
                                                                        If llStartAvailTime < 0 Then
                                                                            llStartAvailTime = 0
                                                                        End If
                                                                        llEndAvailTime = tgSpotSum(llLoop).lTime + llCompTime
                                                                        If llEndAvailTime > 86400 Then
                                                                            llEndAvailTime = 86400
                                                                        End If
                                                                        llIndex = llLoop - 1
                                                                        Do While llIndex >= lgStartIndex
                                                                            If tgSpotSum(llIndex).lTime > llStartAvailTime Then
                                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                                    ilCompOk = False
                                                                                    Exit Do
                                                                                End If
                                                                                If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                                    ilCompOk = False
                                                                                    Exit Do
                                                                                End If
                                                                            Else
                                                                                Exit Do
                                                                            End If
                                                                            llIndex = llIndex - 1
                                                                        Loop
                                                                        If ilCompOk Then
                                                                            llIndex = llLoop + 1
                                                                            Do While llIndex < lgEndIndex
                                                                                If tgSpotSum(llIndex).lTime <= llEndAvailTime Then
                                                                                    If (tgRBofRec(ilRBof).tBof.iMnfComp(0) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(0) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                                        ilCompOk = False
                                                                                        Exit Do
                                                                                    End If
                                                                                    If (tgRBofRec(ilRBof).tBof.iMnfComp(1) <> 0) And ((tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(0)) Or (tgRBofRec(ilRBof).tBof.iMnfComp(1) = tgSpotSum(llIndex).iMnfComp(1))) Then
                                                                                        ilCompOk = False
                                                                                        Exit Do
                                                                                    End If
                                                                                Else
                                                                                    Exit Do
                                                                                End If
                                                                                llIndex = llIndex + 1
                                                                            Loop
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                If ilCompOk Then
                                                    'If (ilSource = 0) Or ((ilSource = 1) And ((tgRBofRec(ilRBof).tBof.lRChfCode = 0) Or (tgRBofRec(ilRBof).tBof.lRChfCode = tgSpotSum(llLoop).lChfCode))) Then
                                                    If (ilSource = 0) Or ((ilSource = 1) And ((tgRBofRec(ilRBof).tBof.lRChfCode <> 0) And (tgRBofRec(ilRBof).tBof.lRChfCode <> tgSpotSum(llLoop).lChfCode))) Then
                                                        tlBofRec = tgRBofRec(ilRBof)
                                                        tlBofRec.tBof.iRAdfCode = tlBofRec.tBof.iAdfCode
                                                        'If tlBofRec.tBof.lRChfCode = 0 Then
                                                        '    tlBofRec.tBof.lRChfCode = tgSpotSum(llLoop).lChfCode
                                                        'End If
                                                        '6/4/16: Replaced GoSub
                                                        'GoSub ReplaceSpot
                                                        mReplaceSpot ilSource, tlBofRec, ilVefCode, ilLnVefCode, ilVpfIndex, ilLnVpfIndex, llDate, slDate, llTime, slTime, ilSeqNo, ilLen, hlODF, hlMcf, hlClf, hlCrf, hlCnf, hlCif, hlCvf, hlCpf, hlLst, hlRsf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo, ilTestRotNo, ilRotNoVefCode, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode, slVehName, llLoop, ilSBof, slLogType, slNewLines(), hlMsg, lbcMsg, ilRsfExist, ilAlert, ilFound
                                                        If ilLen = 30 Then
                                                            ig30StartBofIndex = ilRBof
                                                        ElseIf ilLen = 60 Then
                                                            ig60StartBofIndex = ilRBof
                                                        Else
                                                            igStartBofIndex = ilRBof
                                                        End If
                                                        Exit Do
                                                    End If
                                                End If
                                            End If
                                        End If
                                        ilRBof = ilRBof + 1
                                        If ilRBof >= UBound(tgRBofRec) Then
                                            ilRBof = LBound(tgRBofRec)
                                        End If
                                    Loop Until ilRBof = ilStartBofIndex
                                End If
                                If Not ilFound Then
                                    'Output error message
                                    If ilSource = 0 Then
                                        slMsg = "No Replacement for: " & slNewLines(tgSpotSum(llLoop).iNewIndex)
                                        Print #hlMsg, slMsg
                                        lbcMsg.AddItem slMsg
                                    Else
                                        '6/4/16: Replaced GoSub
                                        'GoSub ReplaceSpot
                                        mReplaceSpot ilSource, tlBofRec, ilVefCode, ilLnVefCode, ilVpfIndex, ilLnVpfIndex, llDate, slDate, llTime, slTime, ilSeqNo, ilLen, hlODF, hlMcf, hlClf, hlCrf, hlCnf, hlCif, hlCvf, hlCpf, hlLst, hlRsf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo, ilTestRotNo, ilRotNoVefCode, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode, slVehName, llLoop, ilSBof, slLogType, slNewLines(), hlMsg, lbcMsg, ilRsfExist, ilAlert, ilFound
                                        '6/4/16: Replaced GoSub
                                        'GoSub ReplaceMissing
                                        mReplaceMissing ilSource, ilVefCode, ilVpfIndex, llDate, llTime, slZone, ilSeqNo, hlODF
                                        slTime = gFormatTimeLong(llTime, "A", "1")
                                        slVehName = ""
                                        'For ilVef = 0 To UBound(tgVehicle) - 1 Step 1
                                        '    slNameCode = tgVehicle(ilVef).sKey 'lbcVehCode.List(llLoop)
                                        '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                        '    If Val(slCode) = ilVefCode Then
                                        '        ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                                        '        ilRet = gParseItem(slVehName, 3, "|", slVehName)
                                        '        Exit For
                                        '    End If
                                        'Next ilVef
                                        ilVef = gBinarySearchVef(ilVefCode)
                                        If ilVef <> -1 Then
                                            slVehName = tgMVef(ilVef).sName
                                        Else
                                            slVehName = "Vehicle" & str(ilVefCode) & " Code missing"
                                        End If
                                        slMsg = "No Replacement Found: " & Trim$(slVehName) & " " & Format$(llDate, "m/d/yy") & " at " & slTime
                                        Print #hlMsg, slMsg
                                        lbcMsg.AddItem slMsg
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ilSBof
                If ilRsfExist Then
                    Exit For
                End If
            Next ilPass
        Else
            If tlBofRec.tBof.lCifCode <= 0 Then
                tmRsfSrchKey.lCode = llRsfCode
                ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hlRsf)
                    ilRsfExist = False
                End If
            End If
            tlBofRec.tBof.lRChfCode = tmRsf.lRChfCode
            '6/4/16: Replaced GoSub
            'GoSub ReplaceSpot
            mReplaceSpot ilSource, tlBofRec, ilVefCode, ilLnVefCode, ilVpfIndex, ilLnVpfIndex, llDate, slDate, llTime, slTime, ilSeqNo, ilLen, hlODF, hlMcf, hlClf, hlCrf, hlCnf, hlCif, hlCvf, hlCpf, hlLst, hlRsf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo, ilTestRotNo, ilRotNoVefCode, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode, slVehName, llLoop, ilSBof, slLogType, slNewLines(), hlMsg, lbcMsg, ilRsfExist, ilAlert, ilFound
        End If
    Next llLoop
    Exit Sub
'ReplaceSpot:
'    lSvCifCode = tlBofRec.tBof.lCifCode
'    llChfCode = tlBofRec.tBof.lRChfCode
'    If tlBofRec.tBof.lCifCode <= 0 Then
'        'Assign copy
'        'Make-up spot and tmAvailTest to obtain Copy
'        tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
'        gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
'        'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'        gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'        tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
'        tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
'        ilRet = btrGetEqual(hlOdf, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'        If (ilRet = BTRV_ERR_NONE) Then
'            tmSdf.lChfCode = tlBofRec.tBof.lRChfCode
'            tmSdf.iAdfCode = tlBofRec.tBof.iRAdfCode
'            tmSdf.iDate(0) = tmOdf.iAirDate(0)
'            tmSdf.iDate(1) = tmOdf.iAirDate(1)
'            tmSdf.iTime(0) = tmOdf.iAirTime(0)
'            tmSdf.iTime(1) = tmOdf.iAirTime(1)
'            tmSdf.iLen = ilLen
'            tmSdf.sSchStatus = "S"
'            tmSdf.iRotNo = 0
'            tmSdf.sPtType = ""
'            tmSdf.lCopyCode = 0
'            tmAvail.ianfCode = tmOdf.ianfCode
'            gObtainAirCopy 3, "", ilVefCode, ilVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
'            ilRotNoVefCode = ilVefCode
'            'tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
'            If (ilVefCode <> ilLnVefCode) Then  'Try Selling vehicle
'                ilLnVpfIndex = 0
'                'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
'                '    If ilLnVefCode = tgVpf(llIndex).iVefKCode Then
'                    llIndex = gBinarySearchVpf(ilLnVefCode)
'                    If llIndex <> -1 Then
'                        ilLnVpfIndex = llIndex
'                '        Exit For
'                    End If
'                'Next llIndex
'                gObtainAirCopy 3, "", ilLnVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
'                If ilTestRotNo > ilRotNo Then
'                    ilRotNoVefCode = ilLnVefCode
'                    ilRotNo = ilTestRotNo
'                End If
'                'tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
'            End If
'            ilRet = gGetCrfVefCode(hlClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode)
'            If ilPkgVefCode <> 0 Then
'                ilLnVpfIndex = 0
'                'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
'                '    If ilPkgVefCode = tgVpf(llIndex).iVefKCode Then
'                    llIndex = gBinarySearchVpf(ilPkgVefCode)
'                    If llIndex <> -1 Then
'                        ilLnVpfIndex = llIndex
'                '        Exit For
'                    End If
'                'Next llIndex
'                gObtainAirCopy 3, "", ilPkgVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
'                If ilTestRotNo > ilRotNo Then
'                    ilRotNoVefCode = ilPkgVefCode
'                End If
'            End If
'            'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
'            '    If ilRotNoVefCode = tgVpf(llIndex).iVefKCode Then
'                llIndex = gBinarySearchVpf(ilRotNoVefCode)
'                If llIndex <> -1 Then
'                    ilLnVpfIndex = llIndex
'            '        Exit For
'                End If
'            'Next llIndex
'            gObtainAirCopy 1, "", ilRotNoVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
'            tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
'        End If
'    End If
'    'Replace spot (Cart ID and Short Title)
'    tmCifSrchKey.lCode = tlBofRec.tBof.lCifCode
'    ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If (ilRet = BTRV_ERR_NONE) Then
'        If tgSpf.sUseCartNo <> "N" Then
'            If tmCif.iMcfCode > 0 Then
'                If tmMcf.iCode <> tmCif.iMcfCode Then
'                    tmMcfSrchKey.iCode = tmCif.iMcfCode
'                    ilRet = btrGetEqual(hlMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        tmMcf.sName = "C"
'                        tmMcf.sPrefix = "C"
'                    End If
'                End If
'            Else
'                tmMcf.sName = ""
'                tmMcf.sPrefix = ""
'            End If
'            If Trim$(tmCif.sCut) = "" Then
'                slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
'            Else
'                slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
'            End If
'        Else
'            slCart = ""
'            tmCif.lcpfCode = 0
'            ilRet = BTRV_ERR_NONE
'        End If
'    ElseIf (tlBofRec.tBof.lCifCode = 0) And (ilSource = 1) Then
'        slCart = ""
'        tmCif.lcpfCode = 0
'        ilRet = BTRV_ERR_NONE
'    End If
'    If (ilRet = BTRV_ERR_NONE) Then
'        If ilSource = 0 Then
'            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 24, Len(slCart)) = slCart
'            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 30, 15) = UCase$(tlBofRec.sShtTitle)
'            tmCpfSrchKey.lCode = tmCif.lcpfCode
'            ilRet = btrGetEqual(hlCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            If (ilRet <> BTRV_ERR_NONE) Then
'                tmCpf.sISCI = ""
'            End If
'            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 53, 20) = UCase$(tmCpf.sISCI)
'        Else
'            'Update ODF and LST
'            Do
'                'tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
'                'gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
'                'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'                'tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
'                'tmOdfSrchKey0.iSeqNo = tgSpotSum(llLoop).iSeqNo
'                tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
'                gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
'                'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'                gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'                tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
'                tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
'                ilRet = btrGetEqual(hlOdf, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                If (ilRet = BTRV_ERR_NONE) Then
'                    tmOdf.iAdfCode = tlBofRec.tBof.iRAdfCode
'                    tmOdf.lCifCode = tlBofRec.tBof.lCifCode
'                    If tgSpf.sUseProdSptScr = "P" Then
'                        tmOdf.sProduct = ""
'                        tmOdf.sShortTitle = tlBofRec.sShtTitle
'                        If (Trim$(tmOdf.sShortTitle) = "") Then
'                            slTime = gFormatTimeLong(llTime, "A", "1")
'                            slVehName = ""
'                            'For ilVef = 0 To UBound(tgVehicle) - 1 Step 1
'                            '    slNameCode = tgVehicle(ilVef).sKey 'lbcVehCode.List(llLoop)
'                            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                            '    If Val(slCode) = ilVefCode Then
'                            '        ilRet = gParseItem(slNameCode, 1, "\", slVehName)
'                            '        ilRet = gParseItem(slVehName, 3, "|", slVehName)
'                            '        Exit For
'                            '    End If
'                            'Next ilVef
'                            ilVef = gBinarySearchVef(ilVefCode)
'                            If ilVef <> -1 Then
'                                slVehName = tgMVef(ilVef).sName
'                            Else
'                                slVehName = "Vehicle" & str(ilVefCode) & " Code missing"
'                            End If
'                            slMsg = "Replacement Found but Short Title Missing: " & Trim$(slVehName) & " " & Format$(llDate, "m/d/yy") & " at " & slTime
'                            Print #hlMsg, slMsg
'                            lbcMsg.AddItem slMsg
'                        End If
'                    Else
'                        tmOdf.sProduct = tlBofRec.sShtTitle
'                        tmOdf.sShortTitle = ""
'                        If (tgSpotSum(llLoop).lLstCode > 0) Then
'                            tmOdf.sShortTitle = slCart
'                        End If
'                    End If
'                    tmOdf.lCntrNo = tlBofRec.lRCntrNo
'                    tmOdf.sBBDesc = ""
'                    ilRet = btrUpdate(hlOdf, tmOdf, imOdfRecLen)
'                Else
'                    Exit Do
'                End If
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If Not ilRsfExist Then
'                tmRsf.lCode = 0
'                tmRsf.lSdfCode = tgSpotSum(llLoop).lSdfCode
'                If tlBofRec.tBof.lCifCode > 0 Then
'                    tmRsf.sPtType = "1"
'                    tmRsf.lCopyCode = tlBofRec.tBof.lCifCode
'                Else
'                    tmRsf.sPtType = "0"
'                    tmRsf.lCopyCode = 0
'                    tmRsf.iRotNo = 0
'                End If
'                tmRsf.lRafCode = 0
'                tmRsf.lSBofCode = tgSBofRec(ilSBof).tBof.lCode
'                tmRsf.lRBofCode = tlBofRec.tBof.lCode
'                tmRsf.sType = "B"
'                tmRsf.iBVefCode = ilVefCode
'                tmRsf.lRChfCode = llChfCode
'                gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
'                tmRsf.sUnused = ""
'                ilRet = btrInsert(hlRsf, tmRsf, imRsfRecLen, INDEXKEY0)
'                ilRsfExist = True
'            End If
'            If (tgSpotSum(llLoop).lLstCode > 0) And (ilRet = BTRV_ERR_NONE) Then
'                'Do
'                    tmLstSrchKey.lCode = tgSpotSum(llLoop).lLstCode
'                    ilRet = btrGetEqual(hlLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                    If (ilRet = BTRV_ERR_NONE) Then
'                        ilRet = btrDelete(hlLst)
'                        'If tmLst.lSdfCode > 0 Then
'                        '    tmLst.lSdfCode = -tmLst.lSdfCode
'                        'End If
'                        tmLst.lCntrNo = tlBofRec.lRCntrNo
'                        tmLst.iAdfCode = tlBofRec.tBof.iRAdfCode
'                        tmLst.iAgfCode = 0
'                        If tgSpf.sUseProdSptScr = "P" Then
'                            tmLst.sProd = ""
'                        Else
'                            tmLst.sProd = tlBofRec.sShtTitle
'                        End If
'                        tmLst.iLineNo = 0
'                        tmLst.iLnVefCode = 0
'                        tmLst.iStartDate(0) = 0
'                        tmLst.iStartDate(1) = 0
'                        tmLst.iEndDate(0) = 0
'                        tmLst.iEndDate(1) = 0
'                        tmLst.iDays(0) = 0
'                        tmLst.iDays(0) = 0
'                        tmLst.iDays(1) = 0
'                        tmLst.iDays(2) = 0
'                        tmLst.iDays(3) = 0
'                        tmLst.iDays(4) = 0
'                        tmLst.iDays(5) = 0
'                        tmLst.iDays(6) = 0
'                        tmLst.iSpotsWk = 0
'                        tmLst.iPriceType = 1
'                        tmLst.lPrice = 0
'                        tmLst.iSpotType = 4
'                        tmLst.sDemo = ""
'                        tmLst.lAud = 0
'                        tmLst.sISCI = ""
'                        If tmCif.lcpfCode > 0 Then
'                            tmCpfSrchKey.lCode = tmCif.lcpfCode
'                            ilRet = btrGetEqual(hlCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'                            If (ilRet = BTRV_ERR_NONE) Then
'                                tmLst.sISCI = tmCpf.sISCI
'                            Else
'                                tmLst.sISCI = ""
'                            End If
'                        Else
'                            tmLst.sISCI = ""
'                        End If
'                        tmLst.sCart = slCart
'                        tmLst.lcpfCode = tmCif.lcpfCode
'                        tmLst.lCrfCsfcode = 0
'                        tmLst.lCifCode = tlBofRec.tBof.lCifCode
'                        tmLst.sImportedSpot = "N"
'                        tmLst.lBkoutLstCode = 0
'                        'ilRet = btrUpdate(hlLst, tmLst, imLstRecLen)
'                        ilRet = btrInsert(hlLst, tmLst, imLstRecLen, INDEXKEY0)
'                        'Checking if date bewteen todays date and last log is not required
'                        'It is Ok if Alert set on when generating Final Logs (The Alert will be removed within mGenLog)
'                        slDate = Format$(llDate, "m/d/yy")
'                        ilAlert = gAlertAdd(slLogType, "S", 0, tmLst.iLogVefCode, slDate)
'                        ilAlert = gAlertAdd(slLogType, "I", 0, tmLst.iLogVefCode, slDate)
'                    Else
'                        'Exit Do
'                    End If
'                'Loop While ilRet = BTRV_ERR_CONFLICT
'            End If
'        End If
'        ilFound = True
'    End If
'    tlBofRec.tBof.lCifCode = lSvCifCode
'    Return
'RemoveSpot:
'    If ilSource = 1 Then
'        If ilRsfExist Then
'            tmRsfSrchKey.lCode = llRsfCode
'            ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'            If ilRet = BTRV_ERR_NONE Then
'                ilRet = btrDelete(hlRsf)
'                ilRsfExist = False
'            End If
'        End If
'        If (tgSpotSum(llLoop).lLstCode > 0) Then
'            'Do
'                tmLstSrchKey.lCode = tgSpotSum(llLoop).lLstCode
'                ilRet = btrGetEqual(hlLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                If (ilRet = BTRV_ERR_NONE) Then
'                    ilRet = btrDelete(hlLst)
'                    tmLst.iType = 1
'                    tmLst.lSdfCode = 0
'                    tmLst.lCntrNo = 0
'                    tmLst.iAdfCode = 0
'                    tmLst.iAgfCode = 0
'                    tmLst.sProd = ""
'                    tmLst.iLineNo = 0
'                    tmLst.iLnVefCode = 0
'                    tmLst.iStartDate(0) = 0
'                    tmLst.iStartDate(1) = 0
'                    tmLst.iEndDate(0) = 0
'                    tmLst.iEndDate(1) = 0
'                    tmLst.iDays(0) = 0
'                    tmLst.iDays(0) = 0
'                    tmLst.iDays(1) = 0
'                    tmLst.iDays(2) = 0
'                    tmLst.iDays(3) = 0
'                    tmLst.iDays(4) = 0
'                    tmLst.iDays(5) = 0
'                    tmLst.iDays(6) = 0
'                    tmLst.iSpotsWk = 0
'                    tmLst.iPriceType = 1
'                    tmLst.lPrice = 0
'                    tmLst.iSpotType = 4
'                    tmLst.sDemo = ""
'                    tmLst.lAud = 0
'                    tmLst.sISCI = ""
'                    tmLst.iUnits = 1
'                    tmLst.sISCI = ""
'                    tmLst.sCart = ""
'                    tmLst.lcpfCode = 0
'                    tmLst.lCrfCsfcode = 0
'                    tmLst.sImportedSpot = "N"
'                    tmLst.lBkoutLstCode = 0
'                    ilRet = btrInsert(hlLst, tmLst, imLstRecLen, INDEXKEY0)
'                    'Checking if date bewteen todays date and last log is not required
'                    'It is Ok if Alert set on when generating Final Logs (The Alert will be removed within mGenLog)
'                    slDate = Format$(llDate, "m/d/yy")
'                    ilAlert = gAlertAdd(slLogType, "S", 0, tmLst.iLogVefCode, slDate)
'                    ilAlert = gAlertAdd(slLogType, "I", 0, tmLst.iLogVefCode, slDate)
'                Else
'                    'Exit Do
'                End If
'            'Loop While ilRet = BTRV_ERR_CONFLICT
'        End If
'    End If
'    Return
'ReplaceMissing:
'    'If (ilSource = 1) And (tgSpf.sUseProdSptScr = "P") Then
'    If (ilSource = 1) Then
'        tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
'        gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
'        'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'        gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
'        tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
'        tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
'        ilRet = btrGetEqual(hlOdf, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'        If (ilRet = BTRV_ERR_NONE) Then
'            If tgVpf(ilVpfIndex).sUnsoldBlank = "N" Then
'                ilRet = btrDelete(hlOdf)
'            Else
'                tmOdf.iAdfCode = 0
'                tmOdf.lCifCode = 0
'                If tgSpf.sUseProdSptScr = "P" Then
'                    tmOdf.sProduct = ""
'                    tmOdf.sShortTitle = ""  '"Replace Missing"
'                Else
'                    tmOdf.sProduct = "" '"Replace Missing"
'                    tmOdf.sShortTitle = ""
'                End If
'                tmOdf.lCntrNo = 0
'                tmOdf.sBBDesc = ""
'                ilRet = btrUpdate(hlOdf, tmOdf, imOdfRecLen)
'            End If
'        End If
'    End If
'    Return
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadBsfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Public Function gReadBofRec(ilSource As Integer, hlBof As Integer, hlCif As Integer, hlPrf As Integer, hlSif As Integer, hlChf As Integer, slInType As String, slActiveDate As String, ilSort As Integer) As Integer
'
'   iRet = gReadBofRec (hlBof, hlCif, hlPrf, hlSif, hlChf, slInType, slActiveDate, ilSort)
'   Where:
'       ilSource(I)- 0=NY; 1=Log
'       slInType(I)-S=Suppression; R=Replacement; B=Both
'       ilSort(I)- 0= for Suppression [Vehicle, Advt; Short Title; Start Time];
'                     for Replacement [Advt; Short Title; Start Date]
'                  1= for Suppression same as 0
'                     for Replacement [Length; Random #]
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim slType As String
    Dim slSDate As String
    Dim llSTime As Long
    Dim slSTime As String
    Dim slStr As String
    Dim ilNum As Integer
    Dim slNum As String
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    If (slInType = "S") Then
        ilSPass = 1
        ilEPass = 1
    End If
    If (slInType = "R") Then
        ilSPass = 2
        ilEPass = 2
    End If
    If (slInType = "B") Then
        ilSPass = 1
        ilEPass = 2
    End If
    If ilSort = 1 Then
        Randomize
    End If
    For ilPass = ilSPass To ilEPass Step 1
        If ilPass = 1 Then
            'ReDim tgSBofRec(1 To 1) As BOFREC
            ReDim tgSBofRec(0 To 0) As BOFREC
            ilUpper = UBound(tgSBofRec)
            slType = "S"
        Else
            'ReDim tgRBofRec(1 To 1) As BOFREC
            ReDim tgRBofRec(0 To 0) As BOFREC
            ilUpper = UBound(tgRBofRec)
            ilExtLen = Len(tgRBofRec(1).tBof)  'Extract operation record size
            slType = "R"
        End If
        imBofRecLen = Len(tmBof)
        imSifRecLen = Len(tmSif)
        imPrfRecLen = Len(tmPrf)
        imCifRecLen = Len(tmCif)
        imCHFRecLen = Len(tmChf)
        ilExtLen = Len(tmBof)  'Extract operation record size
        btrExtClear hlBof   'Clear any previous extend operation
        tmBofSrchKey1.sType = slType
        tmBofSrchKey1.iEndDate(0) = 0
        tmBofSrchKey1.iEndDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hlBof, tmBof, imBofRecLen, tmBofSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'ilRet = btrGetFirst(hmBsf, tgBsfRec(1).tBsf, imBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hlBof, llNoRec, -1, "UC", "BOF", "") '"EG") 'Set extract limits (all records)
            tlCharTypeBuff.sType = slType
            ilOffSet = gFieldOffset("Bof", "BofType")
            ilRet = btrExtAddLogicConst(hlBof, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Bof", "BofEndDate")
            ilRet = btrExtAddLogicConst(hlBof, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_OR, tlDateTypeBuff, 4)
            gPackDate "", tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Bof", "BofEndDate")
            ilRet = btrExtAddLogicConst(hlBof, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            'On Error GoTo mReadBofRecErr
            'gBtrvErrorMsg ilRet, "mReadBofRec (btrExtAddLogicConst):" & "Bof.Btr", Blackout
            'On Error GoTo 0
            ilRet = btrExtAddField(hlBof, 0, ilExtLen) 'Extract the whole record
            'On Error GoTo mReadBofRecErr
            'gBtrvErrorMsg ilRet, "mReadBofRec (btrExtAddField):" & "Bof.Btr", Blackout
            'On Error GoTo 0
            ilRet = btrExtGetNext(hlBof, tmBof, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                'On Error GoTo mReadBofRecErr
                'gBtrvErrorMsg ilRet, "mReadBofRec (btrExtGetNextExt):" & "Bof.Btr", Blackout
                'On Error GoTo 0
                ilExtLen = Len(tmBof)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlBof, tmBof, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    If ((ilSource = 0) And (tmBof.sSource = "N")) Or ((ilSource = 1) And (tmBof.sSource <> "N")) Then
                        gUnpackDateForSort tmBof.iStartDate(0), tmBof.iStartDate(1), slSDate
                        Do While Len(slSDate) < 5
                            slSDate = "0" & slSDate
                        Loop
                        gUnpackTimeLong tmBof.iStartTime(0), tmBof.iStartTime(1), False, llSTime
                        slSTime = Trim$(str$(llSTime))
                        Do While Len(slSTime) < 5
                            slSTime = "0" & slSTime
                        Loop
                        If ilPass = 1 Then
                            tgSBofRec(ilUpper).tBof = tmBof
                            'Get Vehicle Name
                            tgSBofRec(ilUpper).sVefName = ""
                            If tgSBofRec(ilUpper).tBof.iVefCode > 0 Then
                                'For ilLoop = 0 To UBound(tlUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                                '    slNameCode = tlUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                                '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                '    If Val(slCode) = tgSBofRec(ilUpper).tBof.iVefCode Then
                                '        ilRet = gParseItem(slNameCode, 1, "\", slName)
                                '        ilRet = gParseItem(slName, 3, "|", tgSBofRec(ilUpper).sVefName)
                                '        Exit For
                                '    End If
                                'Next ilLoop
                                ilLoop = gBinarySearchVef(tgSBofRec(ilUpper).tBof.iVefCode)
                                If ilLoop <> -1 Then
                                    tgSBofRec(ilUpper).sVefName = tgMVef(ilLoop).sName
                                End If
                            End If
                            'Get Advertiser Name
                            tgSBofRec(ilUpper).sAdfName = ""
                            If tgSBofRec(ilUpper).tBof.iAdfCode > 0 Then
                            '    For ilLoop = 0 To UBound(tlAdvertiser) - 1 Step 1 'Traffic!lbcAdvt.ListCount - 1 Step 1
                            '        slNameCode = tlAdvertiser(ilLoop).sKey 'Traffic!lbcAdvt.List(ilLoop)
                            '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            '        If Val(slCode) = tgSBofRec(ilUpper).tBof.iAdfCode Then
                            '            ilRet = gParseItem(slNameCode, 1, "\", tgSBofRec(ilUpper).sAdfName)
                            '            Exit For
                            '        End If
                            '    Next ilLoop
                                ilLoop = gBinarySearchAdf(tgSBofRec(ilUpper).tBof.iAdfCode)
                                If ilLoop <> -1 Then
                                    If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                                        tgSBofRec(ilUpper).sAdfName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                                    Else
                                        tgSBofRec(ilUpper).sAdfName = tgCommAdf(ilLoop).sName
                                    End If
                                End If
                            End If
                            tgSBofRec(ilUpper).lSCntrNo = 0
                            tgSBofRec(ilUpper).sRAdfName = ""
                            tgSBofRec(ilUpper).lRCntrNo = 0
                            If ilSource = 1 Then
                                If tgSBofRec(ilUpper).tBof.lSChfCode > 0 Then
                                    tmChfSrchKey.lCode = tgSBofRec(ilUpper).tBof.lSChfCode
                                    ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgSBofRec(ilUpper).lSCntrNo = tmChf.lCntrNo
                                '        slName = Trim$(Str$(tmChf.lCntrNo)) & " " & Trim$(Str$(tmChf.iCntRevNo)) & "-" & Trim$(Str$(tmChf.iExtRevNo)) & " : "
                                '        gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStr
                                '        slName = slName & slStr
                                '        gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slStr
                                '        slName = slName & "-" & slStr & " " & Trim$(tmChf.sProduct)
                                '        If tmChf.lVefCode > 0 Then
                                '            For ilLoop = 0 To UBound(tlUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                                '                slNameCode = tlUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                                '                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                '                If Val(slCode) = tmChf.lVefCode Then
                                '                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                                '                    ilRet = gParseItem(slName, 3, "|", slStr)
                                '                    slName = slName & " " & slStr
                                '                    Exit For
                                '                End If
                                '            Next ilLoop
                                '        End If
                                '        tgSBofRec(ilUpper).sSCntr = slName
                                    End If
                                End If
                                'Get Replace Advertiser Name

                                If tgSBofRec(ilUpper).tBof.iRAdfCode > 0 Then
'                                    For ilLoop = 0 To UBound(tlAdvertiser) - 1 Step 1 'Traffic!lbcAdvt.ListCount - 1 Step 1
'                                        slNameCode = tlAdvertiser(ilLoop).sKey 'Traffic!lbcAdvt.List(ilLoop)
'                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                                        If Val(slCode) = tgSBofRec(ilUpper).tBof.iRAdfCode Then
'                                            ilRet = gParseItem(slNameCode, 1, "\", tgSBofRec(ilUpper).sRAdfName)
'                                            Exit For
'                                        End If
'                                    Next ilLoop
                                    ilLoop = gBinarySearchAdf(tgSBofRec(ilUpper).tBof.iRAdfCode)
                                    If ilLoop <> -1 Then
                                        If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                                            tgSBofRec(ilUpper).sRAdfName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                                        Else
                                            tgSBofRec(ilUpper).sRAdfName = tgCommAdf(ilLoop).sName
                                        End If
                                    End If
                                End If
                                If tgSBofRec(ilUpper).tBof.lRChfCode > 0 Then
                                    tmChfSrchKey.lCode = tgSBofRec(ilUpper).tBof.lRChfCode
                                    ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgSBofRec(ilUpper).lRCntrNo = tmChf.lCntrNo
                                        If tgSpf.sUseProdSptScr <> "P" Then  'Short Title
                                            tgSBofRec(ilUpper).sShtTitle = tmChf.sProduct
                                        Else
                                            If (tgSBofRec(ilUpper).tBof.lSifCode = 0) And (ilSource = 1) Then
                                                tgSBofRec(ilUpper).tBof.lSifCode = tmChf.lSifCode
                                            End If
                                            tmSifSrchKey.lCode = tgSBofRec(ilUpper).tBof.lSifCode
                                            ilRet = btrGetEqual(hlSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tgSBofRec(ilUpper).sShtTitle = tmSif.sName
                                            End If
                                        End If
                                '        slName = Trim$(Str$(tmChf.lCntrNo)) & " " & Trim$(Str$(tmChf.iCntRevNo)) & "-" & Trim$(Str$(tmChf.iExtRevNo)) & " : "
                                '        gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStr
                                '        slName = slName & slStr
                                '        gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slStr
                                '        slName = slName & "-" & slStr & " " & Trim$(tmChf.sProduct)
                                '        If tmChf.lVefCode > 0 Then
                                '            For ilLoop = 0 To UBound(tlUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
                                '                slNameCode = tlUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                                '                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                '                If Val(slCode) = tmChf.lVefCode Then
                                '                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                                '                    ilRet = gParseItem(slName, 3, "|", slStr)
                                '                    slName = slName & " " & slStr
                                '                    Exit For
                                '                End If
                                 '           Next ilLoop
                                '        End If
                                '        tgSBofRec(ilUpper).sRCntr = slName
                                    End If
                                End If
                            Else
                                tgSBofRec(ilUpper).sShtTitle = ""
                                If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                                    tmSifSrchKey.lCode = tgSBofRec(ilUpper).tBof.lSifCode
                                    ilRet = btrGetEqual(hlSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgSBofRec(ilUpper).sShtTitle = tmSif.sName
                                    End If
                                Else
                                    tmPrfSrchKey.lCode = tgSBofRec(ilUpper).tBof.lSifCode
                                    ilRet = btrGetEqual(hlPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgSBofRec(ilUpper).sShtTitle = tmPrf.sName
                                    End If
                                End If
                            End If
                            tgSBofRec(ilUpper).sKey = tgSBofRec(ilUpper).sVefName & tgSBofRec(ilUpper).sAdfName & tgSBofRec(ilUpper).sShtTitle & slSTime
                            tgSBofRec(ilUpper).iStatus = 1
                            tgSBofRec(ilUpper).lRecPos = llRecPos
                            ilUpper = ilUpper + 1
                            'ReDim Preserve tgSBofRec(1 To ilUpper) As BOFREC
                            ReDim Preserve tgSBofRec(0 To ilUpper) As BOFREC
                        Else
                            tgRBofRec(ilUpper).tBof = tmBof
                            tgRBofRec(ilUpper).lSCntrNo = 0
                            tgRBofRec(ilUpper).sRAdfName = ""
                            tgRBofRec(ilUpper).lRCntrNo = 0
                            tgRBofRec(ilUpper).sShtTitle = ""
                            tmChf.lSifCode = 0
                            If ilSource = 1 Then
                                If tgRBofRec(ilUpper).tBof.lRChfCode > 0 Then
                                    tmChfSrchKey.lCode = tgRBofRec(ilUpper).tBof.lRChfCode
                                    ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgRBofRec(ilUpper).lRCntrNo = tmChf.lCntrNo
                                        tgRBofRec(ilUpper).tBof.iMnfComp(0) = tmChf.iMnfComp(0)
                                        tgRBofRec(ilUpper).tBof.iMnfComp(1) = tmChf.iMnfComp(1)
                                        If tgSpf.sUseProdSptScr <> "P" Then  'Short Title
                                            tgRBofRec(ilUpper).sShtTitle = tmChf.sProduct
                                        End If
                                    End If
                                End If
                            End If
                            If tgSpf.sUseProdSptScr = "P" Then  'Short Title
                                If (tgRBofRec(ilUpper).tBof.lSifCode = 0) And (ilSource = 1) Then
                                    tgRBofRec(ilUpper).tBof.lSifCode = tmChf.lSifCode
                                End If
                                tmSifSrchKey.lCode = tgRBofRec(ilUpper).tBof.lSifCode
                                ilRet = btrGetEqual(hlSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tgRBofRec(ilUpper).sShtTitle = tmSif.sName
                                End If
                            Else
                                If (ilSource <> 1) Or (Trim$(tgRBofRec(ilUpper).sShtTitle) = "") Then
                                    tmPrfSrchKey.lCode = tgRBofRec(ilUpper).tBof.lSifCode
                                    ilRet = btrGetEqual(hlPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tgRBofRec(ilUpper).sShtTitle = tmPrf.sName
                                    End If
                                End If
                            End If
                            If ilSort = 0 Then
                                'Get Advertiser Name
'                                For ilLoop = 0 To UBound(tlAdvertiser) - 1 Step 1 'Traffic!lbcAdvt.ListCount - 1 Step 1
'                                    slNameCode = tlAdvertiser(ilLoop).sKey 'Traffic!lbcAdvt.List(ilLoop)
'                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                                    If Val(slCode) = tgRBofRec(ilUpper).tBof.iAdfCode Then
'                                        ilRet = gParseItem(slNameCode, 1, "\", tgRBofRec(ilUpper).sAdfName)
'                                        Exit For
'                                    End If
'                                Next ilLoop
                                tgRBofRec(ilUpper).sAdfName = "Advertiser Missing" & str(tgRBofRec(ilUpper).tBof.iAdfCode)
                                ilLoop = gBinarySearchAdf(tgRBofRec(ilUpper).tBof.iAdfCode)
                                If ilLoop <> -1 Then
                                    If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                                        tgRBofRec(ilUpper).sAdfName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                                    Else
                                        tgRBofRec(ilUpper).sAdfName = tgCommAdf(ilLoop).sName
                                    End If
                                End If
                                tgRBofRec(ilUpper).sKey = tgRBofRec(ilUpper).sAdfName & tgRBofRec(ilUpper).sShtTitle & slSDate
                            Else
                                tmCifSrchKey.lCode = tgRBofRec(ilUpper).tBof.lCifCode
                                If tgRBofRec(ilUpper).tBof.lCifCode > 0 Then
                                    ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                Else
                                    tmCif.iLen = 0
                                End If
                                slStr = Trim$(str$(tmCif.iLen))
                                Do While Len(slStr) < 3
                                    slStr = "0" & slStr
                                Loop
                                ilNum = Int(10000 * Rnd + 1)
                                slNum = Trim$(str$(ilNum))
                                Do While Len(slNum) < 5
                                    slNum = "0" & slNum
                                Loop
                                tgRBofRec(ilUpper).sKey = slStr & slNum
                                tgRBofRec(ilUpper).iLen = tmCif.iLen
                            End If
                            'Get Vehicle Name
                            tgRBofRec(ilUpper).sVefName = ""
'                            For ilLoop = 0 To UBound(tlUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
'                                slNameCode = tlUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
'                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                                If Val(slCode) = tgRBofRec(ilUpper).tBof.iVefCode Then
'                                    ilRet = gParseItem(slNameCode, 1, "\", slName)
'                                    ilRet = gParseItem(slName, 3, "|", tgRBofRec(ilUpper).sVefName)
'                                    Exit For
'                                End If
'                            Next ilLoop
                            ilLoop = gBinarySearchVef(tgRBofRec(ilUpper).tBof.iVefCode)
                            If ilLoop <> -1 Then
                                tgRBofRec(ilUpper).sVefName = tgMVef(ilLoop).sName
                            End If
                            tgRBofRec(ilUpper).iStatus = 1
                            tgRBofRec(ilUpper).lRecPos = llRecPos
                            ilUpper = ilUpper + 1
                            'ReDim Preserve tgRBofRec(1 To ilUpper) As BOFREC
                            ReDim Preserve tgRBofRec(0 To ilUpper) As BOFREC
                        End If
                    End If
                    ilRet = btrExtGetNext(hlBof, tmBof, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlBof, tmBof, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        'Test if records missing
        If ilPass = 1 Then
            ilUpper = UBound(tgSBofRec)
            If ilUpper > 1 Then
                ArraySortTyp fnAV(tgSBofRec(), 1), UBound(tgSBofRec) - 1, 0, LenB(tgSBofRec(1)), 0, LenB(tgSBofRec(1).sKey), 0
            End If
        Else
            ilUpper = UBound(tgRBofRec)
            If ilUpper > 1 Then
                ArraySortTyp fnAV(tgRBofRec(), 1), UBound(tgRBofRec) - 1, 0, LenB(tgRBofRec(1)), 0, LenB(tgRBofRec(1).sKey), 0
            End If
        End If
    Next ilPass
    'mInitBlackoutCtrls
    gReadBofRec = True
    Exit Function

    On Error GoTo 0
    gReadBofRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gTestForAirCopy                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test for Air copy and if exist *
'*                      add to RSF                     *
'*                                                     *
'*******************************************************
Public Sub gTestForAirCopy(ilType As Integer, slVefType As String, ilVefCode As Integer, ilVpfIndex As Integer, tlSdf As SDF, tlAvail As AVAILSS, hlCrf As Integer, hlCnf As Integer, hlCif As Integer, hlRsf As Integer, hlCvf As Integer, hlClf As Integer, slZone As String, ilZoneFd As Integer, ilCopyReplaced As Integer, ilRotNo As Integer)
    '10451 added hlClf for gObtainAirCopy
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilFound As Integer
    Dim slType As String
    Dim ilAZoneFd As Integer
    Dim ilACopyReplaced As Integer
    Dim ilARotNo As Integer
    'Copy rotation record information
    'Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim tlCrfSrchKey4 As CRFKEY4 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    Dim ilCrfCode As Long
    Dim blVefFound As Boolean

    ilCopyReplaced = False

    ilRotNo = -1
    If (slVefType <> "A") Or (tgVpf(ilVpfIndex).sCopyOnAir <> "Y") Then
        Exit Sub
    End If
    ilFound = False
    imRsfRecLen = Len(tmRsf)
    ilCrfRecLen = Len(tlCrf)
    '7/15/14
    tmRsfSrchKey4.lSdfCode = tlSdf.lCode
    tmRsfSrchKey4.sType = "A"
    tmRsfSrchKey4.iBVefCode = ilVefCode
    ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tlSdf.lCode) And (tmRsf.sType = "A")
        If (tmRsf.iBVefCode = ilVefCode) Then
            slType = "A"
            'tlCrfSrchKey1.sRotType = slType
            'tlCrfSrchKey1.iEtfCode = 0
            'tlCrfSrchKey1.iEnfCode = 0
            'tlCrfSrchKey1.iadfCode = tlSdf.iadfCode
            'tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
            'tlCrfSrchKey1.iVefCode = ilVefCode  'tmVef.iCode
            'tlCrfSrchKey1.iRotNo = tmRsf.iRotNo
            'ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            'If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iadfCode = tlSdf.iadfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode) And (tlCrf.iRotNo = tmRsf.iRotNo) Then    'tmVef.iCode)    'ilVefCode)
            tlCrfSrchKey4.sRotType = slType
            tlCrfSrchKey4.iEtfCode = 0
            tlCrfSrchKey4.iEnfCode = 0
            tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
            tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
            tlCrfSrchKey4.lFsfCode = 0
            tlCrfSrchKey4.iRotNo = tmRsf.iRotNo
            ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            blVefFound = False
            Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iRotNo = tmRsf.iRotNo)
                blVefFound = gCheckCrfVehicle(ilVefCode, tlCrf, hlCvf)
                If blVefFound Then
                    Exit Do
                End If
                ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (blVefFound) And (tlCrf.iRotNo = tmRsf.iRotNo) Then    'tmVef.iCode)    'ilVefCode)
                If ilType = 0 Then
                    If (Trim$(tlCrf.sZone) = "") Or (StrComp(Trim$(tlCrf.sZone), Trim$(slZone), vbTextCompare) = 0) Then
                        ilFound = True
                        ilZoneFd = True
                        'Any Superseding instructions
                        '10451
                        gObtainAirCopy 4, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                        If (ilARotNo > tlCrf.iRotNo) Then
                            If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo < ilARotNo)) Or (tlSdf.lCopyCode <= 0) Then
                                '10451
                                gObtainAirCopy 0, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                                If (ilACopyReplaced = True) Then
                                    'Update rotation
                                    Do
                                        tmRsf.lCrfCode = tmCrf.lCode
                                        tmRsf.iRotNo = ilARotNo
                                        tmRsf.sPtType = "1"
                                        tmRsf.lCopyCode = tlSdf.lCopyCode
                                        tmRsf.lSBofCode = CLng(ilZoneFd)
                                        gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
                                        tmRsf.sUnused = ""
                                        ilRet = btrUpdate(hlRsf, tmRsf, imRsfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            tmRsfSrchKey.lCode = tmRsf.lCode
                                            ilCRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    ilCopyReplaced = True
                                    ilRotNo = ilARotNo
                                    tlSdf.sPtType = "1"
                                    tlSdf.lCopyCode = tmRsf.lCopyCode
                                    ilZoneFd = CInt(tmRsf.lSBofCode)
                                End If
                                Exit Sub
                            End If
                        Else
                            If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo < tlCrf.iRotNo)) Or (tlSdf.lCopyCode <= 0) Then
                                ilCopyReplaced = True
                                ilRotNo = tlCrf.iRotNo
                                tlSdf.sPtType = "1"
                                tlSdf.lCopyCode = tmRsf.lCopyCode
                                ilZoneFd = CInt(tmRsf.lSBofCode)
                                Exit Sub
                            End If
                            '1/5/15: Same copy
                            If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo = tlCrf.iRotNo)) Then   ' And (tlSdf.lCopyCode = tmRsf.lCopyCode)) Then
                                ilCopyReplaced = True
                                ilRotNo = tlCrf.iRotNo
                                tlSdf.sPtType = "1"
                                tlSdf.lCopyCode = tmRsf.lCopyCode
                                ilZoneFd = CInt(tmRsf.lSBofCode)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    If StrComp(Trim$(tlCrf.sZone), Trim$(slZone), vbTextCompare) = 0 Then
                        ilFound = True
                        ilZoneFd = True
                        'Any Superseding instructions
                        '10451
                        gObtainAirCopy 6, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                        If (ilARotNo > tlCrf.iRotNo) Then
                            If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo < ilARotNo)) Or (tlSdf.lCopyCode <= 0) Then
                                '10451
                                gObtainAirCopy 2, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                                If (ilACopyReplaced = True) Then
                                    'Update rotation
                                    Do
                                        tmRsf.lCrfCode = tmCrf.lCode
                                        tmRsf.iRotNo = ilARotNo
                                        tmRsf.sPtType = "1"
                                        tmRsf.lCopyCode = tlSdf.lCopyCode
                                        tmRsf.lSBofCode = CLng(ilZoneFd)
                                        gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
                                        tmRsf.sUnused = ""
                                        ilRet = btrUpdate(hlRsf, tmRsf, imRsfRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            tmRsfSrchKey.lCode = tmRsf.lCode
                                            ilCRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    ilCopyReplaced = True
                                    ilRotNo = ilARotNo
                                    tlSdf.sPtType = "1"
                                    tlSdf.lCopyCode = tmRsf.lCopyCode
                                    ilZoneFd = CInt(tmRsf.lSBofCode)
                                End If
                                Exit Sub
                            End If
                        Else
                            If ((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo < tlCrf.iRotNo)) Or (tlSdf.lCopyCode <= 0) Then
                                ilCopyReplaced = True
                                ilRotNo = tlCrf.iRotNo
                                tlSdf.sPtType = "1"
                                tlSdf.lCopyCode = tmRsf.lCopyCode
                                ilZoneFd = CInt(tmRsf.lSBofCode)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hlRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    '10451
    gObtainAirCopy ilType, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
    If ilCopyReplaced Then
        tmRsf.lCode = 0
        tmRsf.lSdfCode = tlSdf.lCode
        tmRsf.sPtType = "1"
        tmRsf.lCopyCode = tlSdf.lCopyCode
        tmRsf.iRotNo = ilRotNo
        tmRsf.lRafCode = 0
        tmRsf.lSBofCode = CLng(ilZoneFd)
        tmRsf.lRBofCode = 0
        tmRsf.sType = "A"
        tmRsf.iBVefCode = ilVefCode
        tmRsf.lRChfCode = tlSdf.lChfCode
        tmRsf.lCrfCode = tmCrf.lCode
        gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
        tmRsf.sUnused = ""
        ilRet = btrInsert(hlRsf, tmRsf, imRsfRecLen, INDEXKEY0)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gTestForAirCopy                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test for Air copy and if exist *
'*                      add to RSF                     *
'*                                                     *
'*******************************************************
'4/30/11:  Add Region copy by Airing vehicle
Public Sub gTestForRegionAirCopy(slVefType As String, ilVefCode As Integer, ilVpfIndex As Integer, tlSdf As SDF, tlAvail As AVAILSS, hlCrf As Integer, hlCnf As Integer, hlCif As Integer, hlRsf As Integer, hlCvf As Integer, hlClf As Integer)
    '10451 added hlClf
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilFound As Integer
    Dim slType As String
    Dim ilAZoneFd As Integer
    Dim ilACopyReplaced As Integer
    Dim ilARotNo As Integer
    'Copy rotation record information
    'Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim tlCrfSrchKey4 As CRFKEY4 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    Dim slZone As String
    Dim ilZoneFd As Integer
    Dim ilCopyReplaced As Integer
    Dim ilRotNo As Integer
    Dim ilType As Integer
    Dim ilCrf As Long
    Dim blFirstTime As Boolean
    Dim tlSvSdf As SDF
    ReDim llCrfCode(0 To 0) As Long
    Dim blVefFound As Boolean
    
    tlSvSdf = tlSdf
    slZone = "R"
    ilCopyReplaced = False
    ilRotNo = -1
    ilType = 2
    If (slVefType <> "A") Or (tgVpf(ilVpfIndex).sCopyOnAir <> "Y") Then
        Exit Sub
    End If
    ilFound = False
    imRsfRecLen = Len(tmRsf)
    ilCrfRecLen = Len(tlCrf)
    '7/15/14
    tmRsfSrchKey4.lSdfCode = tlSdf.lCode
    tmRsfSrchKey4.sType = "R"
    tmRsfSrchKey4.iBVefCode = ilVefCode
    ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tlSdf.lCode) And (tmRsf.sType = "R")
        If (tmRsf.iBVefCode = ilVefCode) Then
            slType = "A"
            'tlCrfSrchKey1.sRotType = slType
            'tlCrfSrchKey1.iEtfCode = 0
            'tlCrfSrchKey1.iEnfCode = 0
            'tlCrfSrchKey1.iadfCode = tlSdf.iadfCode
            'tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
            'tlCrfSrchKey1.iVefCode = ilVefCode  'tmVef.iCode
            'tlCrfSrchKey1.iRotNo = tmRsf.iRotNo
            'ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            tlCrfSrchKey4.sRotType = slType
            tlCrfSrchKey4.iEtfCode = 0
            tlCrfSrchKey4.iEnfCode = 0
            tlCrfSrchKey4.iAdfCode = tlSdf.iAdfCode
            tlCrfSrchKey4.lChfCode = tlSdf.lChfCode
            tlCrfSrchKey4.lFsfCode = 0
            tlCrfSrchKey4.iRotNo = tmRsf.iRotNo
            ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            blVefFound = False
            Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iRotNo = tmRsf.iRotNo)
                blVefFound = gCheckCrfVehicle(ilVefCode, tlCrf, hlCvf)
                If blVefFound Then
                    Exit Do
                End If
                ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iadfCode = tlSdf.iadfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode) And (tlCrf.iRotNo = tmRsf.iRotNo) Then    'tmVef.iCode)    'ilVefCode)
            If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (blVefFound) And (tlCrf.iRotNo = tmRsf.iRotNo) Then    'tmVef.iCode)    'ilVefCode)
                If StrComp(Trim$(tlCrf.sZone), Trim$(slZone), vbTextCompare) = 0 Then
                    ilFound = True
                    ilZoneFd = True
                    'Any Superseding instructions
                    '10451
                    gObtainAirCopy 6, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                    If (ilARotNo > tlCrf.iRotNo) Then
                        If (tmRsf.lRafCode = tmCrf.lRafCode) And (((tlSdf.lCopyCode > 0) And (tlSdf.iRotNo < ilARotNo)) Or (tlSdf.lCopyCode <= 0)) Then
                            '10451
                            gObtainAirCopy 2, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
                            If (ilACopyReplaced = True) Then
                                '8/27/16: If added via several it will be re-added below because it was not marked as dormnant, to avoid that add to array of those processed
                                llCrfCode(UBound(llCrfCode)) = tlCrf.lCode
                                ReDim Preserve llCrfCode(0 To UBound(llCrfCode) + 1) As Long
                                'Update rotation
                                Do
                                    tmRsf.lCrfCode = tmCrf.lCode
                                    tmRsf.iRotNo = ilARotNo
                                    tmRsf.sPtType = "1"
                                    tmRsf.lCopyCode = tlSdf.lCopyCode
                                    tmRsf.lSBofCode = 0
                                    gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
                                    tmRsf.sUnused = ""
                                    ilRet = btrUpdate(hlRsf, tmRsf, imRsfRecLen)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        tmRsfSrchKey.lCode = tmRsf.lCode
                                        ilCRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                            End If
                            'Exit Sub
                        End If
                    Else
                        'If (tmRsf.lRafCode = tmCrf.lRafCode) Then
                        '    Exit Sub
                        'End If
                    End If
                    llCrfCode(UBound(llCrfCode)) = tmCrf.lCode
                    ReDim Preserve llCrfCode(0 To UBound(llCrfCode) + 1) As Long
                End If
            End If
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hlRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    blFirstTime = True
    Do
        tlSdf = tlSvSdf
        ilRet = mObtainCrfAirCopy(blFirstTime, 2, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, slZone)
        If Not ilRet Then
            Exit Do
        End If
        blFirstTime = False
        ilFound = False
        For ilCrf = 0 To UBound(llCrfCode) - 1 Step 1
            If tmCrf.lCode = llCrfCode(ilCrf) Then
                ilFound = True
            End If
        Next ilCrf
        '8/28/16: Check if superseded, if so don't add
        If Not ilFound Then
            tlCrf = tmCrf
            '10451
            gObtainAirCopy 6, slVefType, ilVefCode, ilVpfIndex, tlSdf, tlAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilAZoneFd, ilACopyReplaced, ilARotNo
            If (ilARotNo > tlCrf.iRotNo) Then
                tmCrf = tlCrf
                ilFound = True
            Else
                tmCrf = tlCrf
            End If
        End If
        If Not ilFound Then
            tmRsf.lCode = 0
            tmRsf.lSdfCode = tlSdf.lCode
            tmRsf.lCrfCode = tmCrf.lCode
            tmRsf.sPtType = "1"
            tmRsf.lCopyCode = tlSdf.lCopyCode
            tmRsf.iRotNo = tlSdf.iRotNo
            tmRsf.lRafCode = tmCrf.lRafCode
            tmRsf.lSBofCode = 0
            tmRsf.lRBofCode = 0
            tmRsf.sType = "R"
            tmRsf.iBVefCode = ilVefCode
            tmRsf.lRChfCode = tlSdf.lChfCode
            gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
            tmRsf.sUnused = ""
            ilRet = btrInsert(hlRsf, tmRsf, imRsfRecLen, INDEXKEY0)
        End If
    Loop
    tlSdf = tlSvSdf
End Sub
Public Function gSDFBlackoutTest(hlRsf As Integer, hlBof As Integer, llSdfCode As Long, ilVefCode As Integer, llChfCode As Long, llCifCode As Long, llSifCode As Long, llBofCode As Long) As Integer
    'hlRsf(I)- Regional Spot file handle
    'hlBof(I)- Blackout file handle
    'llSdfCode(I)- Spot Code
    'ilVefCode (I)- Airing or conventional vehicle code
    'llChfCode(O)- Contract Header Code
    'llCifCode(O)- Copy Inventory Code
    'llSifCode(O)- Short Title code
    Dim ilRet As Integer

    llChfCode = 0
    llCifCode = 0
    llSifCode = 0
    llBofCode = 0
    gSDFBlackoutTest = False
    imRsfRecLen = Len(tmRsf)
    imBofRecLen = Len(tmBof)
    '7/15/14
    tmRsfSrchKey4.lSdfCode = llSdfCode
    tmRsfSrchKey4.sType = "B"
    tmRsfSrchKey4.iBVefCode = ilVefCode
    ilRet = btrGetEqual(hlRsf, tmRsf, imRsfRecLen, tmRsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = llSdfCode) And (tmRsf.sType = "B")
        If (tmRsf.iBVefCode = ilVefCode) Then
            llChfCode = tmRsf.lRChfCode
            If tmRsf.lRBofCode <> tmRsf.lSBofCode Then
                tmBofSrchKey0.lCode = tmRsf.lRBofCode
            Else
                tmBofSrchKey0.lCode = tmRsf.lSBofCode
            End If
            gSDFBlackoutTest = True
            ilRet = btrGetEqual(hlBof, tmBof, imBofRecLen, tmBofSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                llBofCode = tmBof.lCode
                llCifCode = tmBof.lCifCode
                If tgSpf.sUseProdSptScr = "P" Then
                    llSifCode = tmBof.lSifCode
                End If
            End If
            Exit Do
        Else
            Exit Do
        End If
        ilRet = btrGetNext(hlRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Function
End Function

Public Function gCheckCrfVehicle(ilVefCode, tlCrf As CRF, hlCvf As Integer) As Boolean
    Dim ilRet As Integer
    Dim ilCvf As Integer
    Dim blVefFound As Boolean
    
    blVefFound = False
    igAsgnVehRowNo = -1
    If ilVefCode <= 0 Then
        gCheckCrfVehicle = blVefFound
        Exit Function
    End If
    If tlCrf.iVefCode > 0 Then
        If tlCrf.iVefCode = ilVefCode Then
            blVefFound = True
        End If
    Else
        'If tmCvf.lCode <> tlCrf.lCvfCode Then
            imCvfRecLen = Len(tmCvf)
            tmCvfSrchKey.lCode = tlCrf.lCvfCode
            ilRet = btrGetEqual(hlCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'Else
        '    ilRet = BTRV_ERR_NONE
        'End If
        Do While ilRet = BTRV_ERR_NONE
            For ilCvf = 0 To 99 Step 1
                If tmCvf.iVefCode(ilCvf) > 0 Then
                    If tmCvf.iVefCode(ilCvf) = ilVefCode Then
                        blVefFound = True
                        igAsgnVehRowNo = ilCvf
                        tgAsgnVehCvf = tmCvf
                        Exit Do
                    End If
                End If
            Next ilCvf
            If tmCvf.lLkCvfCode <= 0 Then
                Exit Do
            End If
            tmCvfSrchKey.lCode = tmCvf.lLkCvfCode
            ilRet = btrGetEqual(hlCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    gCheckCrfVehicle = blVefFound
End Function

Private Sub mReplaceSpot(ilSource As Integer, tlBofRec As BOFREC, ilVefCode As Integer, ilLnVefCode As Integer, ilVpfIndex As Integer, ilLnVpfIndex As Integer, llDate As Long, slDate As String, llTime As Long, slTime As String, ilSeqNo As Integer, ilLen As Integer, hlODF As Integer, hlMcf As Integer, hlClf As Integer, hlCrf As Integer, hlCnf As Integer, hlCif As Integer, hlCvf As Integer, hlCpf As Integer, hlLst As Integer, hlRsf As Integer, slZone As String, ilZoneFd As Integer, ilCopyReplaced As Integer, ilRotNo As Integer, ilTestRotNo As Integer, ilRotNoVefCode As Integer, ilCrfVefCode As Integer, ilPkgVefCode As Integer, ilCLnVefCode As Integer, slLive As String, ilRdfCode As Integer, slVehName As String, llLoop As Long, ilSBof As Integer, slLogType As String, slNewLines() As String * 72, hlMsg As Integer, lbcMsg As control, ilRsfExist As Integer, ilAlert As Integer, ilFound As Integer)
    Dim lSvCifCode As Long
    Dim llChfCode As Long
    Dim ilRet As Integer
    Dim llIndex As Long
    Dim slCart As String
    Dim ilVef As Integer
    Dim slMsg As String
    
    lSvCifCode = tlBofRec.tBof.lCifCode
    llChfCode = tlBofRec.tBof.lRChfCode
    If tlBofRec.tBof.lCifCode <= 0 Then
        'Assign copy
        'Make-up spot and tmAvailTest to obtain Copy
        tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
        gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
        'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
        gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
        tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
        tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
        ilRet = btrGetEqual(hlODF, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            tmSdf.lChfCode = tlBofRec.tBof.lRChfCode
            tmSdf.iAdfCode = tlBofRec.tBof.iRAdfCode
            tmSdf.iDate(0) = tmOdf.iAirDate(0)
            tmSdf.iDate(1) = tmOdf.iAirDate(1)
            tmSdf.iTime(0) = tmOdf.iAirTime(0)
            tmSdf.iTime(1) = tmOdf.iAirTime(1)
            tmSdf.iLen = ilLen
            tmSdf.sSchStatus = "S"
            tmSdf.iRotNo = 0
            tmSdf.sPtType = ""
            tmSdf.lCopyCode = 0
            tmAvail.ianfCode = tmOdf.ianfCode
            '10451
            gObtainAirCopy 3, "", ilVefCode, ilVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
            ilRotNoVefCode = ilVefCode
            'tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
            If (ilVefCode <> ilLnVefCode) Then  'Try Selling vehicle
                ilLnVpfIndex = 0
                'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If ilLnVefCode = tgVpf(llIndex).iVefKCode Then
                    llIndex = gBinarySearchVpf(ilLnVefCode)
                    If llIndex <> -1 Then
                        ilLnVpfIndex = llIndex
                '        Exit For
                    End If
                'Next llIndex
                '10451
                gObtainAirCopy 3, "", ilLnVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
                If ilTestRotNo > ilRotNo Then
                    ilRotNoVefCode = ilLnVefCode
                    ilRotNo = ilTestRotNo
                End If
                'tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
            End If
            ilRet = gGetCrfVefCode(hlClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilCLnVefCode, slLive, ilRdfCode)
            If ilPkgVefCode <> 0 Then
                ilLnVpfIndex = 0
                'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If ilPkgVefCode = tgVpf(llIndex).iVefKCode Then
                    llIndex = gBinarySearchVpf(ilPkgVefCode)
                    If llIndex <> -1 Then
                        ilLnVpfIndex = llIndex
                '        Exit For
                    End If
                'Next llIndex
                '10451
                gObtainAirCopy 3, "", ilPkgVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
                If ilTestRotNo > ilRotNo Then
                    ilRotNoVefCode = ilPkgVefCode
                End If
            End If
            'For llIndex = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
            '    If ilRotNoVefCode = tgVpf(llIndex).iVefKCode Then
                llIndex = gBinarySearchVpf(ilRotNoVefCode)
                If llIndex <> -1 Then
                    ilLnVpfIndex = llIndex
            '        Exit For
                End If
            'Next llIndex
            '10451
            gObtainAirCopy 1, "", ilRotNoVefCode, ilLnVpfIndex, tmSdf, tmAvail, hlCrf, hlCnf, hlCif, hlCvf, hlClf, slZone, ilZoneFd, ilCopyReplaced, ilTestRotNo
            tlBofRec.tBof.lCifCode = tmSdf.lCopyCode
        End If
    End If
    'Replace spot (Cart ID and Short Title)
    tmCifSrchKey.lCode = tlBofRec.tBof.lCifCode
    ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        If tgSpf.sUseCartNo <> "N" Then
            If tmCif.iMcfCode > 0 Then
                If tmMcf.iCode <> tmCif.iMcfCode Then
                    tmMcfSrchKey.iCode = tmCif.iMcfCode
                    ilRet = btrGetEqual(hlMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmMcf.sName = "C"
                        tmMcf.sPrefix = "C"
                    End If
                End If
            Else
                tmMcf.sName = ""
                tmMcf.sPrefix = ""
            End If
            If Trim$(tmCif.sCut) = "" Then
                slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
            Else
                slCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
            End If
        Else
            slCart = ""
            tmCif.lcpfCode = 0
            ilRet = BTRV_ERR_NONE
        End If
    ElseIf (tlBofRec.tBof.lCifCode = 0) And (ilSource = 1) Then
        slCart = ""
        tmCif.lcpfCode = 0
        ilRet = BTRV_ERR_NONE
    End If
    If (ilRet = BTRV_ERR_NONE) Then
        If ilSource = 0 Then
            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 24, Len(slCart)) = slCart
            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 30, 15) = UCase$(tlBofRec.sShtTitle)
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hlCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Then
                tmCpf.sISCI = ""
            End If
            Mid$(slNewLines(tgSpotSum(llLoop).iNewIndex), 53, 20) = UCase$(tmCpf.sISCI)
        Else
            'Update ODF and LST
            Do
                'tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
                'gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
                'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
                'tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
                'tmOdfSrchKey0.iSeqNo = tgSpotSum(llLoop).iSeqNo
                tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
                gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
                'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
                gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
                tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
                tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
                ilRet = btrGetEqual(hlODF, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) Then
                    tmOdf.iAdfCode = tlBofRec.tBof.iRAdfCode
                    tmOdf.lCifCode = tlBofRec.tBof.lCifCode
                    If tgSpf.sUseProdSptScr = "P" Then
                        tmOdf.sProduct = ""
                        tmOdf.sShortTitle = tlBofRec.sShtTitle
                        If (Trim$(tmOdf.sShortTitle) = "") Then
                            slTime = gFormatTimeLong(llTime, "A", "1")
                            slVehName = ""
                            'For ilVef = 0 To UBound(tgVehicle) - 1 Step 1
                            '    slNameCode = tgVehicle(ilVef).sKey 'lbcVehCode.List(llLoop)
                            '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            '    If Val(slCode) = ilVefCode Then
                            '        ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                            '        ilRet = gParseItem(slVehName, 3, "|", slVehName)
                            '        Exit For
                            '    End If
                            'Next ilVef
                            ilVef = gBinarySearchVef(ilVefCode)
                            If ilVef <> -1 Then
                                slVehName = tgMVef(ilVef).sName
                            Else
                                slVehName = "Vehicle" & str(ilVefCode) & " Code missing"
                            End If
                            slMsg = "Replacement Found but Short Title Missing: " & Trim$(slVehName) & " " & Format$(llDate, "m/d/yy") & " at " & slTime
                            Print #hlMsg, slMsg
                            lbcMsg.AddItem slMsg
                        End If
                    Else
                        tmOdf.sProduct = tlBofRec.sShtTitle
                        tmOdf.sShortTitle = ""
                        If (tgSpotSum(llLoop).lLstCode > 0) Then
                            tmOdf.sShortTitle = slCart
                        End If
                    End If
                    tmOdf.lCntrNo = tlBofRec.lRCntrNo
                    tmOdf.sBBDesc = ""
                    ilRet = btrUpdate(hlODF, tmOdf, imOdfRecLen)
                Else
                    Exit Do
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If Not ilRsfExist Then
                tmRsf.lCode = 0
                tmRsf.lSdfCode = tgSpotSum(llLoop).lSdfCode
                If tlBofRec.tBof.lCifCode > 0 Then
                    tmRsf.sPtType = "1"
                    tmRsf.lCopyCode = tlBofRec.tBof.lCifCode
                Else
                    tmRsf.sPtType = "0"
                    tmRsf.lCopyCode = 0
                    tmRsf.iRotNo = 0
                End If
                tmRsf.lRafCode = 0
                tmRsf.lSBofCode = tgSBofRec(ilSBof).tBof.lCode
                tmRsf.lRBofCode = tlBofRec.tBof.lCode
                tmRsf.sType = "B"
                tmRsf.iBVefCode = ilVefCode
                tmRsf.lRChfCode = llChfCode
                gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
                tmRsf.sUnused = ""
                ilRet = btrInsert(hlRsf, tmRsf, imRsfRecLen, INDEXKEY0)
                ilRsfExist = True
            End If
            If (tgSpotSum(llLoop).lLstCode > 0) And (ilRet = BTRV_ERR_NONE) Then
                'Do
                    tmLstSrchKey.lCode = tgSpotSum(llLoop).lLstCode
                    ilRet = btrGetEqual(hlLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) Then
                        ilRet = btrDelete(hlLst)
                        'If tmLst.lSdfCode > 0 Then
                        '    tmLst.lSdfCode = -tmLst.lSdfCode
                        'End If
                        tmLst.lCntrNo = tlBofRec.lRCntrNo
                        tmLst.iAdfCode = tlBofRec.tBof.iRAdfCode
                        tmLst.iAgfCode = 0
                        If tgSpf.sUseProdSptScr = "P" Then
                            tmLst.sProd = ""
                        Else
                            tmLst.sProd = tlBofRec.sShtTitle
                        End If
                        tmLst.iLineNo = 0
                        tmLst.iLnVefCode = 0
                        tmLst.iStartDate(0) = 0
                        tmLst.iStartDate(1) = 0
                        tmLst.iEndDate(0) = 0
                        tmLst.iEndDate(1) = 0
                        tmLst.iDays(0) = 0
                        tmLst.iDays(0) = 0
                        tmLst.iDays(1) = 0
                        tmLst.iDays(2) = 0
                        tmLst.iDays(3) = 0
                        tmLst.iDays(4) = 0
                        tmLst.iDays(5) = 0
                        tmLst.iDays(6) = 0
                        tmLst.iSpotsWk = 0
                        tmLst.iPriceType = 1
                        tmLst.lPrice = 0
                        tmLst.iSpotType = 4
                        tmLst.sDemo = ""
                        tmLst.lAud = 0
                        tmLst.sISCI = ""
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hlCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                            If (ilRet = BTRV_ERR_NONE) Then
                                tmLst.sISCI = tmCpf.sISCI
                            Else
                                tmLst.sISCI = ""
                            End If
                        Else
                            tmLst.sISCI = ""
                        End If
                        tmLst.sCart = slCart
                        tmLst.lcpfCode = tmCif.lcpfCode
                        tmLst.lCrfCsfcode = 0
                        tmLst.lCifCode = tlBofRec.tBof.lCifCode
                        tmLst.sImportedSpot = "N"
                        tmLst.lBkoutLstCode = 0
                        'ilRet = btrUpdate(hlLst, tmLst, imLstRecLen)
                        ilRet = btrInsert(hlLst, tmLst, imLstRecLen, INDEXKEY0)
                        'Checking if date bewteen todays date and last log is not required
                        'It is Ok if Alert set on when generating Final Logs (The Alert will be removed within mGenLog)
                        slDate = Format$(llDate, "m/d/yy")
                        ilAlert = gAlertAdd(slLogType, "S", 0, tmLst.iLogVefCode, slDate)
                        ilAlert = gAlertAdd(slLogType, "I", 0, tmLst.iLogVefCode, slDate)
                    Else
                        'Exit Do
                    End If
                'Loop While ilRet = BTRV_ERR_CONFLICT
            End If
        End If
        ilFound = True
    End If
    tlBofRec.tBof.lCifCode = lSvCifCode
End Sub

Private Sub mReplaceMissing(ilSource As Integer, ilVefCode As Integer, ilVpfIndex As Integer, llDate As Long, llTime As Long, slZone As String, ilSeqNo As Integer, hlODF As Integer)
    Dim ilRet As Integer
    If (ilSource = 1) Then
        tmOdfSrchKey0.iVefCode = ilVefCode  'tgSpotSum(llLoop).iVefCode
        gPackDateLong llDate, tmOdfSrchKey0.iAirDate(0), tmOdfSrchKey0.iAirDate(1)
        'gPackTimeLong tgSpotSum(llLoop).lTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
        gPackTimeLong llTime, tmOdfSrchKey0.iLocalTime(0), tmOdfSrchKey0.iLocalTime(1)
        tmOdfSrchKey0.sZone = slZone    'tgSpotSum(llLoop).sZone   'tmDlf.sZone
        tmOdfSrchKey0.iSeqNo = ilSeqNo  'tgSpotSum(llLoop).iSeqNo
        ilRet = btrGetEqual(hlODF, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            If tgVpf(ilVpfIndex).sUnsoldBlank = "N" Then
                ilRet = btrDelete(hlODF)
            Else
                tmOdf.iAdfCode = 0
                tmOdf.lCifCode = 0
                If tgSpf.sUseProdSptScr = "P" Then
                    tmOdf.sProduct = ""
                    tmOdf.sShortTitle = ""  '"Replace Missing"
                Else
                    tmOdf.sProduct = "" '"Replace Missing"
                    tmOdf.sShortTitle = ""
                End If
                tmOdf.lCntrNo = 0
                tmOdf.sBBDesc = ""
                ilRet = btrUpdate(hlODF, tmOdf, imOdfRecLen)
            End If
        End If
    End If
End Sub
Function gGetClfLive(hlClf As Integer, tlSdf As SDF) As String
    '10451  Returns 'slLive'
    
    Dim ilRet As Integer
    Dim tlClfSrchKey As CLFKEY0 'CLF key record image
    Dim ilClfRecLen As Integer  'CLF record length
    Dim tlClf As CLF            'CLF record image
    Dim slLive As String
    
    slLive = "R"
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
        If (tlClf.sLiveCopy = "L") Or (tlClf.sLiveCopy = "M") Or (tlClf.sLiveCopy = "S") Or (tlClf.sLiveCopy = "P") Or (tlClf.sLiveCopy = "Q") Then
            slLive = tlClf.sLiveCopy
        End If
    End If
    gGetClfLive = slLive
End Function

