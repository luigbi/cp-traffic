Attribute VB_Name = "COPYASGNSUB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyasgn.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSsfSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule function
'
'
Option Explicit
Option Compare Text

Public lgAssignTotal As Long
Public lgAssignStart As Long
Public lgAssignEnd As Long

'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
'Dim tmSsfOld As SSF
Dim tmSsfSrchKey2 As SSFKEY2
Dim imSsfRecLen As Integer
Dim tmAvail As AVAILSS
'Spot record
Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey As SDFKEY1
Dim tmSdfSrchKey6 As SDFKEY6
'Contract record information
Dim hmClf As Integer        'Contract line file handle
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Copy rotation record information
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrfSrchKey As LONGKEY0 'CIF key record image
Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim tmCrfSrchKey4 As CRFKEY4 'CRF key record image
Dim imCrfRecLen As Integer  'CRF record length
Dim tmCrf As CRF            'CRF record image
'Copy instruction record information
Dim hmCnf As Integer        'Copy instruction file handle
Dim tmCnfSrchKey As CNFKEY0 'CNF key record image
Dim imCnfRecLen As Integer  'CNF record length
Dim tmCnf As CNF            'CNF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
'Time zone Copy inventory record information
Dim hmTzf As Integer        'Time zone Copy inventory file handle
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer  'TZF record length
Dim tmTzf As TZF            'TZF record image
'Regional Scheduled Copy record information
Dim hmRsf As Integer        'Regional Scheduled copy file handle
Dim tmRsfSrchKey1 As LONGKEY0 'RSF key record image
Dim tmRsfSrchKey3 As RSFKEY3
Dim tmRsfSrchKey4 As RSFKEY4
Dim imRsfRecLen As Integer  'RSF record length
Dim tmRsf As RSF            'RSF record image
'Copy Air Game
Dim tmCaf As CAF            'CAF record image
Dim tmCafSrchKey As LONGKEY0  'CAF key record image
Dim tmCafSrchKey1 As CAFKEY1  'CAF key record image
Dim hmCaf As Integer        'CAF Handle
Dim imCafRecLen As Integer      'CAF record length
'Copy Vehicles
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0  'CVF key record image
Dim tmCvfSrchKey1 As LONGKEY0  'CVF key record image
Dim hmCvf As Integer        'CVF Handle
Dim imCvfRecLen As Integer      'CVF record length
'Game schedule
Dim hmGsf As Integer
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length
Dim tmGsf As GSF

Dim smNowDate As String
Dim smNowTime As String
Public tgCAVehicle() As SORTCODE
Public sgCAVehicleTag As String
'*******************************************************
'*                                                     *
'*      Procedure Name:gAssignCopy                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Assign copy                    *
'*                                                     *
'*******************************************************
Function gAssignCopyToSpots(ilType As Integer, ilVefCode As Integer, ilPFAsgn As Integer, slStartDate As String, slEndDate As String, slStartTime As String, slEndTime As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   ilRet = gAssignCopyToSpots(ilVefCode, ilPFAsgn, slStartDate, slEndDate, slStartTime, slEndTime)
'
'   Where:
'       ilType(I) 0=Non Game or Game; -1=Billboards
'       ilVefCode (I)-Vehicle code number
'       ilPFAsgn(I) - 0=Prelimary; 1=Final
'       slZone(I)- Zone name (EST, CST<..) Or Blank for all
'       slStartDate (I)- Start Date that that spots are to be removed
'       slEndDate (I)- End Date that that spots are to be removed
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llSpotTime As Long
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilAsgnDate0 As Integer
    Dim ilAsgnDate1 As Integer
    Dim ilAsgnTime0 As Integer
    Dim ilAsgnTime1 As Integer
    Dim slType As String
    'Dim llCrfRecPos As Long
    Dim llCrfCode As Long
    Dim llSdfRecPos As Long
    Dim ilDay As Integer
    Dim ilSpotAsgn As Integer
    Dim ilAllZones As Integer    'All zones assigned
    Dim ilDayDone As Integer
    Dim llAvailTime As Long
    Dim ilAvailOk As Integer
    Dim ilEvtIndex As Integer
    Dim ilCrfVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim ilBypassCrf As Integer
    Dim ilAssgToSpot As Integer
    Dim ilOrigCrfVefCode As Integer
    Dim hlSsf As Integer        'Spot summary file handle
'    Dim hlRdf As Integer
    Dim hlSmf As Integer
    Dim slStr As String
    Dim slVehType As String
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim blVefFound As Boolean
    Dim blCrfFound As Boolean
    Dim ilCvf As Integer
    Dim ilCrf As Integer
    Dim ilRotNo As Integer
    Dim ilVefPass As Integer
    Dim tlCrf As CRF
    Dim blAnyAssigned As Boolean
    ReDim llProcessedCrfCode(0 To 0) As Long

    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    If llEndDate < llStartDate Then
        gAssignCopyToSpots = True
        Exit Function
    End If
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    gPackTime slStartTime, ilAsgnTime0, ilAsgnTime1
    smNowDate = Format$(gNow(), "m/d/yy")
    smNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    hmSdf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
    hmCrf = CBtrvTable(TWOHANDLES)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imCrfRecLen = Len(tmCrf)
    hmCnf = CBtrvTable(ONEHANDLE)        'Create CNF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imCnfRecLen = Len(tmCnf)
    hmCif = CBtrvTable(TWOHANDLES)        'Create CNF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        gAssignCopyToSpots = False
        Exit Function
    End If
    imCifRecLen = Len(tmCif)
    hmTzf = CBtrvTable(TWOHANDLES)        'Create CNF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imTzfRecLen = Len(tmTzf)
    hmClf = CBtrvTable(ONEHANDLE)        'Create CNF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imClfRecLen = Len(tmClf)
    hlSsf = CBtrvTable(ONEHANDLE)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        gAssignCopyToSpots = False
        Exit Function
    End If
    hmRsf = CBtrvTable(TWOHANDLES)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hmRsf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        btrDestroy hmRsf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imRsfRecLen = Len(tmRsf)
    hlSmf = CBtrvTable(ONEHANDLE)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hlSmf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        btrDestroy hmRsf
        btrDestroy hlSmf
        gAssignCopyToSpots = False
        Exit Function
    End If
    hmCaf = CBtrvTable(ONEHANDLE)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCaf, "", sgDBPath & "Caf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hmCaf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        btrDestroy hmRsf
        btrDestroy hlSmf
        btrDestroy hmCaf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imCafRecLen = Len(tmCaf)
    hmGsf = CBtrvTable(ONEHANDLE)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hmCaf)
        ilRet = btrClose(hmGsf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        btrDestroy hmRsf
        btrDestroy hlSmf
        btrDestroy hmCaf
        btrDestroy hmGsf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imGsfRecLen = Len(tmGsf)
    
    hmCvf = CBtrvTable(ONEHANDLE)        'Create CRF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmCnf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hmCaf)
        ilRet = btrClose(hmGsf)
        ilRet = btrClose(hmCvf)
        btrDestroy hmSdf
        btrDestroy hmCrf
        btrDestroy hmCnf
        btrDestroy hmCif
        btrDestroy hmTzf
        btrDestroy hmClf
        btrDestroy hlSsf
        btrDestroy hmRsf
        btrDestroy hlSmf
        btrDestroy hmCaf
        btrDestroy hmGsf
        btrDestroy hmCvf
        gAssignCopyToSpots = False
        Exit Function
    End If
    imCvfRecLen = Len(tmCvf)

'    hlRdf = CBtrvTable(ONEHANDLE)        'Create CRF object handle
'    On Error GoTo 0
'    ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSdf)
'        ilRet = btrClose(hmCrf)
'        ilRet = btrClose(hmCnf)
'        ilRet = btrClose(hmCif)
'        ilRet = btrClose(hmTzf)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hlSsf)
'        ilRet = btrClose(hmRsf)
'        ilRet = btrClose(hlSmf)
'        ilRet = btrClose(hlRdf)
'        btrDestroy hmSdf
'        btrDestroy hmCrf
'        btrDestroy hmCnf
'        btrDestroy hmCif
'        btrDestroy hmTzf
'        btrDestroy hmClf
'        btrDestroy hlSsf
'        btrDestroy hmRsf
'        btrDestroy hlSmf
'        btrDestroy hlRdf
'        gAssignCopyToSpots = False
'        Exit Function
'    End If
    'ilRet = btrBeginTrans(hlSdf, 1000)
    lgAssignStart = timeGetTime
    
    slVehType = ""
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        slVehType = tgMVef(ilVef).sType
    End If
    For llDate = llStartDate To llEndDate Step 1
        slDate = Format$(llDate, "m/d/yy")
        gPackDate slDate, ilAsgnDate0, ilAsgnDate1
        ilDay = gWeekDayStr(slDate)
        ilDayDone = False
        If ilType <> -1 Then
            'tmSsfSrchKey.iType = ilType
            'tmSsfSrchKey.iVefCode = ilVefCode
            'tmSsfSrchKey.iDate(0) = ilAsgnDate0
            'tmSsfSrchKey.iDate(1) = ilAsgnDate1
            'tmSsfSrchKey.iStartTime(0) = 0
            'tmSsfSrchKey.iStartTime(1) = 0
            imSsfRecLen = Len(tmSsf)
            tmSsfSrchKey2.iVefCode = ilVefCode
            tmSsfSrchKey2.iDate(0) = ilAsgnDate0
            tmSsfSrchKey2.iDate(1) = ilAsgnDate1
            ilRet = gSSFGetEqualKey2(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        End If
        'If ((ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1)) Or (ilType = -1) Then
        Do While ((ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1)) Or (ilType = -1)
            ilEvtIndex = 1
            If (ilType = -1) Or (tmSsf.iType = 0) Then
                tmSdfSrchKey.iVefCode = ilVefCode
                tmSdfSrchKey.iDate(0) = ilAsgnDate0
                tmSdfSrchKey.iDate(1) = ilAsgnDate1
                tmSdfSrchKey.iTime(0) = ilAsgnTime0
                tmSdfSrchKey.iTime(1) = ilAsgnTime1
                tmSdfSrchKey.sSchStatus = ""
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Else
                tmGsfSrchKey3.iVefCode = ilVefCode
                tmGsfSrchKey3.iGameNo = tmSsf.iType
                ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = tmSsf.iType)
                    If (tmGsf.iAirDate(0) = tmSsf.iDate(0)) And (tmGsf.iAirDate(1) = tmSsf.iDate(1)) Then
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = tmSsf.iType) Then
                    'tmSdfSrchKey6.iVefCode = ilVefCode
                    'tmSdfSrchKey6.iGameNo = tmSsf.iType
                    'ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey6, INDEXKEY6, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iGameNo = tmSsf.iType)
                    '    If (tmSdf.iDate(0) = ilAsgnDate0) And (tmSdf.iDate(1) = ilAsgnDate1) Then
                    '        Exit Do
                    '    End If
                    '    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    'Loop
                    tmSdfSrchKey.iVefCode = ilVefCode
                    tmSdfSrchKey.iDate(0) = ilAsgnDate0
                    tmSdfSrchKey.iDate(1) = ilAsgnDate1
                    tmSdfSrchKey.iTime(0) = ilAsgnTime0
                    tmSdfSrchKey.iTime(1) = ilAsgnTime1
                    tmSdfSrchKey.sSchStatus = ""
                    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode)
                        If (tmSdf.iGameNo = tmSsf.iType) Then
                            Exit Do
                        End If
                        If (tmSdf.iDate(0) <> ilAsgnDate0) Or (tmSdf.iDate(1) <> ilAsgnDate1) Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    Loop
                End If
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iDate(0) = ilAsgnDate0) And (tmSdf.iDate(1) = ilAsgnDate1) And ((ilType = -1) Or (tmSdf.iGameNo = tmSsf.iType))
                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                llSpotTime = CLng(gTimeToCurrency(slTime, False)) '- 1
                If llSpotTime > llEndTime Then
                    Exit Do
                End If
                If (ilType = -1) And (slVehType = "G") And (tmSdf.iGameNo <= 0) Then
                    Exit Do
                End If
                ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                ilAssgToSpot = False
                If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (tmSdf.iRotNo <> -1) Then    'Add spot
                    ilAssgToSpot = True
                Else
                    'Remove assign copy to missed- this will be done in invoicing- it was required
                    'for as order billing
                    'If (tgSpf.sInvAirOrder = "O") And ((tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "U")) Then
                    '    ilAssgToSpot = True
                    'End If
                End If
                If ilType <> -1 Then
                    If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                        ilAssgToSpot = False
                    End If
                Else
                    If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        ilAssgToSpot = False
                    End If
                End If
                If ilAssgToSpot Then
                    ilSchPkgVefCode = 0
                    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSpotType = "X") Then
'                        slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode, hlSmf, hmCrf, hlRdf)
                        slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, hlSmf, hmCrf)
                        If (slStr = "S") Or (slStr = "B") Then
                            ilSchPkgVefCode = gGetMGPkgVefCode(hmClf, tmSdf)
                        End If
                        If slStr = "O" Then
                            ilCrfVefCode = ilLnVefCode
                            ilLnVefCode = 0
                        ElseIf slStr = "S" Then
                            ilPkgVefCode = ilSchPkgVefCode
                            ilSchPkgVefCode = 0
                            ilLnVefCode = 0
                        Else
                            If ilPkgVefCode = ilSchPkgVefCode Then
                                ilSchPkgVefCode = 0
                            End If
                        End If
                    Else
                        ilLnVefCode = 0
                    End If
                    ilOrigCrfVefCode = ilCrfVefCode
                    '12/31/14: Move multi-vehicle pass to be within the rotation loop
                    'Do
                        ReDim llProcessedCrfCode(0 To 0) As Long
                        'Find rotation to assign
                        'Code later- test spot type to determine which rotation type
                        ilSpotAsgn = False
                        ilAllZones = False
                        ilDayDone = False
                        ilAvailOk = True
                        If tmSdf.sSpotType = "O" Then
                            slType = "O"
                        ElseIf tmSdf.sSpotType = "C" Then
                            slType = "C"
                        Else
                            slType = "A"
                        End If
                        'tmCrfSrchKey1.sRotType = slType
                        'tmCrfSrchKey1.iEtfCode = 0
                        'tmCrfSrchKey1.iEnfCode = 0
                        'tmCrfSrchKey1.iadfCode = tmSdf.iadfCode
                        'tmCrfSrchKey1.lChfCode = tmSdf.lChfCode
                        'tmCrfSrchKey1.lFsfCode = 0
                        'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'ilVefCode
                        'tmCrfSrchKey1.iRotNo = 32000
                        'ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.iVefCode = ilCrfVefCode)    'ilVefCode)
                        tmCrfSrchKey4.sRotType = slType
                        tmCrfSrchKey4.iEtfCode = 0
                        tmCrfSrchKey4.iEnfCode = 0
                        tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
                        tmCrfSrchKey4.lChfCode = tmSdf.lChfCode
                        tmCrfSrchKey4.lFsfCode = 0
                        tmCrfSrchKey4.iRotNo = 32000
                        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
                        Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode)
                            'ilRet = btrGetPosition(hmCrf, llCrfRecPos)
                            llCrfCode = tmCrf.lCode
                            ilRotNo = tmCrf.iRotNo
                            llProcessedCrfCode(UBound(llProcessedCrfCode)) = llCrfCode
                            ReDim Preserve llProcessedCrfCode(0 To UBound(llProcessedCrfCode) + 1) As Long
                            'blVefFound = False
                            'If tmCrf.iVefCode > 0 Then
                            '    If tmCrf.iVefCode = ilCrfVefCode Then
                            '        blVefFound = True
                            '    End If
                            'Else
                            '    tmCvfSrchKey.lCode = tmCrf.lCvfCode
                            '    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            '    Do While ilRet = BTRV_ERR_NONE
                            '        For ilCvf = 0 To 99 Step 1
                            '            If tmCvf.iVefCode(ilCvf) > 0 Then
                            '                If tmCvf.iVefCode(ilCvf) = ilCrfVefCode Then
                            '                    blVefFound = True
                            '                    Exit Do
                            '                End If
                            '            End If
                            '        Next ilCvf
                            '        If tmCvf.lLkCvfCode <= 0 Then
                            '            Exit Do
                            '        End If
                            '        tmCvfSrchKey.lCode = tmCvf.lLkCvfCode
                            '        ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            '    Loop
                            'End If
                            '12/31/14: Move vehicle pass here
                            blAnyAssigned = False
                            If tmCrf.sState <> "D" Then
                                tlCrf = tmCrf
                                For ilVefPass = 0 To 3 Step 1
                                    Select Case ilVefPass
                                        Case 0
                                            ilCrfVefCode = ilOrigCrfVefCode
                                        Case 1
                                            ilCrfVefCode = ilPkgVefCode
                                        Case 2
                                            ilCrfVefCode = ilSchPkgVefCode
                                        Case 3
                                            If (ilOrigCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                                                ilCrfVefCode = 0
                                            Else
                                                ilCrfVefCode = ilLnVefCode
                                            End If
                                    End Select
                                    
                                    blVefFound = gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf)
                                    If blVefFound Then
                                        'Test date, time, day and zone
                                        ilSpotAsgn = False
                                        ilBypassCrf = False
                                        'Test if looking for Live or Recorded rotations
                                        If slLive = "L" Then
                                            If tmCrf.sLiveCopy <> "L" Then
                                                ilBypassCrf = True
                                            End If
                                        ElseIf slLive = "M" Then
                                            If tmCrf.sLiveCopy <> "M" Then
                                                ilBypassCrf = True
                                            End If
                                        ElseIf slLive = "S" Then
                                            If tmCrf.sLiveCopy <> "S" Then
                                                ilBypassCrf = True
                                            End If
                                        ElseIf slLive = "P" Then
                                            If tmCrf.sLiveCopy <> "P" Then
                                                ilBypassCrf = True
                                            End If
                                        ElseIf slLive = "Q" Then
                                            If tmCrf.sLiveCopy <> "Q" Then
                                                ilBypassCrf = True
                                            End If
                                        Else
                                            If (tmCrf.sLiveCopy = "L") Or (tmCrf.sLiveCopy = "M") Or (tmCrf.sLiveCopy = "S") Or (tmCrf.sLiveCopy = "P") Or (tmCrf.sLiveCopy = "Q") Then
                                                ilBypassCrf = True
                                            End If
                                        End If
                                        If Not ilBypassCrf Then
                                            If (tmCrf.lRafCode > 0) And (Trim$(tmCrf.sZone) = "R") Then
                                                If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY Then
                                                    'ilVpf = gBinarySearchVpf(tmCrf.iVefCode)
                                                    'If ilVpf <> -1 Then
                                                    '    If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                                    '        ilBypassCrf = True
                                                    '    End If
                                                    'Else
                                                    '    ilBypassCrf = True
                                                    'End If
                                                    ilVef = gBinarySearchVef(ilCrfVefCode)  '(tmCrf.iVefCode)
                                                    If ilVef <> -1 Then
                                                        '5/11/11: Allow selling vehicles to be set to No
                                                        'If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Then
                                                        If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "S") Then
                                                            ilVpf = gBinarySearchVpf(ilCrfVefCode)  '(tmCrf.iVefCode)
                                                            If ilVpf <> -1 Then
                                                                If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                                                    ilBypassCrf = True
                                                                End If
                                                            Else
                                                                ilBypassCrf = True
                                                            End If
                                                        'ElseIf (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "P") Then
                                                        ElseIf (tgMVef(ilVef).sType = "P") Then
                                                            ilBypassCrf = False
                                                        Else
                                                            ilBypassCrf = True
                                                        End If
                                                    Else
                                                        ilBypassCrf = True
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If (tmCrf.sDay(ilDay) = "Y") And (tmSdf.iLen = tmCrf.iLen) And (Not ilBypassCrf) Then
                                            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                                            llSAsgnDate = gDateValue(slDate)
                                            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                                            llEAsgnDate = gDateValue(slDate)
                                            If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                                                gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
                                                llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                                                gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
                                                llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                                                If tmCrf.sAirGameType = "G" Then
                                                    llSAsgnTime = 999999
                                                    llEAsgnTime = -1
                                                    If tmSsf.iType > 0 Then
                                                        tmCafSrchKey1.lCrfCode = tmCrf.lCode
                                                        ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                                            If tmCaf.sType = "G" Then
                                                                If tmCaf.iGameNo = tmSsf.iType Then
                                                                    llSAsgnTime = llSpotTime
                                                                    llEAsgnTime = llSpotTime
                                                                    Exit Do
                                                                End If
                                                            End If
                                                            ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    End If
                                                ElseIf tmCrf.sAirGameType = "T" Then
                                                    llSAsgnTime = 999999
                                                    llEAsgnTime = -1
                                                    If tmSsf.iType > 0 Then
                                                        tmCafSrchKey1.lCrfCode = tmCrf.lCode
                                                        ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                                            If tmCaf.sType = "T" Then
                                                                If (tmCaf.iTeamMnfCode = tmGsf.iHomeMnfCode) Or (tmCaf.iTeamMnfCode = tmGsf.iVisitMnfCode) Then
                                                                    llSAsgnTime = llSpotTime
                                                                    llEAsgnTime = llSpotTime
                                                                    Exit Do
                                                                End If
                                                            End If
                                                            ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    End If
                                                End If
                                                If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                                                    ilAvailOk = True    'Ok to book into
                                                    If ((tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O")) And (ilType <> -1) Then
                                                        'This line was required for assign copy to missed spots
                                                        'it is left in even though assigning to missed has been removed
                                                        If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (tmSdf.iRotNo <> -1) Then    'Add spot
                                                            ilEvtIndex = 1
                                                            Do
                                                                If ilEvtIndex > tmSsf.iCount Then
                                                                    If tmSsf.iType = 0 Then
                                                                        imSsfRecLen = Len(tmSsf)
                                                                        ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                                                            ilEvtIndex = 1
                                                                        Else
                                                                            ilDayDone = True
                                                                            Exit Do 'Ssf can't be found
                                                                        End If
                                                                    Else
                                                                        ilDayDone = True
                                                                        Exit Do
                                                                    End If
                                                                End If
                                                                'Scan for avail that matches time of spot- then test avail name
                                                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
                                                                'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                                                                If (tmAvail.iRecType = 2) Or ((tmAvail.iRecType >= 6) And (tmAvail.iRecType <= 9)) Then 'Contract Avail subrecord
                                                                    'Test time-
                                                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                                                                    If llSpotTime = llAvailTime Then
                                                                        If tmCrf.sInOut = "I" Then
                                                                            If tmCrf.ianfCode <> tmAvail.ianfCode Then
                                                                                ilAvailOk = False   'No
                                                                            End If
                                                                        Else
                                                                            If tmCrf.ianfCode = tmAvail.ianfCode Then
                                                                                ilAvailOk = False   'No
                                                                            End If
                                                                        End If
                                                                        Exit Do
                                                                    ElseIf (llSpotTime < llAvailTime) And (tmSdf.sXCrossMidnight <> "Y") Then
                                                                        'Spot missing from Ssf
                                                                        ilDayDone = True
                                                                        Exit Do         'Don't assign
                                                                    End If
                                                                End If
                                                                ilEvtIndex = ilEvtIndex + 1
                                                            Loop
                                                            '12/31/14
                                                            If ilDayDone Then
                                                                'Exit Do
                                                                ilAvailOk = False
                                                            End If
                                                        End If
                                                    End If
                                                    If ilAvailOk Then
                                                        'Check if rotation previously assign
                                                        'and requires replacing
                                                        '12/31/14
                                                        If Not mSupersedeRot(ilAllZones, ilSpotAsgn) Then
                                                        '    Exit Do
                                                        End If
                                                        If Not ilSpotAsgn Then
                                                            blAnyAssigned = True
                                                            ilRet = mAsgnCopy(llDate, llCrfCode, llSdfRecPos, ilAsgnDate0, ilAsgnDate1, ilPFAsgn, ilAllZones, ilSpotAsgn, ilCrfVefCode)
                                                            If Not ilRet Then
                                                                '6/4/16: Replaced GoSub
                                                                'GoSub AssignClose
                                                                mAssignClose hlSmf, hlSsf
                                                                gAssignCopyToSpots = False
                                                                Exit Function
                                                            End If
                                                            tmCrf = tlCrf
                                                            If tmCrf.iVefCode > 0 Then
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        '12/31/14
                                        'If ilAllZones Then
                                        '    Exit Do
                                        'End If
                                    End If
                                Next ilVefPass
                            End If
                            ''Reposition to Crf so GetNext is correct
                            ''ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            'tmCrfSrchKey.lCode = llCrfCode
                            'ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If blAnyAssigned Then
                                'tmCrfSrchKey1.sRotType = tmCrf.sRotType
                                'tmCrfSrchKey1.iEtfCode = tmCrf.iEtfCode
                                'tmCrfSrchKey1.iEnfCode = tmCrf.iEnfCode
                                'tmCrfSrchKey1.iadfCode = tmCrf.iadfCode
                                'tmCrfSrchKey1.lChfCode = tmCrf.lChfCode
                                'tmCrfSrchKey1.lFsfCode = 0
                                'tmCrfSrchKey1.iVefCode = tmCrf.iVefCode
                                'tmCrfSrchKey1.iRotNo = tmCrf.iRotNo
                                'ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                tmCrfSrchKey4.sRotType = slType
                                tmCrfSrchKey4.iEtfCode = 0
                                tmCrfSrchKey4.iEnfCode = 0
                                tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
                                tmCrfSrchKey4.lChfCode = tmSdf.lChfCode
                                tmCrfSrchKey4.lFsfCode = 0
                                tmCrfSrchKey4.iRotNo = ilRotNo
                                ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
                            Else
                                ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            End If
                            If blAnyAssigned Then
                                Do While ilRet = BTRV_ERR_NONE
                                    blCrfFound = False
                                    For ilCrf = 0 To UBound(llProcessedCrfCode) - 1 Step 1
                                        If llProcessedCrfCode(ilCrf) = tmCrf.lCode Then
                                            blCrfFound = True
                                            Exit For
                                        End If
                                    Next ilCrf
                                    If Not blCrfFound Then
                                        Exit Do
                                    End If
                                    ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                        Loop
                    '    If ilDayDone Then
                    '        Exit Do
                    '    End If
                    '    If ilPkgVefCode > 0 Then
                    '        ilCrfVefCode = ilPkgVefCode
                    '        ilPkgVefCode = 0
                    '    ElseIf ilSchPkgVefCode > 0 Then
                    '        ilCrfVefCode = ilSchPkgVefCode
                    '        ilSchPkgVefCode = 0
                    '    Else
                    '        If (ilOrigCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                    '            Exit Do
                    '        End If
                    '        ilCrfVefCode = ilLnVefCode
                    '        ilLnVefCode = 0
                    '    End If
                    'Loop While ilCrfVefCode > 0
                End If
                Do
                    ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    If tmSdf.iVefCode <> ilVefCode Then
                        Exit Do
                    End If
                    If (tmSdf.iDate(0) <> ilAsgnDate0) Or (tmSdf.iDate(1) <> ilAsgnDate1) Then
                        Exit Do
                    End If
                    If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        If ilType <> -1 Then
                            If tmSdf.iGameNo = tmSsf.iType Then
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                Loop
            Loop
            If (ilType = -1) Then
                Exit Do
            End If
            imSsfRecLen = Len(tmSsf)
            ilRet = btrGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next llDate
    lgAssignEnd = timeGetTime
    lgAssignTotal = lgAssignTotal + (lgAssignEnd - lgAssignStart)
    'ilRet = btrEndTrans(hlSdf)
    '6/4/16: Replaced GoSub
    'GoSub AssignClose
    mAssignClose hlSmf, hlSsf
    gAssignCopyToSpots = True
    Exit Function
'AssignClose:
''    btrDestroy hlRdf
'    On Error Resume Next
'    btrDestroy hmCvf
'    btrDestroy hmGsf
'    btrDestroy hmCaf
'    btrDestroy hlSmf
'    btrDestroy hmRsf
'    btrDestroy hlSsf
'    btrDestroy hmClf
'    btrDestroy hmTzf
'    btrDestroy hmCif
'    btrDestroy hmCnf
'    btrDestroy hmCrf
'    btrDestroy hmSdf
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAsgnCopy                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Assign copy to a spot          *
'*                                                     *
'*******************************************************
Function mAsgnCopy(llDate As Long, llCrfCode As Long, llSdfRecPos As Long, ilAsgnDate0 As Integer, ilAsgnDate1 As Integer, ilPFAsgn As Integer, ilAllZones As Integer, ilSpotAsgn As Integer, ilAsgnVefCode As Integer) As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilFound As Integer
    Dim llCifRecPos As Long
    Dim llTzfRecPos As Long
    Dim slPtType As String
    Dim ilZone As Integer
    Dim slDate As String
    Dim ilVpf As Integer
    Dim llLastLogDate As Long
    Dim llNowDate As Long
    Dim slLogDate As String

    ilAllZones = False
    ilRet = btrBeginTrans(hmSdf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        ilCRet = btrAbortTrans(hmSdf)
        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/BeginTrans, Try Later", vbOKOnly + vbExclamation, "Erase")
        mAsgnCopy = False
        Exit Function
    End If
    '4/2/12: To aviod deadlock with BookSpot, read sdf prior to crf
    If tmSdf.iGameNo <= 0 Then
        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
    Else
        'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY6, BTRV_LOCK_NONE)
        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
    End If
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        ilCRet = btrAbortTrans(hmSdf)
        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Sdf(23), Try Later", vbOKOnly + vbExclamation, "Erase")
        mAsgnCopy = False
        Exit Function
    End If
    
    'Assign copy, find next instruction to be assigned
    '9/30/16: add sync rotation test
    If tmCrf.iBkoutInstAdfCode <> POOLROTATION Then
        tmCnfSrchKey.lCrfCode = tmCrf.lCode
        If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
            If ilPFAsgn = 1 Then    'Final
                tmCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextFinal(igAsgnVehRowNo)
            Else
                tmCnfSrchKey.iInstrNo = tgAsgnVehCvf.iNextPrelim(igAsgnVehRowNo)
            End If
            If tmCnfSrchKey.iInstrNo = 0 Then
                tmCnfSrchKey.iInstrNo = 1
            End If
        Else
            If ilPFAsgn = 1 Then    'Final
                tmCnfSrchKey.iInstrNo = tmCrf.iNextFinal
            Else
                tmCnfSrchKey.iInstrNo = tmCrf.iNextPrelim
            End If
        End If
        ilRet = btrGetEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            tmCnfSrchKey.lCrfCode = tmCrf.lCode
            tmCnfSrchKey.iInstrNo = 1
            ilRet = btrGetEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
    Else
        ilRet = BTRV_ERR_NONE
    End If
    If ilRet = BTRV_ERR_NONE Then
        'One inventory per instruction
        If tmCrf.iBkoutInstAdfCode <> POOLROTATION Then
            If tmCnf.lCifCode > 0 Then
                tmCifSrchKey.lCode = tmCnf.lCifCode
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Else
                'Not coded
                ilRet = Not BTRV_ERR_NONE
            End If
        Else
            ilRet = BTRV_ERR_NONE
            tmCnf.iInstrNo = 0
            tmCrf.iNextFinal = 0
        End If
        If ilRet = BTRV_ERR_NONE Then
            'ilRet = btrGetPosition(hmCrf, llCrfRecPos)
            Do
                'ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, llCrfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                tmCrfSrchKey.lCode = llCrfCode
                ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmSdf)
                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Crf(1), Try Later", vbOKOnly + vbExclamation, "Erase")
                    mAsgnCopy = False
                    Exit Function
                End If
                'tmSRec = tmCrf
                'ilRet = gGetByKeyForUpdate("Crf", hmCrf, tmSRec)
                'tmCrf = tmSRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilCRet = btrAbortTrans(hmSdf)
                '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Crf(2), Try Later", vbOkOnly + vbExclamation, "Erase")
                '    mAsgnCopy = False
                '    Exit Function
                'End If
                If ilPFAsgn = 1 Then    'Final
                    tmCrf.iNextFinal = tmCnf.iInstrNo + 1
                    tmCrf.iNextPrelim = tmCrf.iNextFinal
                Else
                    tmCrf.iNextPrelim = tmCnf.iInstrNo + 1
                End If
                If (tmCrf.iEarliestDateAssg(0) = 0) And (tmCrf.iEarliestDateAssg(1) = 0) Then
                    tmCrf.iEarliestDateAssg(0) = ilAsgnDate0
                    tmCrf.iEarliestDateAssg(1) = ilAsgnDate1
                    tmCrf.iLatestDateAssg(0) = ilAsgnDate0
                    tmCrf.iLatestDateAssg(1) = ilAsgnDate1
                Else
                    If (ilAsgnDate1 < tmCrf.iEarliestDateAssg(1)) Or ((ilAsgnDate1 = tmCrf.iEarliestDateAssg(1)) And (ilAsgnDate0 < tmCrf.iEarliestDateAssg(0))) Then
                        tmCrf.iEarliestDateAssg(0) = ilAsgnDate0
                        tmCrf.iEarliestDateAssg(1) = ilAsgnDate1
                    End If
                    If (ilAsgnDate1 > tmCrf.iLatestDateAssg(1)) Or ((ilAsgnDate1 = tmCrf.iLatestDateAssg(1)) And (ilAsgnDate0 > tmCrf.iLatestDateAssg(0))) Then
                        tmCrf.iLatestDateAssg(0) = ilAsgnDate0
                        tmCrf.iLatestDateAssg(1) = ilAsgnDate1
                    End If
                End If
                tmCrf.iLastDateAssg(0) = ilAsgnDate0
                tmCrf.iLastDateAssg(1) = ilAsgnDate1
                gPackDate smNowDate, tmCrf.iDateAssgDone(0), tmCrf.iDateAssgDone(1)
                gPackTime smNowTime, tmCrf.iTimeAssgDone(0), tmCrf.iTimeAssgDone(1)
                ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmSdf)
                ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Crf(3), Try Later", vbOKOnly + vbExclamation, "Erase")
                mAsgnCopy = False
                Exit Function
            End If
            '9/30/16: Update pointers
            If tmCrf.iBkoutInstAdfCode <> POOLROTATION Then
                If (tgSaf(0).sSyncCopyInRot = "Y") And (igAsgnVehRowNo <> -1) Then
                    imCvfRecLen = Len(tmCvf)
                    tmCvfSrchKey.lCode = tgAsgnVehCvf.lCode
                    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If ilPFAsgn = 1 Then    'Final
                            tmCvf.iNextFinal(igAsgnVehRowNo) = tmCnf.iInstrNo + 1
                            tmCvf.iNextPrelim(igAsgnVehRowNo) = tmCnf.iInstrNo + 1
                        Else
                            tmCvf.iNextPrelim(igAsgnVehRowNo) = tmCnf.iInstrNo + 1
                        End If
                        ilRet = btrUpdate(hmCvf, tmCvf, imCvfRecLen)
                        tgAsgnVehCvf = tmCvf
                    End If
                End If
                If ilPFAsgn = 1 Then   'Final
                    ilRet = btrGetPosition(hmCif, llCifRecPos)
                    Do
                        ilRet = btrGetDirect(hmCif, tmCif, imCifRecLen, llCifRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            ilCRet = btrAbortTrans(hmSdf)
                            ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Cif(4), Try Later", vbOKOnly + vbExclamation, "Erase")
                            mAsgnCopy = False
                            Exit Function
                        End If
                        'tmSRec = tmCif
                        'ilRet = gGetByKeyForUpdate("Cif", hmCif, tmSRec)
                        'tmCif = tmSRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    ilCRet = btrAbortTrans(hmSdf)
                        '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Cif(5), Try Later", vbOkOnly + vbExclamation, "Erase")
                        '    mAsgnCopy = False
                        '    Exit Function
                        'End If
                        If (tmCif.iUsedDate(0) = 0) And (tmCif.iUsedDate(1) = 0) Then
                            tmCif.iUsedDate(0) = ilAsgnDate0
                            tmCif.iUsedDate(1) = ilAsgnDate1
                        Else
                            gUnpackDate tmCif.iUsedDate(0), tmCif.iUsedDate(1), slDate
                            If llDate > gDateValue(slDate) Then
                                tmCif.iUsedDate(0) = ilAsgnDate0
                                tmCif.iUsedDate(1) = ilAsgnDate1
                            'Else
                            '    Exit Do 'Don't update date
                            End If
                        End If
                        tmCif.iNoTimesAir = tmCif.iNoTimesAir + 1
                        'DL:7/1/03, Wrap value around to avoid overflow error
                        If tmCif.iNoTimesAir > 32000 Then
                            tmCif.iNoTimesAir = 0
                        End If
                        ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        ilCRet = btrAbortTrans(hmSdf)
                        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Cif(6), Try Later", vbOKOnly + vbExclamation, "Erase")
                        mAsgnCopy = False
                        Exit Function
                    End If
                End If
            End If
            If (tmCrf.lRafCode <= 0) Or (Trim$(tmCrf.sZone) <> "R") Then

                slPtType = tmSdf.sPtType
                If slPtType = "3" Then 'Time zone copy- determine if supersede or add to tzf
                    If Trim$(tmCrf.sZone) = "" Then 'All zone copy- test if add or superseded
                        'Determine if all previous Zone and Other copy superseded
                        ilFound = False
                        For ilZone = 1 To 6 Step 1
                            If (tmTzf.lCifZone(ilZone - 1) <> 0) And (tmCrf.iRotNo > tmTzf.iRotNo(ilZone - 1)) Then
                                tmTzf.sZone(ilZone - 1) = ""
                                tmTzf.lCifZone(ilZone - 1) = 0
                                tmTzf.iRotNo(ilZone - 1) = 0
                            ElseIf tmTzf.lCifZone(ilZone - 1) <> 0 Then
                                ilFound = True  'Instruction not superseded
                            End If
                        Next ilZone
                        If ilFound Then
                            'Insert into first hold
                            For ilZone = 1 To 6 Step 1
                                If tmTzf.lCifZone(ilZone - 1) = 0 Then
                                    ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                                    Do
                                        ilRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                                ilRet = csiHandleValue(0, 7)
                                            End If
                                            ilCRet = btrAbortTrans(hmSdf)
                                            ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Tzf(7), Try Later", vbOKOnly + vbExclamation, "Erase")
                                            mAsgnCopy = False
                                            Exit Function
                                        End If
                                        'tmSRec = tmTzf
                                        'ilRet = gGetByKeyForUpdate("Tzf", hmTzf, tmSRec)
                                        'tmTzf = tmSRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    ilCRet = btrAbortTrans(hmSdf)
                                        '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Tzf(8), Try Later", vbOkOnly + vbExclamation, "Erase")
                                        '    mAsgnCopy = False
                                        '    Exit Function
                                        'End If
                                        If tmCnf.lCifCode > 0 Then
                                            tmTzf.lCifZone(ilZone - 1) = tmCnf.lCifCode
                                        ElseIf tmCnf.lCifCode < 0 Then
                                            tmTzf.lCifZone(ilZone - 1) = -tmCnf.lCifCode
                                        End If
                                        If ilPFAsgn = 1 Then    'Final
                                            tmTzf.iRotNo(ilZone - 1) = tmCrf.iRotNo
                                        Else
                                            tmTzf.iRotNo(ilZone - 1) = 0
                                        End If
                                        tmTzf.sZone(ilZone - 1) = "Oth"
                                        ilRet = btrUpdate(hmTzf, tmTzf, imTzfRecLen)
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                            ilRet = csiHandleValue(0, 7)
                                        End If
                                        ilCRet = btrAbortTrans(hmSdf)
                                        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Tzf(9), Try Later", vbOKOnly + vbExclamation, "Erase")
                                        mAsgnCopy = False
                                        Exit Function
                                    End If
                                    Exit For
                                End If
                            Next ilZone
                        Else
                            'Remove Tzf
                            ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                            Do
                                ilRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                If ilRet <> BTRV_ERR_NONE Then
                                    ilCRet = btrAbortTrans(hmSdf)
                                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Tzf(10), Try Later", vbOKOnly + vbExclamation, "Erase")
                                    mAsgnCopy = False
                                    Exit Function
                                End If
                                'tmSRec = tmTzf
                                'ilRet = gGetByKeyForUpdate("Tzf", hmTzf, tmSRec)
                                'tmTzf = tmSRec
                                ilRet = btrDelete(hmTzf)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                ilCRet = btrAbortTrans(hmSdf)
                                ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Delete Tzf(11), Try Later", vbOKOnly + vbExclamation, "Erase")
                                mAsgnCopy = False
                                Exit Function
                            End If
                            slPtType = ""
                        End If
                    Else
                        ilFound = False
                        For ilZone = 1 To 6 Step 1
                            If tmCrf.sZone = tmTzf.sZone(ilZone - 1) Then
                                'Replace zone
                                ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                                Do
                                    ilRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                            ilRet = csiHandleValue(0, 7)
                                        End If
                                        ilCRet = btrAbortTrans(hmSdf)
                                        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Tzf(12), Try Later", vbOKOnly + vbExclamation, "Erase")
                                        mAsgnCopy = False
                                        Exit Function
                                    End If
                                    'tmSRec = tmTzf
                                    'ilRet = gGetByKeyForUpdate("Tzf", hmTzf, tmSRec)
                                    'tmTzf = tmSRec
                                    'If ilRet <> BTRV_ERR_NONE Then
                                    '    ilCRet = btrAbortTrans(hmSdf)
                                    '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Tzf(13), Try Later", vbOkOnly + vbExclamation, "Erase")
                                    '    mAsgnCopy = False
                                    '    Exit Function
                                    'End If
                                    If tmCnf.lCifCode > 0 Then
                                        tmTzf.lCifZone(ilZone - 1) = tmCnf.lCifCode
                                    ElseIf tmCnf.lCifCode < 0 Then
                                        tmTzf.lCifZone(ilZone - 1) = -tmCnf.lCifCode
                                    End If
                                    If ilPFAsgn = 1 Then    'Final
                                        tmTzf.iRotNo(ilZone - 1) = tmCrf.iRotNo
                                    Else
                                        tmTzf.iRotNo(ilZone - 1) = 0
                                    End If
                                    ilRet = btrUpdate(hmTzf, tmTzf, imTzfRecLen)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                If ilRet <> BTRV_ERR_NONE Then
                                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    ilCRet = btrAbortTrans(hmSdf)
                                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Tzf(14), Try Later", vbOKOnly + vbExclamation, "Erase")
                                    mAsgnCopy = False
                                    Exit Function
                                End If
                                ilFound = True
                                Exit For
                            End If
                        Next ilZone
                        'Add copy to tzf
                        If Not ilFound Then
                            For ilZone = 1 To 6 Step 1
                                If tmTzf.lCifZone(ilZone - 1) = 0 Then
                                    ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                                    Do
                                        ilRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                                ilRet = csiHandleValue(0, 7)
                                            End If
                                            ilCRet = btrAbortTrans(hmSdf)
                                            ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Tzf(15), Try Later", vbOKOnly + vbExclamation, "Erase")
                                            mAsgnCopy = False
                                            Exit Function
                                        End If
                                        'tmSRec = tmTzf
                                        'ilRet = gGetByKeyForUpdate("Tzf", hmTzf, tmSRec)
                                        'tmTzf = tmSRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    ilCRet = btrAbortTrans(hmSdf)
                                        '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Tzf(16), Try Later", vbOkOnly + vbExclamation, "Erase")
                                        '    mAsgnCopy = False
                                        '    Exit Function
                                        'End If
                                        If tmCnf.lCifCode > 0 Then
                                            tmTzf.lCifZone(ilZone - 1) = tmCnf.lCifCode
                                        ElseIf tmCnf.lCifCode < 0 Then
                                            tmTzf.lCifZone(ilZone - 1) = -tmCnf.lCifCode
                                        End If
                                        If ilPFAsgn = 1 Then    'Final
                                            tmTzf.iRotNo(ilZone - 1) = tmCrf.iRotNo
                                        Else
                                            tmTzf.iRotNo(ilZone - 1) = 0
                                        End If
                                        tmTzf.sZone(ilZone - 1) = tmCrf.sZone
                                        ilRet = btrUpdate(hmTzf, tmTzf, imTzfRecLen)
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    If ilRet <> BTRV_ERR_NONE Then
                                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                            ilRet = csiHandleValue(0, 7)
                                        End If
                                        ilCRet = btrAbortTrans(hmSdf)
                                        ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Tzf(17), Try Later", vbOKOnly + vbExclamation, "Erase")
                                        mAsgnCopy = False
                                        Exit Function
                                    End If
                                    Exit For
                                End If
                            Next ilZone
                        End If
                    End If
                Else
                    If Trim$(tmCrf.sZone) <> "" Then    'Create time zone copy
                        'For ilZone = 1 To 6 Step 1
                        For ilZone = 0 To 5 Step 1
                            tmTzf.lCifZone(ilZone) = 0
                            tmTzf.iRotNo(ilZone) = 0
                            tmTzf.sZone(ilZone) = ""
                        Next ilZone
                        If tmCnf.lCifCode > 0 Then
                            'tmTzf.lCifZone(1) = tmCnf.lCifCode
                            tmTzf.lCifZone(0) = tmCnf.lCifCode
                        ElseIf tmCnf.lCifCode < 0 Then
                            'tmTzf.lCifZone(1) = -tmCnf.lCifCode
                            tmTzf.lCifZone(0) = -tmCnf.lCifCode
                        End If
                        If ilPFAsgn = 1 Then    'Final
                            'tmTzf.iRotNo(1) = tmCrf.iRotNo
                            tmTzf.iRotNo(0) = tmCrf.iRotNo
                        Else
                            'tmTzf.iRotNo(1) = 0
                            tmTzf.iRotNo(0) = 0
                        End If
                        'tmTzf.sZone(1) = tmCrf.sZone
                        tmTzf.sZone(0) = tmCrf.sZone
                        If (slPtType = "1") Or (slPtType = "2") Then
                            'tmTzf.lCifZone(2) = tmSdf.lCopyCode
                            tmTzf.lCifZone(1) = tmSdf.lCopyCode
                            If ilPFAsgn = 1 Then    'Final
                                'tmTzf.iRotNo(2) = tmSdf.iRotNo
                                tmTzf.iRotNo(1) = tmSdf.iRotNo
                            Else
                                'tmTzf.iRotNo(2) = 0
                                tmTzf.iRotNo(1) = 0
                            End If
                            'tmTzf.sZone(2) = "Oth"
                            tmTzf.sZone(1) = "Oth"
                        End If
                        tmTzf.lCode = 0
                        ilRet = btrInsert(hmTzf, tmTzf, imTzfRecLen, INDEXKEY0)
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            ilCRet = btrAbortTrans(hmSdf)
                            ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Insert Tzf(17), Try Later", vbOKOnly + vbExclamation, "Erase")
                            mAsgnCopy = False
                            Exit Function
                        End If
                    End If
                End If
                If tmSdf.iGameNo <= 0 Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                Else
                    'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY6, BTRV_LOCK_NONE)
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmSdf)
                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Sdf(18), Try Later", vbOKOnly + vbExclamation, "Erase")
                    mAsgnCopy = False
                    Exit Function
                End If
                ilRet = mTestRegionRot(ilPFAsgn)
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmSdf)
                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/mTestRegionRot, Try Later", vbOKOnly + vbExclamation, "Erase")
                    mAsgnCopy = False
                    Exit Function
                End If
                'tmSRec = tmSdf
                'ilRet = gGetByKeyForUpdate("Sdf", hmSdf, tmSRec)
                'tmSdf = tmSRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilCRet = btrAbortTrans(hmSdf)
                '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Sdf(19), Try Later", vbOkOnly + vbExclamation, "Erase")
                '    mAsgnCopy = False
                '    Exit Function
                'End If
                Do
                    If Trim$(tmCrf.sZone) = "" Then
                        If slPtType <> "3" Then
                            If ilPFAsgn = 1 Then    'Final
                                tmSdf.iRotNo = tmCrf.iRotNo
                            Else
                                tmSdf.iRotNo = 0
                            End If
                            If tmCnf.lCifCode > 0 Then
                                tmSdf.sPtType = "1"
                                tmSdf.lCopyCode = tmCnf.lCifCode
                            ElseIf tmCnf.lCifCode < 0 Then
                                tmSdf.sPtType = "2"
                                tmSdf.lCopyCode = -tmCnf.lCifCode
                            End If
                        Else
                            tmSdf.sPtType = "3"
                            tmSdf.lCopyCode = tmTzf.lCode
                        End If
                    Else
                        tmSdf.sPtType = "3"
                        tmSdf.lCopyCode = tmTzf.lCode
                    End If
                    ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        If tmSdf.iGameNo <= 0 Then
                            ilCRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                        Else
                            'ilCRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY6, BTRV_LOCK_NONE)
                            ilCRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                        End If
                        If ilCRet <> BTRV_ERR_NONE Then
                            ilRet = ilCRet
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            ilCRet = btrAbortTrans(hmSdf)
                            ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/GetDirect Sdf(20), Try Later", vbOKOnly + vbExclamation, "Erase")
                            mAsgnCopy = False
                            Exit Function
                        End If
                        'tmSRec = tmSdf
                        'ilCRet = gGetByKeyForUpdate("Sdf", hmSdf, tmSRec)
                        'tmSdf = tmSRec
                        'If ilCRet <> BTRV_ERR_NONE Then
                        '    ilRet = ilCRet
                        '    ilCRet = btrAbortTrans(hmSdf)
                        '    ilRet = MsgBox("Assign Not Completed" & Str$(ilRet) & "/GetByKey Sdf(21), Try Later", vbOkOnly + vbExclamation, "Erase")
                        '    mAsgnCopy = False
                        '    Exit Function
                        'End If
                        If (tmSdf.sSchStatus <> "S") And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                            ilAllZones = True
                            Exit Do
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmSdf)
                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/Update Sdf(22), Try Later", vbOKOnly + vbExclamation, "Erase")
                    mAsgnCopy = False
                    Exit Function
                End If
                'Reset to spot by index1 as gGetByKey.. is by Index3
                If tmSdf.iGameNo <= 0 Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                Else
                    'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY6, BTRV_LOCK_NONE)
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                End If
                ilSpotAsgn = True
                If Trim$(tmCrf.sZone) = "" Then
                    ilAllZones = True
                    '3/24/12:  Temporary save rotation number in sdf
                    If (ilPFAsgn = 0) And (tmCrf.iRotNo > tmSdf.iRotNo) Then    'Prel
                        tmSdf.iRotNo = tmCrf.iRotNo
                    End If
                End If
            Else
                'Reset to spot by index1 as gGetByKey.. is by Index3
                If tmSdf.iGameNo <= 0 Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                Else
                    'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY6, BTRV_LOCK_NONE)
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                End If
                ilRet = mAsgnRegionCopy(ilPFAsgn, ilAsgnVefCode)
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmSdf)
                    ilRet = MsgBox("Assign Not Completed" & str$(ilRet) & "/mAsgnRegionCopy, Try Later", vbOKOnly + vbExclamation, "Erase")
                    mAsgnCopy = False
                    Exit Function
                End If
            End If
        End If
    End If
    ilRet = btrEndTrans(hmSdf)
    'ilVpf = gBinarySearchVpfPlus(tmSdf.iVefCode)
    'If ilVpf > 0 Then
    '    If (tgVpf(ilVpf).iLLD(0) <> 0) Or (tgVpf(ilVpf).iLLD(1) <> 0) Then
    '        gUnpackDateLong tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), llLastLogDate
    '    Else
    '        llLastLogDate = -1
    '    End If
    '    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    '    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slLogDate
    '    If (gDateValue(slLogDate) > llNowDate) And (gDateValue(slLogDate) <= llLastLogDate) Then
    '        ilRet = gAlertAdd("L", "C", 0, tmSdf.iVefCode, slLogDate)
    '    End If
    'End If
    gMakeLogAlert tmSdf, "C", hmGsf
    mAsgnCopy = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAsgnRegionCopy                 *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Assign Regional Copy           *
'*                                                     *
'*******************************************************
Function mAsgnRegionCopy(ilPFAsgn As Integer, ilAsgnVefCode As Integer) As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    ilFound = False
    '7/15/14
    tmRsfSrchKey3.lSdfCode = tmSdf.lCode
    tmRsfSrchKey3.lRafCode = tmCrf.lRafCode
    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
        If (tmRsf.sType <> "B") And (tmRsf.sType <> "A") Then
            If tmRsf.lRafCode = tmCrf.lRafCode Then
                ilFound = True
                Exit Do
            Else
                'ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Exit Do
            End If
        Else
            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop
    If ilPFAsgn = 1 Then    'Final
        tmRsf.iRotNo = tmCrf.iRotNo
    Else
        tmRsf.iRotNo = 0
    End If
    If tmCrf.iBkoutInstAdfCode <> POOLROTATION Then
        If tmCnf.lCifCode > 0 Then
            tmRsf.sPtType = "1"
            tmRsf.lCopyCode = tmCnf.lCifCode
        ElseIf tmCnf.lCifCode < 0 Then
            tmRsf.sPtType = "2"
            tmRsf.lCopyCode = -tmCnf.lCifCode
        End If
    End If
    tmRsf.lCrfCode = tmCrf.lCode
    If ilFound Then
        gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
        tmRsf.sUnused = ""
        ilRet = btrUpdate(hmRsf, tmRsf, imRsfRecLen)
    Else
        tmRsf.lCode = 0
        tmRsf.lSdfCode = tmSdf.lCode
        tmRsf.lRafCode = tmCrf.lRafCode
        tmRsf.sType = "R"
        tmRsf.lSBofCode = 0
        tmRsf.lRBofCode = 0
        tmRsf.iBVefCode = ilAsgnVefCode 'tmCrf.iVefCode
        tmRsf.lRChfCode = tmCrf.lChfCode
        gPackDate Format(gNow(), "m/d/yy"), tmRsf.iDateAdded(0), tmRsf.iDateAdded(1)
        tmRsf.sUnused = ""
        If tmCrf.iBkoutInstAdfCode = POOLROTATION Then
            'required to know that copy needs to be found and assigned
            tmRsf.sPtType = ""
            tmRsf.lCopyCode = 0
        End If
        ilRet = btrInsert(hmRsf, tmRsf, imRsfRecLen, INDEXKEY0)
    End If
    gMakeLogAlert tmSdf, "C", hmGsf
    mAsgnRegionCopy = ilRet
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSupersedeRot                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check Rotation number          *
'*                                                     *
'*******************************************************
Function mSupersedeRot(ilAllZones As Integer, ilSpotAsgn As Integer) As Integer
    Dim ilRet As Integer
    Dim ilZone As Integer
    mSupersedeRot = True
    If (tmCrf.lRafCode <= 0) Or (Trim$(tmCrf.sZone) <> "R") Then
        If tmSdf.sPtType = "3" Then 'Time zone
            tmTzfSrchKey.lCode = tmSdf.lCopyCode
            ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If Trim$(tmCrf.sZone) = "" Then 'All zones (note zone copy superseded by all zones will be check when assigning)
                    For ilZone = 1 To 6 Step 1
                        If ((Trim$(tmTzf.sZone(ilZone - 1)) = "") Or (StrComp(tmTzf.sZone(ilZone - 1), "Oth", 1) = 0)) And (tmTzf.lCifZone(ilZone - 1) <> 0) Then
                            If tmTzf.iRotNo(ilZone - 1) >= tmCrf.iRotNo Then
                                'Zone copy must be Ok because of way rot # assigned
                                ilAllZones = True
                                ilSpotAsgn = True
                                mSupersedeRot = False
                                Exit Function
                            End If
                        End If
                    Next ilZone
                Else
                    For ilZone = 1 To 6 Step 1
                        If tmTzf.sZone(ilZone - 1) = tmCrf.sZone Then
                            If tmTzf.iRotNo(ilZone - 1) >= tmCrf.iRotNo Then
                                'Zone copy Ok- All Zone must be OK because of way rot # assigned
                                'The above statement isn't true if the user defines superseding zone
                                'copy (same zone) before the first is assigned
                                'ilAllZones = True
                                ilSpotAsgn = True
                                Exit For
                            End If
                        End If
                    Next ilZone
                End If
            End If
        Else
            If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Then
                If tmSdf.iRotNo >= tmCrf.iRotNo Then
                    ilAllZones = True
                    ilSpotAsgn = True
                    mSupersedeRot = False
                    Exit Function
                End If
            End If
        End If
    Else
        '3/29/13: Verify that the region being assigned is newer then the generic rotation assigned previously
        If tmCrf.iRotNo > tmSdf.iRotNo Then
            '7/15/14
            tmRsfSrchKey3.lSdfCode = tmSdf.lCode
            tmRsfSrchKey3.lRafCode = tmCrf.lRafCode
            ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                If tmRsf.lRafCode = tmCrf.lRafCode Then
                    If tmRsf.iRotNo >= tmCrf.iRotNo Then
                        'ilAllZones = True
                        ilSpotAsgn = True
                        'mSupersedeRot = False
                        Exit Function
                    End If
                Else
                    Exit Do
                End If
                ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Else
            ilSpotAsgn = True
            Exit Function
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRegionRot                  *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test Rotation Numbers          *
'*                                                     *
'*******************************************************
Function mTestRegionRot(ilPFAsgn As Integer) As Integer
    Dim ilRet As Integer
    Dim ilValue As Integer
    Dim ilRotNo As Integer
    Dim tlCrf As CRF

    mTestRegionRot = BTRV_ERR_NONE
    ilValue = Asc(tgSpf.sUsingFeatures2)  'Option Fields in Orders/Proposals
    If (Trim$(tmCrf.sZone) = "") And (((ilValue And REGIONALCOPY) = REGIONALCOPY) Or ((ilValue And SPLITCOPY) = SPLITCOPY)) Then 'All zone copy- test if add or superseded
        '7/15/14
        tmRsfSrchKey3.lSdfCode = tmSdf.lCode
        tmRsfSrchKey3.lRafCode = tmCrf.lRafCode
        'ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        ilRet = btrGetGreaterOrEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
            '''4/30/11:  Added rafCode test
            ''3/21/12:  Remove all region superseded by the Generic
            ''          If region assigned, then remove only those regions that match the current region
            ''If (tmRsf.iRotNo < tmCrf.iRotNo) And (tmRsf.lRafCode = tmCrf.lRafCode) Then
            '4/20/14: Preliminary sets RsfRotNo = 0
            ilRotNo = tmRsf.iRotNo
            If (ilRotNo <= 0) Then
                tmCrfSrchKey.lCode = tmRsf.lCrfCode
                ilRet = btrGetEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRotNo = tlCrf.iRotNo
                    If ilPFAsgn <> 0 Then
                        tmRsf.iRotNo = ilRotNo
                        ilRet = btrUpdate(hmRsf, tmRsf, imRsfRecLen)
                    End If
                End If
            End If
            'If (tmRsf.iRotNo < tmCrf.iRotNo) And ((tmRsf.lRafCode = tmCrf.lRafCode) Or (tmCrf.lRafCode <= 0)) Then
            If (ilRotNo < tmCrf.iRotNo) And ((tmRsf.lRafCode = tmCrf.lRafCode) Or (tmCrf.lRafCode <= 0)) Then
                ilRet = btrDelete(hmRsf)
                If (ilRet <> BTRV_ERR_NONE) Then
                    If (ilRet <> BTRV_ERR_CONFLICT) Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mTestRegionRot = ilRet
                        Exit Function
                    End If
                End If
                tmRsfSrchKey3.lSdfCode = tmSdf.lCode
                tmRsfSrchKey3.lRafCode = tmCrf.lRafCode
                'ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                ilRet = btrGetGreaterOrEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_KEY_NOT_FOUND) Or (ilRet = BTRV_ERR_END_OF_FILE) Then
                    mTestRegionRot = BTRV_ERR_NONE
                    Exit Function
                End If
                If (ilRet <> BTRV_ERR_NONE) Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mTestRegionRot = ilRet
                    Exit Function
                End If
            Else
                ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        Loop
    End If
End Function

'Public Function gGetMGCopyAssign(tlInSdf As SDF, ilPkgVefCode As Integer, ilLnVefCode As Integer, slLive As String, ilRdfCode As Integer, hlSmf As Integer, hlCrf As Integer, hlRdf As Integer) As String
Public Function gGetMGCopyAssign(tlInSdf As SDF, ilPkgVefCode As Integer, ilLnVefCode As Integer, slLive As String, hlSmf As Integer, hlCrf As Integer) As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlRdf                         ilRdfRecLen                   tlRdfSrchKey              *
'*                                                                                        *
'******************************************************************************************

    Dim ilValue As Integer
    Dim slType As String
    Dim ilRet As Integer
    Dim ilBypassCrf As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSpotTime As Long
    Dim slMGCrf As String
    Dim ilRotNo As Integer
    Dim ilCrfOk As Integer
    Dim slTime As String
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim tlSdf As SDF
    Dim tlCrf As CRF
    Dim tlSmf As SMF
    Dim tlCrfSrchKey1 As CRFKEY1

    If tlInSdf.sSpotType = "X" Then
        ilValue = Asc(tgSpf.sMOFCopyAssign)
        If (ilValue And FILLORIGVEHONLY) = FILLORIGVEHONLY Then
            gGetMGCopyAssign = "O"
        ElseIf (ilValue And FILLSCHVEHONLY) = FILLSCHVEHONLY Then
            gGetMGCopyAssign = "S"
        Else
            gGetMGCopyAssign = "B"
        End If
        Exit Function
    Else
        'If we want to use the MG setting in Rotation, then replace this code with getting the rotation
        'for the Missed date and time. Check how it is set below instead of site.  Use site if rotation can't be found
        ilValue = Asc(tgSpf.sMOFCopyAssign)
        If (ilValue And MGORIGVEHONLY) = MGORIGVEHONLY Then
            gGetMGCopyAssign = "O"
        ElseIf (ilValue And MGSCHVEHONLY) = MGSCHVEHONLY Then
            gGetMGCopyAssign = "S"
        Else
            gGetMGCopyAssign = "B"
        End If
        If (ilValue And MGRULESINCOPY) <> MGRULESINCOPY Then
            Exit Function
        End If
        If (tlInSdf.sSchStatus <> "G") And (tlInSdf.sSchStatus <> "O") And (tlInSdf.sSchStatus <> "M") Then
            Exit Function
        End If
    End If
    imCrfRecLen = Len(tmCrf)
    ilRotNo = -1
    slMGCrf = ""
    tlSdf = tlInSdf
    If tlInSdf.sSchStatus <> "M" Then
        ilRet = gFindSmf(tlSdf, hlSmf, tlSmf)
        tlSdf.iDate(0) = tlSmf.iMissedDate(0)
        tlSdf.iDate(1) = tlSmf.iMissedDate(1)
        tlSdf.iTime(0) = tlSmf.iMissedTime(0)
        tlSdf.iTime(1) = tlSmf.iMissedTime(1)
    End If
'    ilRdfRecLen = Len(tlRdf)
'    tlRdfSrchKey.iCode = ilRdfCode  ' Rate card program/time File Code
'    ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If tlSdf.sSpotType = "O" Then
        slType = "O"
    ElseIf tlSdf.sSpotType = "C" Then
        slType = "C"
    Else
        slType = "A"
    End If
    tlSdf.iVefCode = ilLnVefCode
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
    ilDay = gWeekDayLong(llDate)
    gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llSpotTime
    Do
        tlCrfSrchKey1.sRotType = slType
        tlCrfSrchKey1.iEtfCode = 0
        tlCrfSrchKey1.iEnfCode = 0
        tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
        tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
        tlCrfSrchKey1.lFsfCode = 0
        tlCrfSrchKey1.iVefCode = tlSdf.iVefCode
        tlCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, imCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = tlSdf.iVefCode)    'ilVefCode)
            ilBypassCrf = False
            'Test if looking for Live or Recorded rotations
            If tlCrf.sState <> "D" Then
                If slLive = "L" Then
                    If tlCrf.sLiveCopy <> "L" Then
                        ilBypassCrf = True
                    End If
                ElseIf slLive = "M" Then
                    If tlCrf.sLiveCopy <> "M" Then
                        ilBypassCrf = True
                    End If
                ElseIf slLive = "S" Then
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
                Else
                    If (tlCrf.sLiveCopy = "L") Or (tlCrf.sLiveCopy = "M") Or (tlCrf.sLiveCopy = "S") Or (tlCrf.sLiveCopy = "P") Or (tlCrf.sLiveCopy = "Q") Then
                        ilBypassCrf = True
                    End If
                End If
            Else
                ilBypassCrf = True
            End If
            If (tlCrf.sDay(ilDay) = "Y") And (tlSdf.iLen = tlCrf.iLen) And (Not ilBypassCrf) Then
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
                        ilCrfOk = True    'Ok to book into
'                        If ((tlCrf.sInOut = "I") Or (tlCrf.sInOut = "O")) And ((tlRdf.sInOut = "I") Or (tlRdf.sInOut = "O")) Then
'                            'Check RDF
'                            If (tlCrf.sInOut = "I") And (tlRdf.sInOut = "I") Then
'                                If tlCrf.ianfCode <> tlRdf.ianfCode Then
'                                    ilCrfOk = False   'No
'                                End If
'                            ElseIf (tlCrf.sInOut = "O") And (tlRdf.sInOut = "O") Then
'                                If tlCrf.ianfCode = tlRdf.ianfCode Then
'                                    ilCrfOk = False   'No
'                                End If
'                            Else
'                                ilCrfOk = False
'                            End If
'                        End If
                        If ilCrfOk Then
                            If ilRotNo = -1 Then
                                ilRotNo = tlCrf.iRotNo
                                slMGCrf = tlCrf.sMGCopyAssign
                            Else
                                If tlCrf.iRotNo > ilRotNo Then
                                    ilRotNo = tlCrf.iRotNo
                                    slMGCrf = tlCrf.sMGCopyAssign
                                End If
                            End If
                            Exit Do
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop

        If ilPkgVefCode > 0 Then
            If tlSdf.iVefCode = ilPkgVefCode Then
                tlSdf.iVefCode = 0
                Exit Do
            End If
            tlSdf.iVefCode = ilPkgVefCode
        Else
            tlSdf.iVefCode = 0
            Exit Do
        End If
    Loop While tlSdf.iVefCode > 0
    If ilRotNo <> -1 Then
        If (slMGCrf = "O") Or (slMGCrf = "S") Or (slMGCrf = "B") Then
            gGetMGCopyAssign = slMGCrf
        End If
    End If
End Function

Public Function gGetMGPkgVefCode(hlClf As Integer, tlInSdf As SDF) As Long
    Dim ilLineNo As Integer
    Dim ilRet As Integer
    Dim tlClfSrchKey As CLFKEY0 'CLF key record image
    Dim ilClfRecLen As Integer  'CLF record length
    Dim tlClf As CLF            'CLF record image

    gGetMGPkgVefCode = 0
    'First check if vehicle scheduled onto is the same as the line and that the line is part of a package.  If so, return package vehicle.
    tlClfSrchKey.lChfCode = tlInSdf.lChfCode
    tlClfSrchKey.iLine = tlInSdf.iLineNo
    tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
    tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
    ilClfRecLen = Len(tlClf)
    ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = tlInSdf.iLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F"))  'And (tlClf.sSchStatus = "A")
        ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = tlInSdf.iLineNo) Then
        If tlClf.iVefCode = tlInSdf.iVefCode Then
            If tlClf.sType = "H" Then
                ilLineNo = tlClf.iPkLineNo
                tlClfSrchKey.lChfCode = tlInSdf.lChfCode
                tlClfSrchKey.iLine = ilLineNo
                tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = ilLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F")) 'And (tlClf.sSchStatus = "A")
                    ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = ilLineNo) Then
                    gGetMGPkgVefCode = tlClf.iVefCode
                End If
            End If
            Exit Function
        End If
    End If
    'Look to line that spot was scheduled into.  If hidden, use package vehicle
    tlClfSrchKey.lChfCode = tlInSdf.lChfCode
    tlClfSrchKey.iLine = 0
    tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
    tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
    ilClfRecLen = Len(tlClf)
    ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode)
        If (tlClf.iVefCode = tlInSdf.iVefCode) And (tlClf.sSchStatus = "F") Then
            If tlClf.sType = "H" Then
                ilLineNo = tlClf.iPkLineNo
                tlClfSrchKey.lChfCode = tlInSdf.lChfCode
                tlClfSrchKey.iLine = ilLineNo
                tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = ilLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F")) 'And (tlClf.sSchStatus = "A")
                    ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlInSdf.lChfCode) And (tlClf.iLine = ilLineNo) Then
                    gGetMGPkgVefCode = tlClf.iVefCode
                    Exit Function
                End If
            End If
        End If
        ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Function

Private Sub mAssignClose(hlSmf As Integer, hlSsf As Integer)
    On Error Resume Next
    btrDestroy hmCvf
    btrDestroy hmGsf
    btrDestroy hmCaf
    btrDestroy hlSmf
    btrDestroy hmRsf
    btrDestroy hlSsf
    btrDestroy hmClf
    btrDestroy hmTzf
    btrDestroy hmCif
    btrDestroy hmCnf
    btrDestroy hmCrf
    btrDestroy hmSdf
End Sub
