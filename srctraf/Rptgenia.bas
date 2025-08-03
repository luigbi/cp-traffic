Attribute VB_Name = "RPTGENIA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptgenia.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmCffSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: rptgenia.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Rm**'Declare Sub ISortT2 Lib "QPRO200.DLL" (Array As Any, Index%, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%)
'Rm**'Declare Sub ArraySortTyp Lib "QPRO200.DLL" (Array() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%, ByVal CaseSenitive%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
'Public tgMkMnf() As MNF 'Market
'Global tmsRec As LPOPREC

Public igFoundLocl As Integer
Public igFoundNatl As Integer
Public igFoundRegl As Integer

Dim hmLef As Integer        'Library Event file handle
Dim tmLef As LEF
Dim imLefRecLen As Integer
Dim tmLefSrchKey As LEFKEY0
Dim imInvNoFound As Integer
Type VEHARRAY
    sVehType As String
    iVehCode As Integer
    iSellAirFlag As Integer
End Type
Type VEH_SIMULCAST_ARRAY
    iParentVehCode As Integer
    iChildVehCode As Integer
    sChildVehName As String
End Type
Type PHFRVFARRAY
    lInvNo As Long
    lCntrNo As Long
    lChfCode As Long
    iAdfCode As Integer
    iAgfCode As Integer
    iSlfCode As Integer
End Type

'D.S. Start 1-23-02
Type LOCNATREG
    iSlfCode As Integer
    iShowLocNatReg As Integer
End Type
'D.S. End 1-23-02


'In RptGen
'Type ODFEXT
'    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sZone As String * 3
'    iEtfCode As Integer         'Event type code
'    iEnfCode As Integer         'Event name code
'    sProgCode As String * 5 'Program code #
'    ianfCode As Integer 'Avail name code
'    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sProduct As String * 35 'Product (either from contract or copy)
'    iMnfSubFeed As Integer
'    iBreakNo As Integer 'Reset at start of each program
'    iPositionNo As Integer
'    lCefCode As Long
'    sShortTitle As String * 15
'End Type
'In RptGenCB
'Type TYPESORT
'    sKey As String * 100
'    lRecPos As Long
'End Type
'In RptGenCB
'Type SPOTTYPESORT
'    sKey As String * 80 'Office Advertiser Contract
'    iVefcode As Integer 'line airing vehicle
'    sCostType As String * 12    'string of spot type (0,00, bonus, adu, recapturable, etc)
'    tSdf As SDF
'End Type
'In RptGenCB
'Type COPYCNTRSORT
'    sKey As String * 80 'Agency, City Contrct # ID Vehicle Len
'    lChfCode As Long
'    iVefcode As Integer
'    sVehName As String * 20
'    iLen As Integer
'    iNoSpots As Integer 'Number of spots that have no copy
'    iNoUnAssg As Integer    'Number of spots not assigned
'    iNoToReassg As Integer    'Number of spots should be reassigned
'End Type
'Type COPYROTNO
'    iRotNo As Integer
'    sZone As String * 3
'End Type
'In RptGenCB
'Type COPYSORT
'    sKey As String * 80 'Agency, City ID
'    iCopyStatus As Integer '0=No Copy; 1=Assigned; 2=Copy but not assigned; 3= Supersede; 4=Zone missing
'    tSdf As SDF
'    sVehName As String * 20
'End Type
'In RptGenCB
'Type SPOTSALE
'    sKey As String * 100    'Vehicle|sSofName|AdvtName|Date or 99999 if total
'    iVefcode As Integer
'    sVehName As String * 20
'    sSofName As String * 20
'    sAdvtName As String * 30
'    lCntrNo As Long
'    lDate As Long
'    sDate As String * 8
'    iCNoSpots As Long            'chged to long 5-12-99 from integer
'    sCGross As String * 12
'    sCCommission As String * 12
'    sCNet As String * 12
'    iTNoSpots As Integer
'    sTGross As String * 12
'    sTCommission As String * 12
'    sTNet As String * 12
'End Type
'In RptGenCB
'Type CODESTNCONV
'    sName As String * 20
'    sCodeStn As String * 5
'End Type
'In RptGenCB
'Type DALLASFDSORT
'    sKey As String * 30
'    sRecord As String * 104
'End Type
'In RptGenCB
'Type VEHICLELLD
'    iVefcode As Integer             'vehicle code
'    iLLD(0 To 1) As Integer         'vehicles last log date
'End Type
'Dim tmSort() As TYPESORT
'Dim tmPlSdf() As SPOTTYPESORT
'Dim tmSpotSOF() As SPOTTYPESORT
'Dim tmCopyCntr() As COPYCNTRSORT
'Dim tmCopy() As COPYSORT
'Dim tmSelAdvt() As Integer
Dim tmSelChf() As Long
Dim tmSelAgf() As Integer
Dim tmSelSlf() As Integer
Dim tmSelAdf() As Integer
'Dim tmRotNo(1 To 6) As COPYROTNO
'Dim tmCodeStn() As CODESTNCONV
'Dim tmDallasFdSort() As DALLASFDSORT
'Dim imSpotSaleVefCode() As Integer
'Dim tmSdfExtSort() As SDFEXTSORT
'Dim tmSdfExt() As SDFEXT
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim imAnfRecLen As Integer        'ANF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0            'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF
Dim hmVsf As Integer            'Vehicle combo file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer            'Agency name file handle
Dim tmAgf As AGF                'AGF record image
Dim tmAgfSrchKey As INTKEY0            'AGF record image
Dim imAgfRecLen As Integer        'AGF record length

Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF

Dim hmSmf As Integer            'Spot makegoods file handle

'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
' Copy Combo Inventory File
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer     'TZF record length
'  Media code File

Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmSlf As Integer            'Salesoerson file handle
Dim tmSlf As SLF                'SLF record image
Dim imSlfRecLen As Integer        'SLF record length
Dim hmMnf As Integer            'MultiName file handle
Dim tmMnf As MNF                'MNF record image
Dim tmMnfSrchKey As INTKEY0            'MNF record image
Dim imMnfRecLen As Integer        'MNF record length
Dim hmSof As Integer            'Sales Office file handle
Dim tmSof As SOF                'SOF record image
Dim tmSofSrchKey As INTKEY0            'SOF record image
Dim imSofRecLen As Integer        'SOF record length


Dim tmIvr As IVR
Dim hmIvr As Integer
Dim imIvrRecLen As Integer        'GPF record length

Dim tmFsf As FSF                    'Feed buffer
Dim hmFsf As Integer                'Feed handle
Dim tmFSFSrchKey As LONGKEY0        'Feed key search
Dim imFsfRecLen As Integer        'FSF record length

Dim tmFnf As FNF                    'Feed name buffer
Dim hmFnf As Integer                'Feed name handle
Dim tmFnfSrchKey As INTKEY0        'Feed name key search
Dim imFnfRecLen As Integer        'FNF record length

Dim tmPrf As PRF                    'Product buffer
Dim hmPrf As Integer                'Product handle
Dim tmPrfSrchKey As LONGKEY0        'Product key search
Dim imPrfRecLen As Integer        'Product record length

'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim tmRvfSrchKey As INTKEY0 'RVF key record image
Dim imRvfRecLen As Integer  'RVF record length
'  Payment History File
Dim hmPhf As Integer        'payment history
Dim tmPhf As RVF            'PHF record image
Dim tmPhfSrchKey As INTKEY0 'PHF key record image
Dim imPhfRecLen As Integer  'PHF record length
Dim hmPnf As Integer            'Personnel file handle
Dim tmPnf As PNF                'PNF record image
Dim tmPnfSrchKey As INTKEY0            'PNF record image
Dim imPnfRecLen As Integer        'PNF record length
Dim hmArf As Integer            'Name/Address file handle
Dim tmArf As ARF                'ARF record image
Dim tmArfSrchKey As INTKEY0            'ARF record image
Dim imArfRecLen As Integer        'ARF record length
Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmProg As PROGRAMSS

Dim tmIihf As IIHF                   'Imported station invoice header file
Dim hmIihf As Integer                 'Handle
Dim imIihfRecLen As Integer          'record length
Dim tmIihfSrchKey1 As IIHFKEY1

Dim tmPostBarterCounts() As POSTBARTERCOUNTS
Type POSTBARTERCOUNTS               'array of contracts and vehicles that are posted by counts (rather than sot times)
    iVefCode As Integer
    lChfCode As Long
End Type

Const LBONE = 1


'********************************************************************************
'*
'*      Procedure Name:gInvAffRpt
'*
'*            Created: 2/16/01     By:D. Smith
'*            Modified:            By:
'*
'*            Comments: Generate Affidavit of Performance reports
'*
'*       8-25-01 Ignore reservation spots.  They were intermixed with
'*               regular spots causing affidavit to be incorrect
'*
'*       1/11/02 D.S. Added Simulcast;
'*                    Added Error checking and more graceful shutdown
'*
'*       1/22/02 D.S. Added local and national filtering
'*
'*       1/30/02 D.S. Added checking for tmPhf.sTranType = "HI" to PHF
'*                    in order to pick up invoice numbers on national contracts
'*
'*       2/4/02 D.S.  Added code to bypass contracts that are assoc. w/markets
'*                    defined as clusters equal yes. The check is done in
'*                    mResetFieldsForIvr
'*      3-23-03 Test advt to see if fill/extra should be shown on inv
'*      1-19-04 change manner in which fill/extra are shown.  Only test advt
'               if the spot price type doesnt have an "-" or "+"
'*      7-27-04 include/exclude contract/feed spots
'       5-26-06 Adjust spot date to true date of air/sched if its a cross-midnight spot
'       11-07-06 determine which contracts the user is allowed to see
'*********************************************************************************
Sub gInvAffRpt()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim tlVef As VEF
    Dim slStartStd As String
    Dim llStartStd As Long
    Dim slEndStd As String
    Dim llEndStd As Long
    Dim llInvDate As Long
    Dim llSptDate As Long
    Dim slTemp As String
    Dim slBuildDate As String
    Dim ilIdx As Integer
    Dim ilVehIdx As Integer
    Dim llNumSdfRecs As Long
    Dim ilSpotOK As Integer
    Dim ilFound As Integer
    Dim slName As String
    Dim slType As String
    Dim llTempCntrNum As Long
    Dim iSavePrgEnfCode As Integer
    Dim tlChfSrchKey1 As CHFKEY1
    Dim tlSSource() As LOCNATREG
    Dim ilLocNatRegFlag As Integer


    ReDim ilStartStd(0 To 1) As Integer
    ReDim ilEndStd(0 To 1) As Integer
    ReDim tlVehArray(0 To 0) As VEHARRAY
    ReDim tlsimulcast(0 To 0) As VEH_SIMULCAST_ARRAY
    ReDim tlPhfRvf(0 To 0) As PHFRVFARRAY
    'ReDim slProduct(1 To 6) As String
    ReDim slProduct(0 To 6) As String   'Index zero ignored
    'ReDim slCopy(1 To 6) As String
    ReDim slCopy(0 To 6) As String  'Index zero ignored
    Dim slShowOnInv As String * 1
    Dim ilDate(0 To 1) As Integer
    Dim ilOKToSeeCnt As Integer     '12-07-06
    Dim slPrice As String
    Dim ilGameSelect As Integer
    Dim ilLineSelect As Integer
    Dim llChfSelect As Long
    Dim ilVpfIndex As Integer
    Dim slLLDate As String
    Dim llLLDate As Long
    Dim slDate As String
    Dim ilLoopOnSelection As Integer
    Dim blSelectionOK As Boolean
    Dim ilLoopOnFile As Integer
    Dim ilIncludeCodes As Integer           'test for inclusion or exclusion of codes
    Dim ilVff As Integer
    Dim slPostLogSource As String * 1
    Dim blIsItBarterByCount As Boolean
    Dim ilDay As Integer
    Dim llWeekStartDate As Long
    Dim ilHowManyWeeks As Integer
    Dim llAirDate As Long
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    
    Screen.MousePointer = vbHourglass
    'Open all necessary files
    If Not mOpenInvAffFiles() Then
        MsgBox "Some Of Necessary Files Could Not Be Opened "
        mCloseInvAffFiles
        Exit Sub
    End If
    

    'D.S. Start 1-23-02
    'We only want to do this work of we have a combination of local and national spots
    ilLocNatRegFlag = False
    'If (igFoundNatl And igFoundLocl) Then  '3-21-08 always test for combination of local, natl, regional
        ilLocNatRegFlag = True
        ilRet = gObtainSalesperson()
        ReDim tlSSource(0 To 0)
        For ilLoop = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            'tlSSource(ilLoop - 1).iSlfCode = tgMSlf(ilLoop).iCode
            tlSSource(ilLoop).iSlfCode = tgMSlf(ilLoop).iCode
            tmSofSrchKey.iCode = tgMSlf(ilLoop).iSofCode
            ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                tmMnfSrchKey.iCode = tmSof.iMnfSSCode
                ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'tlSSource(ilLoop - 1).iShowLocNatReg = False
                    If RptSelIA!ckcSelC3(0).Value And (tmMnf.iGroupNo = 3) Then
                        'tlSSource(ilLoop - 1).iShowLocNatReg = True
                        tlSSource(ilLoop).iShowLocNatReg = True
                    End If
                    If RptSelIA!ckcSelC3(1).Value And (tmMnf.iGroupNo = 1) Then
                        'tlSSource(ilLoop - 1).iShowLocNatReg = True
                        tlSSource(ilLoop).iShowLocNatReg = True
                    End If
                    If RptSelIA!ckcSelC3(2).Value And (tmMnf.iGroupNo = 2) Then
                        'tlSSource(ilLoop - 1).iShowLocNatReg = True
                        tlSSource(ilLoop).iShowLocNatReg = True
                    End If
                Else
                    MsgBox "Failed to Read from MNF file. "
                    mCloseInvAffFiles
                    Exit Sub
                End If
            Else
                MsgBox "Failed to Read from SOF file. "
                mCloseInvAffFiles
                Exit Sub
            End If
            ReDim Preserve tlSSource(0 To UBound(tlSSource) + 1)
        Next ilLoop
    'End If
    'D.S. End 1-23-02
    '8-23-11 remove to do standard bdcst month; use dates entered due to End of Contract Billing
'    slTemp = RptSelIA!edcSelCFrom.Text
    slTemp = RptSelIA!CSI_CalFrom.Text          '8-27-19 use csi calendar control vs edit box
    llStartStd = gDateValue(slTemp)
    slStartStd = Format$(llStartStd, "m/d/yy")       'ensure year in conversion
   
'    slTemp = RptSelIA!edcSelCFrom1.Text
    slTemp = RptSelIA!CSI_CalTo.Text
    llEndStd = gDateValue(slTemp)
    slEndStd = Format$(llEndStd, "m/d/yy")       'ensure year in conversion
    
    
    'backup to Monday
    llWeekStartDate = llStartStd
    ilDay = gWeekDayLong(llWeekStartDate)
    Do While ilDay <> 0
        llWeekStartDate = llWeekStartDate - 1
        ilDay = gWeekDayLong(llWeekStartDate)
    Loop
    
    If RptSelIA!ckcUseCountAff.Value = vbChecked Then       'using special air time + count aff of performance form
                                                            'send report formula for starting monday to calculate start date of weeks
        slTemp = Format$(llWeekStartDate, "m/d/yy")
        gObtainYearMonthDayStr slTemp, True, slYear, slMonth, slDay
        If Not gSetFormula("MondayStartDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            MsgBox "Invalid Monday Start Date: gInvAffRpt"
            mCloseInvAffFiles
            Exit Sub
        End If
    End If
    
    'determine how many weeks
    ilHowManyWeeks = (llEndStd - llWeekStartDate) / 7 + 1
    
'    'Concatenate user entered month & "/15/" & user entered year
'    slTemp = RptSelIA.edcSelCFrom.Text   'Month user input
'    slBuildDate = mVerifyMonth(slTemp)
'    slBuildDate = slBuildDate & "/15/"
'    slTemp = RptSelIA!edcSelCFrom1.Text  'Year user input
'    slBuildDate = slBuildDate & Trim$(Str$(gVerifyYear(slTemp)))
'
'    'Get standard broadcast month start and end dates based off user input month and year
'    slStartStd = gObtainStartStd(slBuildDate)
'    llStartStd = gDateValue(slStartStd)
'    slEndStd = gObtainEndStd(slBuildDate)
'    llEndStd = gDateValue(slEndStd)
    
    'Btrieve date values for start and end of requested period; no longer has to be a standard bdcst month
    gPackDate slStartStd, ilStartStd(0), ilStartStd(1)
    gPackDate slEndStd, ilEndStd(0), ilEndStd(1)

    'Biuld array of vehicles that are Conv, Selling or Airing
    tmVefSrchKey.iCode = 0
    ilRet = btrGetGreaterOrEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    ilIdx = 0
    Do While (ilRet = BTRV_ERR_NONE)
        'If ((tlVef.sType = "C") Or (tlVef.sType = "S") Or (tlVef.sType = "A")) Or (tlVef.sType = "G") Or ((tlVef.sType = "R") And (tgSpf.sPostCalAff = "D")) Then
        If ((tlVef.sType = "C") Or (tlVef.sType = "S") Or (tlVef.sType = "A")) Or (tlVef.sType = "G") Or ((tlVef.sType = "R") And ((Asc(tgSpf.sUsingFeatures8) And REPBYDT) = REPBYDT)) Then
            tlVehArray(ilIdx).sVehType = tlVef.sType
            tlVehArray(ilIdx).iVehCode = tlVef.iCode
            ilIdx = ilIdx + 1
            ReDim Preserve tlVehArray(0 To ilIdx)
        End If
        ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Loop
    'Biuld array of simulcast vehicles
    tmVefSrchKey.iCode = 0
    ilRet = btrGetGreaterOrEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    ilIdx = 0
    Do While (ilRet = BTRV_ERR_NONE)
        If (tlVef.sType = "T") And (tlVef.sState = "A") Then
            tlsimulcast(ilIdx).iParentVehCode = tlVef.iVefCode
            tlsimulcast(ilIdx).iChildVehCode = tlVef.iCode
            tlsimulcast(ilIdx).sChildVehName = tlVef.sName
            ilIdx = ilIdx + 1
            ReDim Preserve tlsimulcast(0 To ilIdx)
        End If
        ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Loop


    '6-25-14 build array of selections
    ReDim tmSelAdf(0 To 0) As Integer
    If (RptSelIA!edcSelCTo.Text = "") Then
        ReDim tmSelChf(0 To 0) As Long 'We don't want to redim if its a single contract above
    End If
    
    'Build single contract
    ReDim tmSelChf(0 To 0) As Long
    If (RptSelIA!edcSelCTo.Text <> "") Then 'Single Contracts
        tlChfSrchKey1.lCntrNo = Val(RptSelIA!edcSelCTo.Text)
        tlChfSrchKey1.iCntRevNo = 32000
        tlChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(RptSelIA!edcSelCTo.Text)) And (tmChf.sSchStatus <> "F")
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If tmChf.lCntrNo <> Val(RptSelIA!edcSelCTo.Text) Then
            ilRet = 1
        End If
        If (ilRet = BTRV_ERR_NONE) Then
            'determine if cntr OK for user to see
            ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
            If ilOKToSeeCnt Then
                tmSelChf(UBound(tmSelChf)) = tmChf.lCode
                ReDim Preserve tmSelChf(0 To UBound(tmSelChf) + 1) As Long
            End If
        End If
    End If

    If ((RptSelIA!rbcSelCSelect(0).Value) And (RptSelIA!edcSelCTo.Text = "")) Then 'Advertiser/Contracts, no single contract entered
        If Not RptSelIA!ckcAll.Value = vbChecked Then             'selected advts, not all
            RptSelIA!ckcContrFeed(1).Value = vbUnchecked
            RptSelIA!ckcContrFeed(0).Value = vbUnchecked      'default contracts disallowed until one selected found
        End If
        'get list of select advt
'        For ilLoop = 0 To RptSelIA!lbcSelection(5).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(5).Selected(ilLoop) Then
'                slNameCode = tgAdvertiser(ilLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelAdf(UBound(tmSelAdf)) = Val(slCode)
'                ReDim Preserve tmSelAdf(0 To UBound(tmSelAdf) + 1) As Integer
'            End If
'        Next ilLoop
        gObtainCodesForMultipleLists 5, tgAdvertiser(), ilIncludeCodes, tmSelAdf(), RptSelIA

    End If

    If ((RptSelIA!rbcSelCSelect(0).Value) And (RptSelIA!edcSelCTo.Text = "")) Then 'Advertiser/Contracts, no single contract entered, get the contracts selected for the advt
        For ilLoop = 0 To RptSelIA!lbcSelection(0).ListCount - 1 Step 1
            If RptSelIA!lbcSelection(0).Selected(ilLoop) Then
                slNameCode = RptSelIA!lbcCntrCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmSelChf(UBound(tmSelChf)) = Val(slCode)
                If Not RptSelIA!ckcAll = vbChecked Then
                    If Val(slCode) = 0 Then
                        RptSelIA!ckcContrFeed(1).Value = vbChecked      'feed allowed from contract list
                    Else
                        RptSelIA!ckcContrFeed(0).Value = vbChecked      'locals allowed from contract list
                    End If
                End If
                ReDim Preserve tmSelChf(0 To UBound(tmSelChf) + 1) As Long
            End If
        Next ilLoop
    End If
    
    'Build array of agf codes
    ReDim tmSelAgf(0 To 0) As Integer
    If RptSelIA!rbcSelCSelect(1).Value Then 'Agency
'        For ilLoop = 0 To RptSelIA!lbcSelection(1).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(1).Selected(ilLoop) Then
'                slNameCode = tgAgency(ilLoop).sKey '
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelAgf(UBound(tmSelAgf)) = Val(slCode)
'                ReDim Preserve tmSelAgf(0 To UBound(tmSelAgf) + 1) As Integer
'            End If
'        Next ilLoop
        
        gObtainCodesForMultipleLists 1, tgAgency(), ilIncludeCodes, tmSelAgf(), RptSelIA

    End If
    
    'Build Array of Slf codes
    ReDim tmSelSlf(0 To 0) As Integer
    If RptSelIA!rbcSelCSelect(2).Value Then 'Salesperson
'        For ilLoop = 0 To RptSelIA!lbcSelection(2).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(2).Selected(ilLoop) Then
'                slNameCode = tgSalesperson(ilLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelSlf(UBound(tmSelSlf)) = Val(slCode)
'                ReDim Preserve tmSelSlf(0 To UBound(tmSelSlf) + 1) As Integer
'            End If
'        Next ilLoop
        
        gObtainCodesForMultipleLists 2, tgSalesperson(), ilIncludeCodes, tmSelSlf(), RptSelIA

    End If
    
    
    'Biuld PhfRvf array of PHF contract numbers and invoice numbers
'    ilRet = btrGetGreaterOrEqual(hmPhf, tmPhf, imPhfRecLen, tmPhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    ilIdx = 0
'    Do While (ilRet = BTRV_ERR_NONE)
'        If (tmPhf.sTranType = "IN") Or (tmPhf.sTranType = "HI") Then
'            gUnpackDateLong tmPhf.iInvDate(0), tmPhf.iInvDate(1), llInvDate
'            If ((llInvDate >= llStartStd) And (llInvDate <= llEndStd)) Then
'                ilFound = False
'                For ilLoop = 0 To UBound(tlPhfRvf) - 1 Step 1  'Test to see if Inv and Cntr already in array
'                    If (tlPhfRvf(ilIdx).lInvNo = tmPhf.lInvNo) And (tlPhfRvf(ilIdx).lCntrNo = tmPhf.lCntrNo) Then
'                        ilFound = True
'                        Exit For
'                    End If
'                Next ilLoop
'                If Not ilFound Then
'                    tlPhfRvf(ilIdx).lInvNo = tmPhf.lInvNo
'                    tlPhfRvf(ilIdx).lCntrNo = tmPhf.lCntrNo
'                    tlChfSrchKey1.lCntrNo = tlPhfRvf(ilIdx).lCntrNo
'                    tlChfSrchKey1.iCntRevNo = 32000
'                    tlChfSrchKey1.iPropVer = 32000
'                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlPhfRvf(ilIdx).lCntrNo) And (tmChf.sSchStatus <> "F")
'                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                    If (ilRet = BTRV_ERR_NONE) And (tlPhfRvf(ilIdx).lCntrNo = tmChf.lCntrNo) Then
'                        ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
'                        If ilOKToSeeCnt Then
'                            tlPhfRvf(ilIdx).lChfCode = tmChf.lChfCode
'                            ilIdx = ilIdx + 1
'                            ReDim Preserve tlPhfRvf(0 To ilIdx)
'                        End If
'                    Else
'                        'initialize fields so gCntrOKForUser can set the
'                        tmChf.lVefCode = tmPhf.iBillVefCode
'                        tmChf.lCode = 0
'                        tmChf.lCntrNo = tmPhf.lCntrNo
'                        For ilLoop = 0 To 9
'                            tmChf.iSlfCode(ilLoop) = 0
'                        Next ilLoop
'                        tmChf.iSlfCode(0) = tmPhf.iSlfCode
'                        ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
'                        If ilOKToSeeCnt Then
'                            tlPhfRvf(ilIdx).lChfCode = 0
'                            ilIdx = ilIdx + 1
'                            ReDim Preserve tlPhfRvf(0 To ilIdx)
'                        End If
'                    End If
'                    'ilIdx = ilIdx + 1
'                    'ReDim Preserve tlPhfRvf(0 To ilIdx)
'                End If
'            End If
'        End If
'        ilRet = btrGetNext(hmPhf, tmPhf, imPhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'    Loop

    ilIdx = 0
    'Build PhfRvf array of RVF contract numbers and invoice numbers
    'We continue adding to the tlPhfRfv array - we do not reset ilIdx back to zero
    hmRvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRvfRecLen = Len(tmRvf)
    '2 passes, first on PHF (receivables History), then RVF (receivables).  Both files have same structure
    For ilLoopOnFile = 1 To 2
        ilRet = btrGetGreaterOrEqual(hmRvf, tmRvf, imRvfRecLen, tmRvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE)
            If tmRvf.sTranType = "IN" Then
                gUnpackDateLong tmRvf.iInvDate(0), tmRvf.iInvDate(1), llInvDate
                If ((llInvDate >= llStartStd) And (llInvDate <= llEndStd)) Then
                    
                    blSelectionOK = False
                     
                    'test advertisers
                    If RptSelIA!rbcSelCSelect(0).Value = True Then
'                        For ilLoopOnSelection = 0 To UBound(tmSelAdf) - 1 Step 1
'                            If tmRvf.iAdfCode = tmSelAdf(ilLoopOnSelection) Then
'                                blSelectionOK = True
'                                Exit For
'                            End If
'                        Next ilLoopOnSelection
                        If gFilterLists(tmRvf.iAdfCode, ilIncludeCodes, tmSelAdf()) Then
                            blSelectionOK = True
                        End If
                    End If
                     
                     'Test Agency
                    If RptSelIA!rbcSelCSelect(1).Value = True Then
'                        For ilLoopOnSelection = 0 To UBound(tmSelAgf) - 1 Step 1
'                            If tmRvf.iAgfCode = tmSelAgf(ilLoopOnSelection) Then
'                                blSelectionOK = True
'                                Exit For
'                            End If
'                        Next ilLoopOnSelection
                        If gFilterLists(tmRvf.iAgfCode, ilIncludeCodes, tmSelAgf()) Then
                            blSelectionOK = True
                        End If

                    End If
            
                    'Test Salesperson
                    If RptSelIA!rbcSelCSelect(2).Value = True Then
'                        For ilLoopOnSelection = 0 To UBound(tmSelSlf) - 1 Step 1
'                             If tmRvf.iSlfCode = tmSelSlf(ilLoopOnSelection) Then
'                                 blSelectionOK = True
'                                 Exit For
'                             End If
'                        Next ilLoopOnSelection
                        If gFilterLists(tmRvf.iSlfCode, ilIncludeCodes, tmSelSlf()) Then
                            blSelectionOK = True
                        End If

                    End If
                    
                    If blSelectionOK Then
                    
                        ilFound = False
                        For ilLoop = 0 To UBound(tlPhfRvf) - 1 Step 1  'Test to see if Inv and Cntr already in array
                            If (tlPhfRvf(ilLoop).lInvNo = tmRvf.lInvNo) And (tlPhfRvf(ilLoop).lCntrNo = tmRvf.lCntrNo) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilFound Then
        
                            tlPhfRvf(ilIdx).lInvNo = tmRvf.lInvNo
                            tlPhfRvf(ilIdx).lCntrNo = tmRvf.lCntrNo
                            tlChfSrchKey1.lCntrNo = tlPhfRvf(ilIdx).lCntrNo
                            tlChfSrchKey1.iCntRevNo = 32000
                            tlChfSrchKey1.iPropVer = 32000
                            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlPhfRvf(ilIdx).lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")    'test for fully scheduled and manually sched
                                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            If (ilRet = BTRV_ERR_NONE) And (tlPhfRvf(ilIdx).lCntrNo = tmChf.lCntrNo) Then
                                ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
                                If ilOKToSeeCnt Then
                                    tlPhfRvf(ilIdx).lChfCode = tmChf.lCode
                                    ilIdx = ilIdx + 1
                                    ReDim Preserve tlPhfRvf(0 To ilIdx)
                                End If
                            Else
                                'initialize fields so gCntrOKForUser can set the
                                tmChf.lVefCode = tmRvf.iBillVefCode
                                tmChf.lCode = 0
                                tmChf.lCntrNo = tmRvf.lCntrNo
                                For ilLoop = 0 To 9
                                    tmChf.iSlfCode(ilLoop) = 0
                                Next ilLoop
                                tmChf.iSlfCode(0) = tmRvf.iSlfCode
                                ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
                                If ilOKToSeeCnt Then
                                    tlPhfRvf(ilIdx).lChfCode = 0
                                    ilIdx = ilIdx + 1
                                    ReDim Preserve tlPhfRvf(0 To ilIdx)
                                End If
                            End If
                        End If
    
                        'ilIdx = ilIdx + 1
                        'ReDim Preserve tlPhfRvf(0 To ilIdx)
                    End If
                End If
            End If
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Loop
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        hmRvf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    Next ilLoopOnFile
    
'    'Build single contract
'    ReDim tmSelChf(0 To 0) As Long
'    If (RptSelIA!edcSelCTo.Text <> "") Then 'Single Contracts
'        tlChfSrchKey1.lCntrNo = Val(RptSelIA!edcSelCTo.Text)
'        tlChfSrchKey1.iCntRevNo = 32000
'        tlChfSrchKey1.iPropVer = 32000
'        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(RptSelIA!edcSelCTo.Text)) And (tmChf.sSchStatus <> "F")
'            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        If tmChf.lCntrNo <> Val(RptSelIA!edcSelCTo.Text) Then
'            ilRet = 1
'        End If
'        If (ilRet = BTRV_ERR_NONE) Then
'            'determine if cntr OK for user to see
'            ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
'            If ilOKToSeeCnt Then
'                tmSelChf(UBound(tmSelChf)) = tmChf.lCode
'                ReDim Preserve tmSelChf(0 To UBound(tmSelChf) + 1) As Long
'            End If
'        End If
'    End If

    'Build array of contracts
'    ReDim tmSelAdf(0 To 0) As Integer
'    If (RptSelIA!edcSelCTo.Text = "") Then
'        ReDim tmSelChf(0 To 0) As Long 'We don't want to redim if its a single contract above
'    End If
'
'    If ((RptSelIA!rbcSelCSelect(0).Value) And (RptSelIA!edcSelCTo.Text = "")) Then 'Advertiser/Contracts
'        If Not RptSelIA!ckcAll.Value = vbChecked Then             'selected advt
'            RptSelIA!ckcContrFeed(1).Value = vbUnchecked
'            RptSelIA!ckcContrFeed(0).Value = vbUnchecked      'default contracts disallowed until one selected found
'        End If
'        For ilLoop = 0 To RptSelIA!lbcSelection(5).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(5).Selected(ilLoop) Then
'                slNameCode = tgAdvertiser(ilLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelAdf(UBound(tmSelAdf)) = Val(slCode)
'                ReDim Preserve tmSelAdf(0 To UBound(tmSelAdf) + 1) As Integer
'            End If
'        Next ilLoop
'    End If
'
'    If ((RptSelIA!rbcSelCSelect(0).Value) And (RptSelIA!edcSelCTo.Text = "")) Then 'Advertiser/Contracts
'        For ilLoop = 0 To RptSelIA!lbcSelection(0).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(0).Selected(ilLoop) Then
'                slNameCode = RptSelIA!lbcCntrCode.List(ilLoop)
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelChf(UBound(tmSelChf)) = Val(slCode)
'                If Not RptSelIA!ckcAll = vbChecked Then
'                    If Val(slCode) = 0 Then
'                        RptSelIA!ckcContrFeed(1).Value = vbChecked      'feed allowed from contract list
'                    Else
'                        RptSelIA!ckcContrFeed(0).Value = vbChecked      'locals allowed from contract list
'                    End If
'                End If
'                ReDim Preserve tmSelChf(0 To UBound(tmSelChf) + 1) As Long
'            End If
'        Next ilLoop
'    End If
'    'Build array of agf codes
'    ReDim tmSelAgf(0 To 0) As Integer
'    If RptSelIA!rbcSelCSelect(1).Value Then 'Agency
'        For ilLoop = 0 To RptSelIA!lbcSelection(1).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(1).Selected(ilLoop) Then
'                slNameCode = tgAgency(ilLoop).sKey '
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelAgf(UBound(tmSelAgf)) = Val(slCode)
'                ReDim Preserve tmSelAgf(0 To UBound(tmSelAgf) + 1) As Integer
'            End If
'        Next ilLoop
'    End If
'    'Build Array of Slf codes
'    ReDim tmSelSlf(0 To 0) As Integer
'    If RptSelIA!rbcSelCSelect(2).Value Then 'Salesperson
'        For ilLoop = 0 To RptSelIA!lbcSelection(2).ListCount - 1 Step 1
'            If RptSelIA!lbcSelection(2).Selected(ilLoop) Then
'                slNameCode = tgSalesperson(ilLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                tmSelSlf(UBound(tmSelSlf)) = Val(slCode)
'                ReDim Preserve tmSelSlf(0 To UBound(tmSelSlf) + 1) As Integer
'            End If
'        Next ilLoop
'    End If
    
    llNumSdfRecs = 0  'just a counter for test purposes
    
'   The vehicle, advertiser and agency tables have already been populated
'    ilRet = gObtainVef()  'retrieve all the vehicles for the latest Book used
'    If Not gObtainAdvt() Then
'        MsgBox "The Advertiser File Could Not Be Opened "
'        mCloseInvAffFiles
'        Exit Sub
'    End If
'    If Not gObtainAgency() Then
'        MsgBox "The Agency File Could Not Be Opened "
'        mCloseInvAffFiles
'        Exit Sub
'    End If

    'Obtain spots for a given vehicle then test spot info against pre-built chf, agf, slf arrays
    For ilVehIdx = 0 To UBound(tlVehArray) - 1 Step 1
        '5-9-11 Remove all the future bb spots
        ilGameSelect = 0
        ilLineSelect = 0
        llChfSelect = 0
        ilRet = gRemoveBBSpots(hmSdf, tlVehArray(ilVehIdx).iVehCode, ilGameSelect, slStartStd, slEndStd, llChfSelect, ilLineSelect)

        '5-8-14 get last log date for vehicle processing; cannot see open/close bb in the future
        ilVpfIndex = gBinarySearchVpf(tlVehArray(ilVehIdx).iVehCode)
    
        'ignore  BB from days in future
        If ilVpfIndex <> -1 Then
            gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slLLDate
            If slLLDate = "" Then
                slLLDate = Format(Now, "m/d/yy")
            Else
                If gDateValue(slLLDate) < gDateValue(Format(Now, "m/d/yy")) Then
                    slLLDate = Format(Now, "m/d/yy")
                End If
            End If
            slLLDate = gIncOneDay(slLLDate)
        Else
            slLLDate = gIncOneDay(Format(Now, "m/d/yy"))
        End If
        llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater
        
        ilVff = gBinarySearchVff(tlVehArray(ilVehIdx).iVehCode)
        If ilVff = -1 Then
            slPostLogSource = ""
        Else
            slPostLogSource = tgVff(ilVff).sPostLogSource
        End If
        'determine if vehicle is a barter and imported invoice type.  If so, is it imported or not.  If imported, are spot times entered with counts only, or automtically
        'with invoice import feature
        ReDim tmPostBarterCounts(0 To 0) As POSTBARTERCOUNTS
        If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) And (gIsOnInsertions(tlVehArray(ilVehIdx).iVehCode) = True) And slPostLogSource = "S" Then
            'check to see if entry exists in the imported posting file
            mGatherPostBarterCounts tlVehArray(ilVehIdx).iVehCode, llStartStd, llEndStd
        End If
        
            tmSdfSrchKey1.iVefCode = tlVehArray(ilVehIdx).iVehCode
            gPackDate slStartStd, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
            tmSdfSrchKey1.iTime(0) = 0
            tmSdfSrchKey1.iTime(1) = 0
            tmSdfSrchKey1.sSchStatus = ""
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            llSptDate = 0
    
            Do While ((ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tlVehArray(ilVehIdx).iVehCode))
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSptDate
                If llSptDate > llEndStd Then
                    Exit Do
                End If
    
                If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) Then  'Scheduled, Makegood, Outside
                    slShowOnInv = "Y"
                    If tmSdf.sSpotType = "X" Then       '3-23-03 some kind of fill/extra spot, determine from advt if its to show on this Affidavit
                        '1-19-04 change way in which fill/extra are shown.  ONly test advt if the spot price type flag hasnt been overridden
                        slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)
                    End If
                    If slShowOnInv <> "N" Then
                        ilSpotOK = False
    
                        'Test Contracts
                        If gSetCheck(RptSelIA!ckcAll.Value) And (RptSelIA!edcSelCTo.Text = "") Then
                            ilSpotOK = True            'All advertisers were selected so get all spots
                        Else
                            'check for selective advertiser and if feed spot matches
                            If tmSdf.lChfCode = 0 And RptSelIA!rbcSelCSelect(0).Value Then 'Adv
                                If gFilterLists(tmSdf.iAdfCode, ilIncludeCodes, tmSelAdf()) Then
                                    ilSpotOK = True
                                End If
    
                            Else
                                For ilIdx = 0 To UBound(tmSelChf) - 1 Step 1
                                    If tmSdf.lChfCode = tmSelChf(ilIdx) Then
                                        ilSpotOK = True
                                        Exit For
                                    End If
                                Next ilIdx
                            End If
    
                        End If
                        'If Selection was by Agency or Salesperson then we need to get the Spots Chf record
                        If ((UBound(tmSelSlf) > 0) Or (UBound(tmSelAgf) > 0)) Then
                            tmChfSrchKey.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            ilRet = Not BTRV_ERR_NONE
                        End If
    
                        'Test Agency
                        If (ilRet = BTRV_ERR_NONE) Then
                            If gFilterLists(tmChf.iAgfCode, ilIncludeCodes, tmSelAgf()) Then
                                ilSpotOK = True
                            End If
    
                        End If
    
                        'Test Salesperson
                        If (ilRet = BTRV_ERR_NONE) Then
                            If gFilterLists(tmChf.iSlfCode(0), ilIncludeCodes, tmSelSlf()) Then
                                ilSpotOK = True
                            End If
                        End If
    
                        'D.S. Start 1-23-02
                        If tmSdf.lChfCode = 0 Then      'feed spot
                            tmFSFSrchKey.lCode = tmSdf.lFsfCode
                            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            tmIvr.iCTSplit = 0  'Percent of trade
    
                        Else                            'contract spot
                            tmChfSrchKey.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If (ilRet = BTRV_ERR_NONE) Then
                                tmIvr.iCTSplit = tmChf.iPctTrade  'Percent of trade
                            End If
    
                            tmClfSrchKey.lChfCode = tmChf.lCode
                            tmClfSrchKey.iLine = tmSdf.iLineNo
                            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                MsgBox "Pervasive Call To The Contract Line File (CLF) Failed "
                                mCloseInvAffFiles
                                Exit Sub
                            End If
    
                            'check if type of contract to be selected (contract spots = all spot types except psa & promo)
                            'chf.stype :  S = psa, M = promo
    
    
                            If tmChf.sType = "S" Then
                                If RptSelIA!ckcContrFeed(2).Value = vbUnchecked Then    'include psa?
                                    ilSpotOK = False
                                End If
                            ElseIf tmChf.sType = "M" Then
                                If RptSelIA!ckcContrFeed(3).Value = vbUnchecked Then        'include promo
                                    ilSpotOK = False
                                End If
                            ElseIf (RptSelIA!ckcContrFeed(0).Value = vbUnchecked) Then        'include all other contract types
                                ilSpotOK = False
                            End If
    
                            If ilLocNatRegFlag Then
                                For ilIdx = 0 To UBound(tlSSource) - 1 Step 1
                                    If tlSSource(ilIdx).iSlfCode = tmChf.iSlfCode(0) Then
                                        If tlSSource(ilIdx).iShowLocNatReg = False Then
                                            ilSpotOK = False
                                            Exit For
                                        End If
                                    End If
                                Next ilIdx
                            End If
                            'D.S. End 1-23-02
                            
                            'Test if Open or Close BB, ignore if in the future
                            If tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C" Then
                                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                If gDateValue(slDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
                                    ilSpotOK = False
                                End If
                            End If
    
                        End If
    
                        'see if the spot is a Barter spot that is invoice with counts.  If so, ignore it
                        blIsItBarterByCount = mTestSpotForBarterCounts()
                        
                        If ilSpotOK = True And gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode()) Then
                            ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                            tmIvr.sARate = Trim$(slPrice)
                            If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                                'is it a .00?
                                If gCompNumberStr(slPrice, "0.00") = 0 Then       'its a .00 spot
                                    tmIvr.lOTotalGross = 0
                                Else
                                    tmIvr.lOTotalGross = gStrDecToLong(slPrice, 2)
                                End If
                            Else
                                'its a bonus, recap, n/c, etc. which is still $0
                                'tmIvr.lATotalGross = 0
                                tmIvr.lOTotalGross = 0      '2-22-08
                            End If
    
                            tmIvr.iGenDate(0) = igNowDate(0)  'todays date used for removal of records
                            tmIvr.iGenDate(1) = igNowDate(1)
                            'tmIvr.iGenTime(0) = igNowTime(0)  'todays time used for removal of records
                            'tmIvr.iGenTime(1) = igNowTime(1)
                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                            tmIvr.lGenTime = lgNowTime
                            tmIvr.iType = 0                   'type 0 = spots
                            tmIvr.lChfCode = tmSdf.lChfCode         'Contract header code
                            tmIvr.iInvStartDate(0) = ilStartStd(0)  'Start date of requested period
                            tmIvr.iInvStartDate(1) = ilStartStd(1)
                            tmIvr.iInvDate(0) = ilEndStd(0)         'End date of requested period
                            tmIvr.iInvDate(1) = ilEndStd(1)
    
                            tmIvr.lInvNo = 0   'Invoice number
                            imInvNoFound = False
                            If tmSdf.lChfCode > 0 Then                  'invoice #s dont apply to feed spots
                                For ilIdx = 0 To UBound(tlPhfRvf) - 1 Step 1
                                    llTempCntrNum = tlPhfRvf(ilIdx).lCntrNo
                                    If tmChf.lCntrNo = llTempCntrNum Then
                                        tmIvr.lInvNo = tlPhfRvf(ilIdx).lInvNo
                                        imInvNoFound = True
                                        Exit For
                                    End If
                                Next ilIdx
                            End If
                            tmIvr.sSpotType = tmSdf.sSpotType
                            tmIvr.iLen = tmSdf.iLen 'Spot len
                            '5-26-06 if spot is cross-midnight, adjust date to next date (true date it ran on)
                            If tmSdf.sXCrossMidnight = "Y" Then
                                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSptDate
                                gPackDate llSptDate + 1, ilDate(0), ilDate(1)
                                gUnpackDate ilDate(0), ilDate(1), slTemp    'tmIvr.sADayDate 'Air date
                            Else
                                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slTemp  'tmIvr.sADayDate 'Air date
                            End If
                            llAirDate = gDateValue(slTemp)
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", tmIvr.sATime 'Air time
                            
                            mAdjDateTime tmSdf.iVefCode, slTemp, tmIvr.sATime
                            Select Case gWeekDayStr(slTemp)
                            Case 0  'Monday
                                tmIvr.sADayDate = "Mo, " & slTemp
                            Case 1  'Tuesday
                                tmIvr.sADayDate = "Tu, " & slTemp
                            Case 2  'Wednesday
                                tmIvr.sADayDate = "We, " & slTemp
                            Case 3  'Thursday
                                tmIvr.sADayDate = "Th, " & slTemp
                            Case 4  'Friday
                                tmIvr.sADayDate = "Fr, " & slTemp
                            Case 5  'Saturday
                                tmIvr.sADayDate = "Sa, " & slTemp
                            Case 6  'Sunday
                                tmIvr.sADayDate = "Su, " & slTemp
                            End Select
    
                            slName = ""
                            'gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", tmIvr.sATime 'Air time
                            'gObtainVehicleName tmSdf.iVefCode, slName, slType  'Vehicle name
                            ilRet = gBinarySearchVef(tmSdf.iVefCode)
                            If ilRet <> -1 Then
                                slName = Trim$(tgMVef(ilRet).sName)
                            End If
                            tmIvr.sAVehName = Trim$(slName)
                            tmIvr.sOVehName = ""
    
                            tmAdfSrchKey.iCode = tmSdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                MsgBox "Pervasive Call To The Advertiser File (ADF) Failed "
                                mCloseInvAffFiles
                                Exit Sub
                            End If
                            If (tmChf.iPnfBuyer > 0) And (tmSdf.lChfCode <> 0) Then
                                tmPnfSrchKey.iCode = tmChf.iPnfBuyer
                                ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    If Trim$(tmPnf.sBillAddr(0)) <> "" Then
                                        tmAdf.sBillAddr(0) = tmPnf.sBillAddr(0)
                                        tmAdf.sBillAddr(1) = tmPnf.sBillAddr(1)
                                        tmAdf.sBillAddr(2) = tmPnf.sBillAddr(2)
                                    ElseIf Trim$(tmPnf.sCntrAddr(0)) <> "" Then
                                        tmAdf.sBillAddr(0) = tmPnf.sCntrAddr(0)
                                        tmAdf.sBillAddr(1) = tmPnf.sCntrAddr(1)
                                        tmAdf.sBillAddr(2) = tmPnf.sCntrAddr(2)
                                    End If
                                End If
                            End If
    
                            tmIvr.iFormType = 0                 'used for feed spots to double use the chfcode field
                            tmIvr.iWkNo = 0              'week index, retained for invoice by spot counts
                            tmIvr.iRdfCode = 0
                            If blIsItBarterByCount Then
                                tmIvr.iFormType = 1
                                ilIdx = (llAirDate - llWeekStartDate) \ 7 + 1
                                tmIvr.iWkNo = ilIdx
                                tmIvr.iRdfCode = tmClf.iRdfCode
                            End If
                            
                            If tmSdf.lChfCode = 0 Then          'feed spot
                                mCreateFeedSortKey
                                imInvNoFound = False 'reset sort key flag
    
                                'clear out any old ISCI numbers
                                For ilIdx = LBound(tmIvr.sACopy) To UBound(tmIvr.sACopy) Step 1
                                    tmIvr.sACopy(ilIdx) = ""
                                Next
                                mObtainIvrCopy slProduct(), slCopy()
                                mGenFeedHeader
                                tmIvr.iFormType = 99                'used to indicate its a feed spot in crystal
                                tmIvr.lChfCode = tmFsf.lCode
    
                                ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
    
                            Else                                'contract spot
                                mCreateSortKey
                                If tmChf.sType <> "V" Then  'ignore reservation spots
                                    imInvNoFound = False 'reset sort key flag
    
                                    'clear out any old ISCI numbers
                                    For ilIdx = LBound(tmIvr.sACopy) To UBound(tmIvr.sACopy) Step 1
                                        tmIvr.sACopy(ilIdx) = ""
                                    Next
                                    mObtainIvrCopy slProduct(), slCopy()
    
                                    mGenHeader
    
                                    'mResetFieldsForIvr() now tests for markets defined as cluster = "Y"
                                    'If yes it returns false and we bypass the contract
                                    If mResetFieldsForIvr() Then
                                        tmIvr.lCode = 0
                                        ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
                                        'Add Simulcast record if one exists - there could be more than one
                                        For ilIdx = 0 To UBound(tlsimulcast) - 1 Step 1
                                            If tlsimulcast(ilIdx).iParentVehCode = tmSdf.iVefCode Then
                                                tmIvr.sAVehName = Trim$(tlsimulcast(ilIdx).sChildVehName)
                                                'Save the old Market code - if the simulcast is 0 use the parents
                                                iSavePrgEnfCode = tmIvr.iPrgEnfCode
                                                If tmIvr.iPrgEnfCode = 0 Then
                                                    tmIvr.iPrgEnfCode = iSavePrgEnfCode
                                                End If
                                                tmIvr.sARate = ".00"    'zero out rate for simulcast spots
                                                tmIvr.lOTotalGross = ".00"     'zero out rate for simulcast spots
                                                tmIvr.lCode = 0
                                                ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
                                            End If
                                        Next ilIdx
                                    End If
                                    ilSpotOK = False                'reset flag
                                    'llNumSdfRecs = llNumSdfRecs + 1 'just a counter for test purposes
                                End If
                            End If                  'tmSdf.lChfCode = 0
                        End If
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
    Next ilVehIdx
    
    
    Erase tlVehArray
    Erase tlsimulcast
    Erase tlPhfRvf
    Erase slProduct
    Erase slCopy
    Erase tmSelChf, tmSelAgf, tmSelSlf, tmSelAdf
    
    Screen.MousePointer = vbDefault
    mCloseInvAffFiles
    Exit Sub
End Sub
'********************************************************
'*                                                      *
'*      Procedure Name:mCloseInvAffFiles                 *
'*                                                      *
'*             Created:2/21/00       By:D. Smith        *
'*             Modified:             By:                *
'*                                                      *
'*            Comments: Close all of the files          *
'*                      for the gInvAffRpt procedure    *
'********************************************************
'
Sub mCloseInvAffFiles()
Dim ilRet As Integer
    Erase tmSelChf
    Erase tmSelAgf
    Erase tmSelSlf
    Erase tmSelAdf

    ilRet = btrClose(hmIvr)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmPhf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmTzf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmPnf)
    ilRet = btrClose(hmArf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmFnf)
    ilRet = btrClose(hmIihf)
    btrDestroy hmIvr
    btrDestroy hmVsf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmPhf
    btrDestroy hmRvf
    btrDestroy hmClf
    btrDestroy hmMnf
    btrDestroy hmCif
    btrDestroy hmCpf
    btrDestroy hmTzf
    btrDestroy hmAgf
    btrDestroy hmArf
    btrDestroy hmPnf
    btrDestroy hmAdf
    btrDestroy hmFsf
    btrDestroy hmAnf
    btrDestroy hmPrf
    btrDestroy hmFnf
    btrDestroy hmIihf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateSortKey                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:2/21/00       By:D. Smith       *
'*                                                     *
'*            Comments: Create the Sort Key (Ivr.sKey) *
'*                      for the Inv. Aff. report       *
'*******************************************************
'
Sub mCreateSortKey()
'
'   Where
'       ilType(I)-0=Sdf (tmSdf); 1=Smf; 2=Psf
'
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim slSort As String
    Dim slStr As String
    Dim slKey As String
    Dim slDate As String
    Dim llTime As Long
    Dim ilFound As Integer
    Dim ilSlf As Integer
    Dim ilBonusSpot As Integer
    Dim tlVef As VEF
    'ReDim slCopyProduct(1 To 6) As String
    ReDim slCopyProduct(0 To 6) As String
    'ReDim slCopy(1 To 6) As String
    ReDim slCopy(0 To 6) As String
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer

    If tmChf.lCode <> tmSdf.lChfCode Then
        tmChfSrchKey.lCode = tmSdf.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    'Bypass Reservation Contracts
    If tmChf.sType = "V" Then
        Exit Sub
    End If
    If tgUrf(0).iSlfCode > 0 Then   'Test if Salesperson- then only allow his contracts
        ilFound = False
        For ilSlf = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
            If tmChf.iSlfCode(ilSlf) <> 0 Then
                If tgUrf(0).iSlfCode = tmChf.iSlfCode(ilSlf) Then
                    ilFound = True
                    Exit For
                End If
            End If
        Next ilSlf
        If Not ilFound Then
            Exit Sub
        End If
    End If
    If (tmClf.lChfCode <> tmChf.lCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
        tmClfSrchKey.lChfCode = tmChf.lCode
        tmClfSrchKey.iLine = tmSdf.iLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        If (tmClf.lChfCode <> tmChf.lCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
            Exit Sub
        End If
    End If
    slName = Format$(tmIvr.lInvNo)
    Do While Len(slName) < 10
        slName = slName & " "
    Loop
    If tmChf.iAgfCode = 0 Then  'Obtain advertiser- direct bill
        'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
        '    If tmChf.iAdfCode = tgCommAdf(ilLoop).iCode Then
            ilLoop = gBinarySearchAdf(tmChf.iAdfCode)
            If ilLoop <> -1 Then
                'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                '    slName = slName & Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                'Else
                    slName = slName & Trim$(tgCommAdf(ilLoop).sName)
                'End If
                tmMnfSrchKey.iCode = tgCommAdf(ilLoop).iMnfSort
        '        Exit For
            End If
        'Next ilLoop
    Else    'Obtain agency
        'For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
        '    If tmChf.iAgfCode = tgCommAgf(ilLoop).iCode Then
            ilLoop = gBinarySearchAgf(tmChf.iAgfCode)
            If ilLoop <> -1 Then
                slName = slName & Trim$(tgCommAgf(ilLoop).sName) & "/" & Trim$(tgCommAgf(ilLoop).sCityID)
                tmMnfSrchKey.iCode = tgCommAgf(ilLoop).iMnfSort
        '        Exit For
            End If
        'Next ilLoop
        If tmMnfSrchKey.iCode <= 0 Then
            'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tmChf.iAdfCode = tgCommAdf(ilLoop).iCode Then
                ilLoop = gBinarySearchAdf(tmChf.iAdfCode)
                If ilLoop <> -1 Then
                    tmMnfSrchKey.iCode = tgCommAdf(ilLoop).iMnfSort
            '        Exit For
                End If
            'Next ilLoop
        End If
    End If
    If tmMnfSrchKey.iCode > 0 Then
        If tmMnf.iCode <> tmMnfSrchKey.iCode Then
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.iCode = 0
            tmMnf.iGroupNo = 0
            'Sort PSA and Promos to end of list
            If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                slSort = "99999"
            Else
                slSort = "99998"
            End If
        Else
            If tmMnf.iGroupNo > 0 Then
                slSort = Trim$(str$(tmMnf.iGroupNo))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            Else
                'Sort PSA and Promos to end of list
                If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                    slSort = "99999"
                Else
                    slSort = "99998"
                End If
            End If
        End If
    Else
        'Sort PSA and Promos to end of list
        If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
            slSort = "99999"
        Else
            slSort = "99998"
        End If
    End If
    Do While Len(slName) < 46
        slName = slName & " "
    Loop
    slKey = slSort & slName
    'Advertiser
    'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
    '    If tmChf.iAdfCode = tgCommAdf(ilLoop).iCode Then
        ilLoop = gBinarySearchAdf(tmChf.iAdfCode)
        If ilLoop <> -1 Then
            ''slKey = slKey & tgCommAdf(ilLoop).sName
            'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
            '    slKey = slKey & Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
            'Else
                slKey = slKey & Trim$(tgCommAdf(ilLoop).sName)
            'End If
    '        Exit For
        End If
    'Next ilLoop
    'Contract number
    slStr = Trim$(str$(tmChf.lCntrNo))
    Do While Len(slStr) < 8
        slStr = "0" & slStr
    Loop
    slKey = slKey & slStr
    'Either contract product or copy product
    ilBonusSpot = False
    If tmSdf.sSpotType = "X" Then
          ilBonusSpot = True
    End If

    If tmChf.sInvGp = "P" Then  'By product-code later (for now use contract product)

        mObtainIvrCopy slCopyProduct(), slCopy()
        If Trim$(slCopyProduct(1)) <> "" Then
            slKey = slKey & slCopyProduct(1)
        Else
            'If copy missing or product missing-> make product ~ plus blanks
            slCopyProduct(1) = "~     "
             slKey = slKey & slCopyProduct(1) 'tmChf.sProduct
        End If

    ElseIf tmChf.sInvGp = "T" Then  'By tag- code later (for now use contract product)
        'If copy missing or tag missing-> make product ~ plus blanks
        slKey = slKey & tmChf.sProduct
    Else
        slKey = slKey & tmChf.sProduct
    End If
    'Spot Vehicle name
    If Not ilBonusSpot Then
        If tmVef.iCode = tmSdf.iVefCode Then
             tlVef = tmVef
        Else
            tmVefSrchKey.iCode = tmSdf.iVefCode
            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tlVef.sName = "Vehicle Missing"
                tlVef.iCode = tmSdf.iVefCode
            End If
        End If
    Else
        If tmVef.iCode = tmSdf.iVefCode Then
            tlVef = tmVef
        Else
            tmVefSrchKey.iCode = tmSdf.iVefCode
            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tlVef.sName = "Vehicle Missing"
                tlVef.iCode = tmSdf.iVefCode
            End If
        End If
    End If
    slKey = slKey & "0"
    slKey = slKey & tlVef.sName
    'Adjust date to true air date if its a spot that crossed midnight
    If tmSdf.sXCrossMidnight = "Y" Then
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
        gPackDate llDate + 1, ilDate(0), ilDate(1)
        gUnpackDateForSort ilDate(0), ilDate(1), slDate
    Else
        gUnpackDateForSort tmSdf.iDate(0), tmSdf.iDate(1), slDate
    End If
    slKey = slKey & slDate
    'Time
    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
    slStr = Trim$(str$(llTime))
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    'tgSort(ilUpper).sKey = slKey & slStr
    tmIvr.sKey = slKey & slStr
    slStr = tmIvr.sKey
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenHeader                      *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:2/23/00       By:D. Smith       *
'*                                                     *
'*            Comments:Generate Header fields          *
'*            for Network/Feed spots                   *
'*
'   10-18-07 Always show Affidavit of Performance vs
'           Invoice and Affidavit
'*******************************************************
'
Sub mGenFeedHeader()
    Dim slStr As String
    Dim slParse As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    If (tgSpf.sBLaserForm = "2") Then
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        tmIvr.sCashTrade = "C"
    ElseIf (tgSpf.sBLaserForm = "3") Then
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        'slStr = "INVOICE" & Chr(10) & "and" & Chr(10) & "AFFIDAVIT"
        tmIvr.sCashTrade = "C"
    Else
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        'slStr = "INVOICE" & Chr(10) & "and" & Chr(10) & "AFFIDAVIT"
        tmIvr.sCashTrade = "C"
    End If

    For ilLoop = LBound(tmIvr.sTitle) To UBound(tmIvr.sTitle) Step 1
        tmIvr.sTitle(ilLoop) = ""
        ilRet = gParseItem(slStr, ilLoop + 1, Chr(10), slParse)
        If (ilRet = CP_MSG_NONE) Then
            tmIvr.sTitle(ilLoop) = slParse
        End If
    Next ilLoop

    tmFnfSrchKey.iCode = tmFsf.iFnfCode     'feed name search
    ilRet = btrGetEqual(hmFnf, tmFnf, imFnfRecLen, tmFnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        tmFnf.iNetArfCode = 0
    End If


    If tmFnf.iNetArfCode > 0 Then
        If tmArf.iCode <> tmFnf.iNetArfCode Then
            tmArfSrchKey.iCode = tmFnf.iNetArfCode
            ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tmArf.sNmAd(0) = "Missing"
                tmArf.sNmAd(1) = " "
                tmArf.sNmAd(2) = " "
                tmArf.sNmAd(3) = " "
            End If
        End If

        'tmIvr.sAddr(1) = Trim$(tmArf.sName)
        'tmIvr.sAddr(2) = Trim$(tmArf.sNmAd(0))
        'tmIvr.sAddr(3) = Trim$(tmArf.sNmAd(1))
        'tmIvr.sAddr(4) = Trim$(tmArf.sNmAd(2))
        'tmIvr.sAddr(5) = Trim$(tmArf.sNmAd(3))
        tmIvr.sAddr(0) = Trim$(tmArf.sName)
        tmIvr.sAddr(1) = Trim$(tmArf.sNmAd(0))
        tmIvr.sAddr(2) = Trim$(tmArf.sNmAd(1))
        tmIvr.sAddr(3) = Trim$(tmArf.sNmAd(2))
        tmIvr.sAddr(4) = Trim$(tmArf.sNmAd(3))
    Else
        'tmIvr.sPayAddr(1) = Trim$(tgSpf.sBPayName)
        'tmIvr.sPayAddr(2) = Trim$(tgSpf.sBPayAddr(0))
        'tmIvr.sPayAddr(3) = Trim$(tgSpf.sBPayAddr(1))
        'tmIvr.sPayAddr(4) = Trim$(tgSpf.sBPayAddr(2))
        tmIvr.sPayAddr(0) = Trim$(tgSpf.sBPayName)
        tmIvr.sPayAddr(1) = Trim$(tgSpf.sBPayAddr(0))
        tmIvr.sPayAddr(2) = Trim$(tgSpf.sBPayAddr(1))
        tmIvr.sPayAddr(3) = Trim$(tgSpf.sBPayAddr(2))
    End If
    tmIvr.iMnfSort = tmAdf.iMnfSort
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainIvrCopy                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:2/21/00       By:D. Smith       *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Sub mObtainIvrCopy(slProduct() As String, slCopy() As String)
'
'   mObtainIvrCopy
'       Where:
'           tmSdf(I)- Spot record
'           tmSmf(I)
'           tmCpf(O)- Product/ISCI record
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilCopy As Integer
    Dim slPtType As String
    Dim llCopyCode As Long


    'For ilIndex = LBound(slCopy) To UBound(slCopy) Step 1
    For ilIndex = LBONE To UBound(slCopy) Step 1
        slCopy(ilIndex) = ""
        slProduct(ilIndex) = ""
        If ilIndex <= 4 Then            'only array of 4 for copy, using the commnet pointers since unused in this report
            tmIvr.lComment(ilIndex - 1) = 0
        End If
    Next ilIndex

    If tmSdf.sSpotType = "X" Then
        If (tmSdf.sSchStatus <> "S") And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
            Exit Sub
        End If
        slPtType = tmSdf.sPtType
        llCopyCode = tmSdf.lCopyCode
    Else
        If (tmSdf.sSchStatus <> "S") And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
            Exit Sub
        End If
        slPtType = tmSdf.sPtType
        llCopyCode = tmSdf.lCopyCode
    End If
    ilCopy = 1
    tmIvr.lClfCxfCode = 0           '4-25-17 init pointer to copy script if requested
    If slPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = llCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.lcpfCode > 0 Then
                tmCpfSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmCpf.sISCI = ""
                    tmCpf.sName = ""
                End If
                slCopy(ilCopy) = Trim$(tmCpf.sISCI)
                slProduct(ilCopy) = Trim$(tmCpf.sName)
                'tmIvr.sACopy(1) = slCopy(ilCopy)
                tmIvr.sACopy(0) = slCopy(ilCopy)
                'tmIvr.lComment(1) = tmCpf.lCode
                tmIvr.lComment(0) = tmCpf.lCode
                                If RptSelIA!ckcShowScript.Value = vbChecked Then                '4-25-17 show script?  if not, do not set up the pointer
                                        tmIvr.lClfCxfCode = tmCif.lCsfCode
                                End If
            End If
        End If
    ElseIf slPtType = "2" Then  '  Combo Copy
    ElseIf slPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = llCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ' Look for the first positive lZone value
            For ilIndex = 1 To 6 Step 1
                If (tmTzf.lCifZone(ilIndex - 1) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), "Oth", 1) <> 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmCpf.sISCI = ""
                                tmCpf.sName = ""
                            End If
                            If Trim$(tmCpf.sISCI) <> "" Then
                                slCopy(ilCopy) = Trim$(tmTzf.sZone(ilIndex - 1)) & ": " & Trim$(tmCpf.sISCI)
                                slProduct(ilCopy) = Trim$(tmCpf.sName)
                                If ilCopy <= 4 Then                           '1-22-10 move from below, tz copy was not showing
                                    tmIvr.sACopy(ilCopy - 1) = slCopy(ilCopy)     'there is only an arry of 4
                                    tmIvr.lComment(ilCopy - 1) = tmCpf.lCode      'send pointer to crystal for creative title output
                                End If
                                ilCopy = ilCopy + 1
                            End If
                            'If ilCopy <= 4 Then
                            '    tmIvr.sACopy(ilCopy) = slCopy(ilCopy)
                            'End If
                        End If
                    End If
                End If
            Next ilIndex
            For ilIndex = 1 To 6 Step 1
                If (tmTzf.lCifZone(ilIndex - 1) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), "Oth", 1) = 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmCpf.sISCI = ""
                                tmCpf.sName = ""
                            End If
                            If Trim$(tmCpf.sISCI) <> "" Then
                                slCopy(ilCopy) = "Other: " & Trim$(tmCpf.sISCI)
                                slProduct(ilCopy) = Trim$(tmCpf.sName)
                                If ilCopy <= 4 Then                         '1-22-10 move from below, tz copy was not showing
                                    tmIvr.sACopy(ilCopy - 1) = slCopy(ilCopy) 'there is only an array of 4
                                    tmIvr.lComment(ilCopy - 1) = tmCpf.lCode  'send pointer to crystal for creative title output
                                End If
                                ilCopy = ilCopy + 1
                            End If
                            'If ilCopy <= 4 Then
                            '    tmIvr.sACopy(ilCopy) = slCopy(ilCopy)
                            'End If
                        End If
                    End If
                End If
            Next ilIndex
        End If
    End If
    If RptSelIA!ckcInclCreativeTitle.Value = vbUnchecked Then           '1-22-10 new option to show creative title
        For ilCopy = 1 To 4
            tmIvr.lComment(ilCopy - 1) = 0        'dont show creative titles
        Next ilCopy
    End If
End Sub
'********************************************************
'*                                                      *
'*      Procedure Name:mOpenInvAffFiles                 *
'*                                                      *
'*             Created:2/21/00       By:D. Smith        *
'*             Modified:             By:                *
'*                                                      *
'*            Comments: Open all of the necessary files *
'*                      for the gInvAffRpt procedure    *
'********************************************************
'
'
Function mOpenInvAffFiles() As Integer
    Dim ilRet As Integer
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)

    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imVefRecLen = Len(tmVef)

    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)

    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)

    hmIvr = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmIvr, "", sgDBPath & "Ivr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imIvrRecLen = Len(tmIvr)
    hmPhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imPhfRecLen = Len(tmPhf)
    hmRvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imClfRecLen = Len(tmClf)
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)

    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imCpfRecLen = Len(tmCpf)

    hmTzf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imTzfRecLen = Len(tmTzf)
    hmPnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imPnfRecLen = Len(tmPnf)
    hmArf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmArf, "", sgDBPath & "Arf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imArfRecLen = Len(tmArf)
    hmAgf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imAgfRecLen = Len(tmAgf)
    mOpenInvAffFiles = True
    mOpenInvAffFiles = True
    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)

    'D.S. Start 1-23-02
    hmSlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imSlfRecLen = Len(tmSlf)

    hmSof = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imSofRecLen = Len(tmSof)

    hmFsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imFsfRecLen = Len(tmFsf)

    hmAnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)

    hmPrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPrf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imPrfRecLen = Len(tmPrf)

    hmFnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFnf
        btrDestroy hmPrf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imFnfRecLen = Len(tmFnf)

    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmFnf
        btrDestroy hmPrf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imCffRecLen = Len(tmCff)

    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmFnf
        btrDestroy hmPrf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    'D.S. End 1-23-02

    hmIihf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIihf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmArf)
        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmPhf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmIihf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmFnf
        btrDestroy hmPrf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmAdf
        btrDestroy hmAgf
        btrDestroy hmArf
        btrDestroy hmPnf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmRvf
        btrDestroy hmPhf
        btrDestroy hmIvr
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCHF

        Screen.MousePointer = vbDefault
        mOpenInvAffFiles = False
        Exit Function
    End If
    imIihfRecLen = Len(tmIihf)
    
    mOpenInvAffFiles = True
End Function
'***********************************************************
'*                                                         *
'*      Procedure Name:mResetFieldsForIvr                  *
'*                                                         *
'*             Created:6/16/93       By:D. LeVine          *
'*            Modified:              By:                   *
'*                                                         *
'*            Comments: Used to get the IvrPrgEnfCode      *
'*                      and test for clusters              *
'*                                                         *
'*            D.S. 2/4/02 Added test for clusters = "Y"    *
'*                                                         *
'***********************************************************
'
Function mResetFieldsForIvr() As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilMnf As Integer
    Dim ilVef As Integer

    mResetFieldsForIvr = True
    If ((tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "A")) And (tgSpf.sBLaserForm = "2") Then
        tmIvr.iType = 0
        tmIvr.iPrgEnfCode = 0
        slStr = Trim$(tmIvr.sAVehName)
        'If market name exist, then add the first 10 characters of market name to vehicle name (combination will only be 40 characters)
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilVef).sName), slStr, 1) = 0 Then
                For ilMnf = LBound(tgMkMnf) To UBound(tgMkMnf) - 1 Step 1
                    If tgMVef(ilVef).iMnfVehGp3Mkt = tgMkMnf(ilMnf).iCode Then
                        tmIvr.iPrgEnfCode = tgMkMnf(ilMnf).iCode
                        'D.S. 2/4/02
                        If (Trim$(tgMkMnf(ilMnf).sRPU)) = "Y" Then
                            mResetFieldsForIvr = False
                        End If
                        Exit For
                    End If
                Next ilMnf
                Exit For
            End If
        Next ilVef
        'tmIvr.sKey = slFields(1) & slFields(2) & slFields(3) & slFields(4) & "0" & slFields(9) & slFields(10) & slFields(11)
    'ElseIf tmSpf.sBLaserForm = "3" Then
    ElseIf tgSpf.sBLaserForm = "3" Then

        'mSetPrgEnfCode ilSdfIndex
        ilPos = InStr(1, tmIvr.sRRemark, "Missed, MG")
        If ilPos = 1 Then
            'tmIvr.sKey = slFields(1) & slFields(2) & slFields(3) & slFields(4) & slFields(8) & slFields(5) & slFields(10) & slFields(11)
        Else
            'tmIvr.sKey = slFields(1) & slFields(2) & slFields(3) & slFields(4) & slFields(8) & slFields(9) & slFields(10) & slFields(11)
        End If
        'If (ilBonusSpot) Or (tmSdf.sSpotType = "X") Then
        '    tmIvr.iType = 1
        'Else
            tmIvr.iType = 0
        'End If
        'Change sort veheicle if Missed, MG type spot
    Else
        tmIvr.iPrgEnfCode = 0
        'tmIvr.sKey = slFields(1) & slFields(2) & slFields(3) & slFields(4) & slFields(5) & slFields(6) & slFields(7) & slFields(8) & slFields(9) & slFields(10) & slFields(11)
        'If (ilBonusSpot) Or (tmSdf.sSpotType = "X") Then
        '    tmIvr.iType = 1
        'Else
            tmIvr.iType = 0
        'End If
    End If
    'If ((tgSpf.sInvAirOrder <> "2") And (tgSpf.sInvAirOrder <> "A")) Or ((tgSpf.sInvAirOrder = "2") And (tgSort(ilSdfIndex).iType = 0) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sInvAirOrder = "2") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Or ((tgSpf.sInvAirOrder = "A") And ((tgSort(ilSdfIndex).iType = 0) Or (tgSort(ilSdfIndex).iType = 1)) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sInvAirOrder = "A") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Then
    'If ((tgSpf.sBLaserForm <> "2") And (tgSpf.sBLaserForm <> "3")) Or ((tgSpf.sBLaserForm = "2") And (tgSort(ilSdfIndex).iType = 0) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sBLaserForm = "2") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Or ((tgSpf.sBLaserForm = "3") And ((tgSort(ilSdfIndex).iType = 0) Or (tgSort(ilSdfIndex).iType = 1)) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sBLaserForm = "3") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Then
    'If ((tgSpf.sBLaserForm <> "2") And (tgSpf.sBLaserForm <> "3")) Or ((tgSpf.sBLaserForm = "2") And (tmIvr.iType = 0) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sBLaserForm = "2") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Or ((tgSpf.sBLaserForm = "3") And ((tmIvr.iType = 0) Or (tmIvr.iType = 1)) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O"))) Or ((tgSpf.sBLaserForm = "3") And ((ilBonusSpot) Or (tmSdf.sSpotType = "X"))) Then
        'If (tgSpf.sBLaserForm = "2") Then
            'If rbcType(3).Value Then
            '    tmIvr.lSpotKeyNo = tmIvr.lSpotKeyNo + 1
            '    ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
            '    ilAnyIvr = True
            'End If
        'Else
            'tmIvr.lSpotKeyNo = tmIvr.lSpotKeyNo + 1
            'ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
            'ilAnyIvr = True
        'End If
    'End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetPrgEnfCode                  *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Comments: Get Program Event Name Code    *
'*                                                     *
'*******************************************************
Sub mSetPrgEnfCode(ilSdfIndex As Integer)
    Dim llTime As Long
    'Dim slType As String
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim slStartTime As String
    Dim slStr As String
    Dim slSTime As String
    Dim slETime As String
    Dim slXMid As String

    tmIvr.iPrgEnfCode = 0
    If (tgSpf.sBLaserForm = "3") And (tmIvr.iType = 0) Then
        If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
            '11/24/12
            'slType = "O"
            'ilType = 0
            ilType = tmSdf.iGameNo
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = tmSdf.iVefCode  'ilVefCode
            tmSsfSrchKey.iDate(0) = tmSdf.iDate(0)  'ilLogDate0
            tmSsfSrchKey.iDate(1) = tmSdf.iDate(1)  'ilLogDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            If (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> tmSdf.iVefCode) Or (tmSsf.iDate(0) <> tmSdf.iDate(0)) Or (tmSsf.iDate(1) <> tmSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Else
                ilRet = BTRV_ERR_NONE
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = tmSdf.iDate(0)) And (tmSsf.iDate(1) = tmSdf.iDate(1))
                For ilLoop = 1 To tmSsf.iCount Step 1
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
                    If ((tmProg.iRecType And &HF) = 1) Then
                        gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llSTime
                        gUnpackTimeLong tmProg.iEndTime(0), tmProg.iEndTime(1), True, llETime
                        If (llTime >= llSTime) And (llTime < llETime) Then
                            gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slStartTime
                            tmLefSrchKey.lLvfCode = tmProg.lLvfCode
                            tmLefSrchKey.iStartTime(0) = 0
                            tmLefSrchKey.iStartTime(1) = 0
                            tmLefSrchKey.iSeqNo = 0
                            ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = tmProg.lLvfCode)
                                If tmLef.iEtfCode = 1 Then
                                    gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
                                    gAddTimeLength slStartTime, slStr, "A", "1", slSTime, slXMid
                                    llSTime = gTimeToLong(slSTime, False)
                                    gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", False, slStr
                                    gAddTimeLength slSTime, slStr, "A", "1", slETime, slXMid
                                    llETime = gTimeToLong(slETime, True)
                                    If (llTime >= llSTime) And (llTime < llETime) Then
                                        tmIvr.iPrgEnfCode = tmLef.iEnfCode
                                    End If
                                End If
                                ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            Exit Do
                        End If
                    End If
                Next ilLoop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
End Sub

'****************************************************************
'*
'*      Procedure Name:mCreatFeedSortKey
'*
'*             Created:3/01/94       By:D. LeVine
'*            Modified:2/21/00       By:D. Smith
'*
'*            Comments: Create the Sort Key (Ivr.sKey) for FEED
'*                     spots for the Inv. Aff. report
'***************************************************************
'
Public Sub mCreateFeedSortKey()
'
'   Where
'       ilType(I)-0=Sdf (tmSdf); 1=Smf; 2=Psf
'
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim slSort As String
    Dim slStr As String
    Dim slKey As String
    Dim slDate As String
    Dim llTime As Long
    Dim ilBonusSpot As Integer
    Dim tlVef As VEF
    Dim slProduct As String * 35

    slName = Format$(tmIvr.lInvNo)
    Do While Len(slName) < 10
        slName = slName & " "
    Loop
    ilLoop = gBinarySearchAdf(tmFsf.iAdfCode)
    If ilLoop <> -1 Then
        ''slName = slName & Trim$(tgCommAdf(ilLoop).sName)
        'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
        '    slName = slName & Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
        'Else
            slName = slName & Trim$(tgCommAdf(ilLoop).sName)
        'End If
        tmMnfSrchKey.iCode = tgCommAdf(ilLoop).iMnfSort
    End If


    If tmMnfSrchKey.iCode > 0 Then
        If tmMnf.iCode <> tmMnfSrchKey.iCode Then
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.iCode = 0
            tmMnf.iGroupNo = 0
            'Sort PSA and Promos to end of list
            'If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
            '    slSort = "99999"
            'Else
                slSort = "99998"
            'End If
        Else
            If tmMnf.iGroupNo > 0 Then
                slSort = Trim$(str$(tmMnf.iGroupNo))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            Else
                slSort = "99998"
            End If
        End If
    Else
        slSort = "99998"
    End If
    Do While Len(slName) < 46
        slName = slName & " "
    Loop
    slKey = slSort & slName
    'Advertiser

        ilLoop = gBinarySearchAdf(tmFsf.iAdfCode)
        If ilLoop <> -1 Then
            ''slKey = slKey & tgCommAdf(ilLoop).sName
            'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
            '    slKey = slKey & Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
            'Else
                slKey = slKey & Trim$(tgCommAdf(ilLoop).sName)
            'End If
        End If
    'Contract number
    slStr = Trim$(tmFsf.sRefID)
    Do While Len(slStr) < 8
        slStr = "0" & slStr
    Loop
    slKey = slKey & slStr
    'Either contract product or copy product
    ilBonusSpot = False
    If tmSdf.sSpotType = "X" Then
          ilBonusSpot = True
    End If

    If tmFsf.lPrfCode = 0 Then
        Do While Len(slProduct) < 35
            slProduct = " " & slProduct
        Loop
        slKey = slKey & slProduct    '35 blanks
    Else
        tmPrfSrchKey.lCode = tmFsf.lPrfCode
        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Do While Len(slProduct) < 35
                slProduct = " " & slProduct
            Loop
            slKey = slKey & slProduct
        Else
            slKey = slKey & tmPrf.sName
        End If
    End If

    'Spot Vehicle name
    If Not ilBonusSpot Then
        If tmVef.iCode = tmSdf.iVefCode Then
             tlVef = tmVef
        Else
            tmVefSrchKey.iCode = tmSdf.iVefCode
            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tlVef.sName = "Vehicle Missing"
                tlVef.iCode = tmSdf.iVefCode
            End If
        End If
    Else
        If tmVef.iCode = tmSdf.iVefCode Then
            tlVef = tmVef
        Else
            tmVefSrchKey.iCode = tmSdf.iVefCode
            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tlVef.sName = "Vehicle Missing"
                tlVef.iCode = tmSdf.iVefCode
            End If
        End If
    End If
    slKey = slKey & "0"
    slKey = slKey & tlVef.sName
    'Date
    gUnpackDateForSort tmSdf.iDate(0), tmSdf.iDate(1), slDate
    slKey = slKey & slDate
    'Time
    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
    slStr = Trim$(str$(llTime))
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    'tgSort(ilUpper).sKey = slKey & slStr
    tmIvr.sKey = slKey & slStr
    slStr = tmIvr.sKey
End Sub

'*************************************************************
'*
'*      Procedure Name:mGenHeader
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:2/23/00       By:D. Smith
'*
'*            Comments:Generate Header fields for contract spot
'*
'   10-18-07 Always show Affidavit of Performance vs
'           Invoice and Affidavit
'*************************************************************
'
Public Sub mGenHeader()
 Dim slStr As String
    Dim slParse As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilAttnCode As Integer
    Dim tlagf As AGF
    If (tgSpf.sBLaserForm = "2") Then
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        tmIvr.sCashTrade = ""               'not applicable for feed spot
    ElseIf (tgSpf.sBLaserForm = "3") Then
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        'slStr = "INVOICE" & Chr(10) & "and" & Chr(10) & "AFFIDAVIT"
        tmIvr.sCashTrade = ""
    Else
        slStr = "AFFIDAVIT" & Chr(10) & "of" & Chr(10) & "PERFORMANCE"
        'slStr = "INVOICE" & Chr(10) & "and" & Chr(10) & "AFFIDAVIT"
        tmIvr.sCashTrade = ""
    End If

    Select Case tmChf.sType
        Case "P"    'Proposal
        Case "H"    'Hold
        Case "V"    'Reservation
        Case "J"    'Rejection
        Case "E"    'Order
        Case "C"    'Contract
        Case "D"    'Deferred
        Case "T"    'Remnant
            'slStr = "REMNANT" & Chr(10) & slStr
        Case "R"    'Direct Response
            slStr = "DIRECT RESPONSE" & Chr(10) & slStr
        Case "Q"    'Direct Response
            slStr = "PER INQUIRY" & Chr(10) & slStr
        Case "S"    'PSA
            slStr = "PSA" & Chr(10) & slStr
        Case "M"    'Promo
            slStr = "PROMO" & Chr(10) & slStr
    End Select

    For ilLoop = LBound(tmIvr.sTitle) To UBound(tmIvr.sTitle) Step 1
        tmIvr.sTitle(ilLoop) = ""
        ilRet = gParseItem(slStr, ilLoop + 1, Chr(10), slParse)
        If (ilRet = CP_MSG_NONE) Then
            tmIvr.sTitle(ilLoop) = slParse
        End If
    Next ilLoop


    If tmChf.iAgfCode > 0 Then
        tmAgfSrchKey.iCode = tmChf.iAgfCode
        ilRet = btrGetEqual(hmAgf, tlagf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            'error statement here
            tmAgfSrchKey.iCode = tmAgfSrchKey.iCode
        End If

        If tgSpf.sBLaserForm = "2" Then
            'If Trim$(tmChf.sBuyer) <> "" Then
            '    ilAttnCode = 1
            '    tmPnf.sName = tmChf.sBuyer
            'Else
            '    ilAttnCode = 0
            'End If
            ilAttnCode = tlagf.iPnfPay
        Else
            ilAttnCode = tlagf.iPnfPay
        End If
        If (tmChf.iPnfBuyer > 0) And (tmSdf.lChfCode <> 0) Then
            tmPnfSrchKey.iCode = tmChf.iPnfBuyer
            ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If Trim$(tmPnf.sBillAddr(0)) <> "" Then
                    tlagf.sBillAddr(0) = tmPnf.sBillAddr(0)
                    tlagf.sBillAddr(1) = tmPnf.sBillAddr(1)
                    tlagf.sBillAddr(2) = tmPnf.sBillAddr(2)
                ElseIf Trim$(tmPnf.sCntrAddr(0)) <> "" Then
                    tlagf.sBillAddr(0) = tmPnf.sCntrAddr(0)
                    tlagf.sBillAddr(1) = tmPnf.sCntrAddr(1)
                    tlagf.sBillAddr(2) = tmPnf.sCntrAddr(2)
                End If
            End If
        End If

        If Trim$(tlagf.sBillAddr(0)) <> "" Then
            If (InStr(1, tlagf.sBillAddr(0), "Attn:", 1) > 0) Or (InStr(1, tlagf.sBillAddr(1), "Attn:", 1) > 0) Then
                slStr = Trim$(tlagf.sName) & Chr$(10) & Trim$(tlagf.sBillAddr(0)) & Chr$(10) & Trim$(tlagf.sBillAddr(1)) & Chr$(10) & Trim$(tlagf.sBillAddr(2)) & Chr$(0)
            Else
                If ilAttnCode > 0 Then
                    'If tgSpf.sBLaserForm <> "2" Then
                    '    tmPnfSrchKey.iCode = ilAttnCode
                    '    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    'Else
                    '    ilRet = BTRV_ERR_NONE
                    'End If
                    tmPnfSrchKey.iCode = ilAttnCode 'tmAgf.iPnfPay
                    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: " & Trim$(tmPnf.sName) & Chr$(10) & Trim$(tlagf.sBillAddr(0)) & Chr$(10) & Trim$(tlagf.sBillAddr(1)) & Chr$(10) & Trim$(tlagf.sBillAddr(2)) & Chr$(0)
                    Else
                        slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tlagf.sBillAddr(0)) & Chr$(10) & Trim$(tlagf.sBillAddr(1)) & Chr$(10) & Trim$(tlagf.sBillAddr(2)) & Chr$(0)
                    End If
                Else
                    slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tlagf.sBillAddr(0)) & Chr$(10) & Trim$(tlagf.sBillAddr(1)) & Chr$(10) & Trim$(tlagf.sBillAddr(2)) & Chr$(0)
                End If
            End If
        Else
            If (InStr(1, tlagf.sCntrAddr(0), "Attn:", 1) > 0) Or (InStr(1, tlagf.sCntrAddr(1), "Attn:", 1) > 0) Then
                slStr = Trim$(tlagf.sName) & Chr$(10) & Trim$(tlagf.sCntrAddr(0)) & Chr$(10) & Trim$(tlagf.sCntrAddr(1)) & Chr$(10) & Trim$(tlagf.sCntrAddr(2)) & Chr$(0)
            Else
                If ilAttnCode > 0 Then
                    'If tgSpf.sBLaserForm <> "2" Then
                    '    tmPnfSrchKey.iCode = ilAttnCode
                    '    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    'Else
                    '    ilRet = BTRV_ERR_NONE
                    'End If
                    tmPnfSrchKey.iCode = ilAttnCode 'tmAgf.iPnfPay
                    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: " & Trim$(tmPnf.sName) & Chr$(10) & Trim$(tlagf.sCntrAddr(0)) & Chr$(10) & Trim$(tlagf.sCntrAddr(1)) & Chr$(10) & Trim$(tlagf.sCntrAddr(2)) & Chr$(0)
                    Else
                        slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tlagf.sCntrAddr(0)) & Chr$(10) & Trim$(tlagf.sCntrAddr(1)) & Chr$(10) & Trim$(tlagf.sCntrAddr(2)) & Chr$(0)
                    End If
                Else
                    slStr = Trim$(tlagf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tlagf.sCntrAddr(0)) & Chr$(10) & Trim$(tlagf.sCntrAddr(1)) & Chr$(10) & Trim$(tlagf.sCntrAddr(2)) & Chr$(0)
                End If
            End If
        End If
        For ilLoop = LBound(tmIvr.sAddr) To UBound(tmIvr.sAddr) Step 1
            tmIvr.sAddr(ilLoop) = ""
            ilRet = gParseItem(slStr, ilLoop + 1, Chr(10), slParse)
            If (ilRet = CP_MSG_NONE) Then
                ilPos = InStr(slParse, Chr$(0))
                If ilPos > 0 Then
                    slParse = Left$(slParse, ilPos - 1)
                End If
                tmIvr.sAddr(ilLoop) = slParse
            End If
        Next ilLoop
    Else
        tmAgf.sPkInvShow = "D"
        If tgSpf.sBLaserForm = "2" Then
            'If Trim$(tmChf.sBuyer) <> "" Then
            '    ilAttnCode = 1
            '    tmPnf.sName = tmChf.sBuyer
            'Else
            '    ilAttnCode = 0
            'End If
            ilAttnCode = tmAdf.iPnfPay
        Else
            ilAttnCode = tmAdf.iPnfPay
        End If
        If Trim$(tmAdf.sBillAddr(0)) <> "" Then
            If (InStr(1, tmAdf.sBillAddr(0), "Attn:", 1) > 0) Or (InStr(1, tmAdf.sBillAddr(1), "Attn:", 1) > 0) Then
                slStr = Trim$(tmAdf.sName) & Chr$(10) & Trim$(tmAdf.sBillAddr(0)) & Chr$(10) & Trim$(tmAdf.sBillAddr(1)) & Chr$(10) & Trim$(tmAdf.sBillAddr(2)) & Chr$(0)
            Else
                If ilAttnCode > 0 Then
                    'If tgSpf.sBLaserForm <> "2" Then
                    '    tmPnfSrchKey.iCode = ilAttnCode
                    '    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    'Else
                    '    ilRet = BTRV_ERR_NONE
                    'End If
                    tmPnfSrchKey.iCode = ilAttnCode 'tmAgf.iPnfPay
                    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: " & Trim$(tmPnf.sName) & Chr$(10) & Trim$(tmAdf.sBillAddr(0)) & Chr$(10) & Trim$(tmAdf.sBillAddr(1)) & Chr$(10) & Trim$(tmAdf.sBillAddr(2)) & Chr$(0)
                    Else
                        slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tmAdf.sBillAddr(0)) & Chr$(10) & Trim$(tmAdf.sBillAddr(1)) & Chr$(10) & Trim$(tmAdf.sBillAddr(2)) & Chr$(0)
                    End If
                Else
                    slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tmAdf.sBillAddr(0)) & Chr$(10) & Trim$(tmAdf.sBillAddr(1)) & Chr$(10) & Trim$(tmAdf.sBillAddr(2)) & Chr$(0)
                End If
            End If
        Else
            If (InStr(1, tmAdf.sCntrAddr(0), "Attn:", 1) > 0) Or (InStr(1, tmAdf.sCntrAddr(1), "Attn:", 1) > 0) Then
                slStr = Trim$(tmAdf.sName) & Chr$(10) & Trim$(tmAdf.sCntrAddr(0)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(1)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(2)) & Chr$(0)
            Else
                If ilAttnCode > 0 Then
                    'If tgSpf.sBLaserForm <> "2" Then
                    '    tmPnfSrchKey.iCode = ilAttnCode
                    '    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    'Else
                    '    ilRet = BTRV_ERR_NONE
                    'End If
                    tmPnfSrchKey.iCode = ilAttnCode 'tmAgf.iPnfPay
                    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: " & Trim$(tmPnf.sName) & Chr$(10) & Trim$(tmAdf.sCntrAddr(0)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(1)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(2)) & Chr$(0)
                    Else
                        slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tmAdf.sCntrAddr(0)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(1)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(2)) & Chr$(0)
                    End If
                Else
                    slStr = Trim$(tmAdf.sName) & Chr$(10) & "Attn: Accounts Payable" & Chr$(10) & Trim$(tmAdf.sCntrAddr(0)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(1)) & Chr$(10) & Trim$(tmAdf.sCntrAddr(2)) & Chr$(0)
                End If
            End If
        End If

        For ilLoop = LBound(tmIvr.sAddr) To UBound(tmIvr.sAddr) Step 1
            tmIvr.sAddr(ilLoop) = ""
            ilRet = gParseItem(slStr, ilLoop + 1, Chr(10), slParse)
            If (ilRet = CP_MSG_NONE) Then
                ilPos = InStr(slParse, Chr$(0))
                If ilPos > 0 Then
                    slParse = Left$(slParse, ilPos - 1)
                End If
                tmIvr.sAddr(ilLoop) = slParse
            End If
        Next ilLoop

        If tmAdf.iArfLkCode > 0 Then
            If tmArf.iCode <> tmAdf.iArfLkCode Then
                tmArfSrchKey.iCode = tmAdf.iArfLkCode
                ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    tmArf.sNmAd(0) = "Missing"
                    tmArf.sNmAd(1) = " "
                    tmArf.sNmAd(2) = " "
                    tmArf.sNmAd(3) = " "
                End If
            End If

            'tmIvr.sPayAddr(1) = Trim$(tmArf.sNmAd(0))
            'tmIvr.sPayAddr(2) = Trim$(tmArf.sNmAd(1))
            'tmIvr.sPayAddr(3) = Trim$(tmArf.sNmAd(2))
            'tmIvr.sPayAddr(4) = Trim$(tmArf.sNmAd(3))
            tmIvr.sPayAddr(0) = Trim$(tmArf.sNmAd(0))
            tmIvr.sPayAddr(1) = Trim$(tmArf.sNmAd(1))
            tmIvr.sPayAddr(2) = Trim$(tmArf.sNmAd(2))
            tmIvr.sPayAddr(3) = Trim$(tmArf.sNmAd(3))
        Else
            'tmIvr.sPayAddr(1) = Trim$(tgSpf.sBPayName)
            'tmIvr.sPayAddr(2) = Trim$(tgSpf.sBPayAddr(0))
            'tmIvr.sPayAddr(3) = Trim$(tgSpf.sBPayAddr(1))
            'tmIvr.sPayAddr(4) = Trim$(tgSpf.sBPayAddr(2))
            tmIvr.sPayAddr(0) = Trim$(tgSpf.sBPayName)
            tmIvr.sPayAddr(1) = Trim$(tgSpf.sBPayAddr(0))
            tmIvr.sPayAddr(2) = Trim$(tgSpf.sBPayAddr(1))
            tmIvr.sPayAddr(3) = Trim$(tgSpf.sBPayAddr(2))
        End If
        tmIvr.iMnfSort = tmAdf.iMnfSort
    End If

End Sub

Private Sub mAdjDateTime(ilVefCode As Integer, slDate As String, slTime As String)
    Dim ilTimeAdj As Integer
    Dim ilZone As Integer
    Dim ilVpf As Integer
    Dim llTime As Long
    
    If (tgSpf.sInvSpotTimeZone = "E") Or (tgSpf.sInvSpotTimeZone = "C") Or (tgSpf.sInvSpotTimeZone = "M") Or (tgSpf.sInvSpotTimeZone = "P") Then
        ilTimeAdj = 0
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            'For ilZone = 1 To 5 Step 1
            For ilZone = LBound(tgVpf(ilVpf).sGZone) To UBound(tgVpf(ilVpf).sGZone) Step 1
                If Left$(tgVpf(ilVpf).sGZone(ilZone), 1) = tgSpf.sInvSpotTimeZone Then
                    ilTimeAdj = tgVpf(ilVpf).iGLocalAdj(ilZone)
                    Exit For
                End If
            Next ilZone
        End If
        If ilTimeAdj <> 0 Then
            llTime = gTimeToLong(slTime, False) + (CLng(ilTimeAdj) * 3600)
            If llTime < 0 Then
                llTime = 86400 - llTime
                slDate = gDecOneDay(slDate)
            ElseIf llTime > 86399 Then
                llTime = llTime - 86400
                slDate = gIncOneDay(slDate)
            End If
            slTime = gFormatTimeLong(llTime, "A", "1")
        End If
    End If
End Sub
'               Barter feature on, vehicle shown on Insertion Orders, and vehicle is posted by radio station invoicing.
'               Read Iihf to see if it exists
Private Sub mGatherPostBarterCounts(ilVefCode As Integer, llStartDate As Long, llEndDate As Long)
Dim ilRet As Integer
Dim llPostedInvDate As Long
Dim ilLoopOnCnt As Long
Dim blFound As Boolean
Dim ilUpper As Integer
Dim ilLoopOnWeek As Integer

        'read all IIhf for the period requested
        'Any that exists and has post by counts must be saved in table
        tmIihfSrchKey1.iVefCode = ilVefCode
        gPackDateLong llStartDate, tmIihfSrchKey1.iInvStartDate(0), tmIihfSrchKey1.iInvStartDate(1)
        
        ilRet = btrGetGreaterOrEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
        gUnpackDateLong tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), llPostedInvDate
        Do While ilRet = BTRV_ERR_NONE And llPostedInvDate >= llStartDate And llPostedInvDate <= llEndDate And tmIihf.iVefCode = ilVefCode
            If Trim$(tmIihf.sSourceForm) = "C" Then        'show aff of performance by counts
                blFound = False
                For ilLoopOnCnt = LBound(tmPostBarterCounts) To UBound(tmPostBarterCounts) - 1
                    If tmPostBarterCounts(ilLoopOnCnt).lChfCode = tmIihf.lChfCode And tmPostBarterCounts(ilLoopOnCnt).iVefCode = tmIihf.iVefCode Then
                        blFound = True
                        Exit For
                    End If
                Next ilLoopOnCnt
                If Not blFound Then
                    ilUpper = UBound(tmPostBarterCounts)
                    tmPostBarterCounts(ilUpper).lChfCode = tmIihf.lChfCode
                    tmPostBarterCounts(ilUpper).iVefCode = tmIihf.iVefCode
                    ReDim Preserve tmPostBarterCounts(LBound(tmPostBarterCounts) To UBound(tmPostBarterCounts) + 1) As POSTBARTERCOUNTS
                End If
            End If
            ilRet = btrGetNext(hmIihf, tmIihf, imIihfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        
        Exit Sub
End Sub
'
'           mTestSpotForBarterCounts - determine if the spot should be shown as an air time spot
'           or just with counts
'           <input>  llStartWeek - start monday of request
'
'           return - true : ok to print, invoicing not by counts
Private Function mTestSpotForBarterCounts() As Boolean
Dim blIsItBarterByCount As Boolean
Dim ilLoopOnCnt As Integer
Dim llSpotDate As Long

            blIsItBarterByCount = False
            For ilLoopOnCnt = LBound(tmPostBarterCounts) To UBound(tmPostBarterCounts) - 1
                If tmPostBarterCounts(ilLoopOnCnt).lChfCode = tmSdf.lChfCode And tmPostBarterCounts(ilLoopOnCnt).iVefCode = tmSdf.iVefCode Then
                    blIsItBarterByCount = True
                    Exit For
                End If
            Next ilLoopOnCnt
        
            mTestSpotForBarterCounts = blIsItBarterByCount
            Exit Function
            
End Function
