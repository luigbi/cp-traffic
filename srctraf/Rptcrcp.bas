Attribute VB_Name = "RPTCRCP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrcp.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCRGet.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report Get Data for Crystal screen code
Option Explicit
Option Compare Text
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
Dim tlChfAdvtExt() As CHFADVTEXT
'The following arrays are built by the schedule line for as many weeks as there are in the order
Dim tmAllCnt() As CPPCPMLIST
Dim lmAllGrimp() As Long
Dim lmAllGrp() As Long
Dim lmAllCost() As Long
Dim imAllRtg() As Integer
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
'Dim hmUrf As Integer            'User file handle
'Dim imUrfRecLen As Integer      'URF record length
'Dim tmUrf As URF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
'Dim hmVsf As Integer            'Virtual Vehicle file handle
'Dim tmVsf As VSF                'VSF record image
'Dim tmVsfSrchKey As LONGKEY0            'VSF record image
'Dim imVsfRecLen As Integer        'VSF record length
'Dim hmVlf As Integer            'Vehicle Link file handle
'Dim tmVlf As VLF                'VLF record image
'Dim tmVlfSrchKey0 As VLFKEY0            'VLF by selling vehicle record image
'Dim tmVlfSrchKey1 As VLFKEY1            'VLF by airing vehicle record image
'Dim imVlfRecLen As Integer        'VLF record length
Dim hmSlf As Integer            'Slsp file handle
Dim tmSlf As SLF                'Slsp record image
Dim imSlfRecLen As Integer        'Slsp record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GRF record length
Dim tmZeroGrf As GRF              'initialized  recd
'  Rating Book File
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length
'  Demo Plus Research File
Dim hmDpf As Integer        'DEmo plus book file handle
Dim tmDpf As DPF            'DPF record image
Dim imDpfRecLen As Integer  'DPF record length
Dim hmDef As Integer
Dim hmRaf As Integer


'
'
'                   gCPPCPMGen - Generate prepass file for
'                   Cost Per Point/ Cost Per Thousands report.
'                   This report generated CPP or CPMs for each
'                   demo category for the vehicle, advertiser, product
'                   and spot length.  CPP or CPMs are shown for up to
'                   4 quarters, plus the yearly CPP or CPM
'
'                   4/10/98 Exclude 100% trade contracts
'       2-4-04 Exclude history lines when gathering schedule lines.  Spots/$ overstated.
'       6-1-04 removed unused variables; implement Demo estimates
'       6-4-04 add selective contract (for debugging)
Sub gCppCpmGen()
Dim ilRet As Integer
Dim slKey As String                     'formatted key field for demos:  5 char demo, 5 char veh, 5 char advt, 3 char length, 35 char prod. name
Dim ilfirstTime As Integer
Dim ilFoundCnt As Integer               'valid contract to process
'Dim ilCurrentRecd As Integer            'number of contracts processed so far
Dim llCurrentRecd As Long               '3-16-10 integer exceeded, chg to long
Dim illoop As Integer                   'temp loop variable
Dim ilLoop3 As Integer                  'temp loop variable
Dim llContrCode As Long                 'Contr ID to process
Dim slStartDate As String               'Contract start date
Dim slEndDate As String                 'contract end date
Dim ilClf As Integer                    'loop for lines
Dim ilCff As Integer                    'loop for flights
Dim slStr As String                     'temp string for conversions
Dim ilQtr As Integer
Dim llDate As Long                      'temp serial date
Dim llDate2 As Long
Dim ilDay As Integer
Dim slNameCode As String
Dim slCode As String
Dim slCntrType As String                'valid contract types (per inq, direct respon, etc) to retrieve
Dim slCntrStatus As String              'valid contr status (working, complete, etc) to retrieve
Dim ilHOState As Integer                'which type of Holds Orders to retrieved (internally WCI)
Dim llPop As Long                       'population obtained per schedule line
Dim llAvgAud As Long                    'avg audience obtained per flight
Dim llGetTo As Long                     'test contract's date entered against this span
Dim llGetFrom As Long                   'test contract's date entered against this span
Dim ilMonth As Integer
Dim ilYear As Integer
Dim ilCorpStd As Integer                '1 = corp, 2 = std
Dim llFltStart As Long
Dim llFltEnd As Long
Dim ilSpots As Integer
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim llLineBase As Long
Dim ilInputQtr As Integer           '# qtrs requested
'ReDim llDateS(1 To 2) As Long        'year dates
ReDim llDateS(0 To 2) As Long        'year dates. Index zero ignored
ReDim ilInputDays(0 To 6) As Integer    'valid days of the week for audience retrieval
Dim ilUpperWk As Integer
Dim llTemp As Long
'Dim ilWinx As Integer
Dim llWinx As Long                  '3-16-10 integer exceeded, chg to long
Dim llTotalCPP As Long
Dim llTotalCPM As Long
Dim llTotalGrImp As Long
Dim llTotalGRP As Long
'Dim llTotalCost As Long
Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
Dim llTotalAvgAud As Long
Dim ilTotalAvgRtg As Integer
'ReDim llQtrSpots(1 To 13) As Integer       'array of weekly spots in qtr
'ReDim llQtrRates(1 To 13) As Long           'array of weekly spot rates in qtr
'ReDim llQtrAvgAud(1 To 13) As Long             'array of wekly avg aud in qtr
'ReDim llQtrPopEst(1 To 13) As Long
ReDim llQtrSpots(0 To 12) As Long       'array of weekly spots in qtr
ReDim llQtrRates(0 To 12) As Long           'array of weekly spot rates in qtr
ReDim llQtrAvgAud(0 To 12) As Long             'array of wekly avg aud in qtr
ReDim llQtrPopEst(0 To 12) As Long
'ReDim ilRtg(1 To 52) As Integer             'return from gAvgAudToLnResearch
'ReDim llGrImp(1 To 52) As Long              'return from gAvgAudToLnResearch
'ReDim llGRP(1 To 52) As Long                'return from gAvgAudToLnResearch
ReDim ilRtg(0 To 51) As Integer             'return from gAvgAudToLnResearch
ReDim llGrImp(0 To 51) As Long              'return from gAvgAudToLnResearch
ReDim llGRP(0 To 51) As Long                'return from gAvgAudToLnResearch

Dim ilTotLnSpts As Integer                  'sum of spots per line
Dim llTotLnGross As Long                    'sum of gross $ per line
Dim ilCorpCalInx As Integer               'index into tgCof for the year processing
ReDim ilEffDate(0 To 1) As Integer            'Effective date entered stored as btrieve date
Dim llPopEst As Long
Dim llSingleCntr As Long                    '6-4-04
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long
Dim slGrossNet As String * 1                '1-25-19 Gross net option
Dim llRate As Long

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    ReDim tgClfCP(0 To 0) As CLFLIST
    tgClfCP(0).iStatus = -1 'Not Used
    tgClfCP(0).lRecPos = 0
    tgClfCP(0).iFirstCff = -1
    ReDim tgCffCP(0 To 0) As CFFLIST
    tgCffCP(0).iStatus = -1 'Not Used
    tgCffCP(0).lRecPos = 0
    tgCffCP(0).iNextCff = -1
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tgClfCP(0).ClfRec)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tgCffCP(0).CffRec)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)

    'hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    'ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    ilRet = btrClose(hmUrf)
    '    ilRet = btrClose(hmMnf)
    '    ilRet = btrClose(hmCff)
    '    ilRet = btrClose(hmClf)
    '    ilRet = btrClose(hmChf)
    '    btrDestroy hmUrf
    '    btrDestroy hmMnf
    '    btrDestroy hmCff
    '    btrDestroy hmClf
    '    btrDestroy hmChf
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'imUrfRecLen = Len(tmUrf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
    '    ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
    '    btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmGrf)
   '     ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmDrf
        btrDestroy hmGrf
   '     btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDrfRecLen = Len(tmDrf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmGrf)
    '    ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmDrf
        btrDestroy hmGrf
   '     btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmGrf)
    '    ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmDpf
        btrDestroy hmSlf
        btrDestroy hmDrf
        btrDestroy hmGrf
   '     btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDpfRecLen = Len(tmDpf)
    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmGrf)
   '     ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmSlf
        btrDestroy hmDrf
        btrDestroy hmGrf
    '    btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmGrf)
   '     ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmSlf
        btrDestroy hmDrf
        btrDestroy hmGrf
    '    btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '7-23-01 setup global variable to determine if demo plus info exists
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    'Build  slsp in memory to avoid rereading for office
    'ReDim tlSlf(1 To 1) As SLF
    ReDim tlSlf(0 To 0) As SLF
    ilRet = gObtainSlf(RptSelCp, hmSlf, tlSlf())

    If RptSelCp!rbcSelCInclude(0).Value Then        'corp
        ilCorpStd = 1
        ilCorpCalInx = gGetCorpCalIndex(igYear)
    Else
        ilCorpStd = 2
    End If

    If RptSelCp!edcContract <> "" Then                      '6-4-04
        llSingleCntr = CLng(RptSelCp!edcContract)
    End If

    slGrossNet = "G"                                '1-25-19 Implement gross net option; default to gross
    If RptSelCp!rbcGrossNet(1).Value Then
        slGrossNet = "N"
    End If
    
'    slStr = RptSelCp!edcSelCFrom.Text               'effective date entred
    slStr = RptSelCp!CSI_CalFrom.Text               '12-13-19 change to use csi calendar control vs edit text box; effective date entred
    'obtain the entered dates year based on the std month
    llGetTo = gDateValue(slStr)                     'gather contracts thru this date
    slStr = Format$(llGetTo, "m/d/yy")               'reformat date to insure year is there
    gPackDate slStr, ilEffDate(0), ilEffDate(1)

    If ilCorpStd = 1 Then                               'corporate
        'gUnpackDate tgMCof(ilCorpCalInx).iStartDate(0, 1), tgMCof(ilCorpCalInx).iStartDate(1, 1), slStr         'convert last bdcst billing date to string
        gUnpackDate tgMCof(ilCorpCalInx).iStartDate(0, 0), tgMCof(ilCorpCalInx).iStartDate(1, 0), slStr         'convert last bdcst billing date to string
        llGetFrom = gDateValue(gObtainStartCorp(slStr, True))  'gather contracts from this date thru effective entered date
        illoop = (igMonthOrQtr - 1) * 3 + 1             'determine starting month based on qtr entred
        gUnpackDate tgMCof(ilCorpCalInx).iStartDate(0, illoop - 1), tgMCof(ilCorpCalInx).iStartDate(1, illoop - 1), slStr     'convert last bdcst billing date to string
        llDateS(1) = gDateValue(gObtainStartCorp(slStr, True))  'gather contracts from this date thru effective entered date
    Else                                                'standard
        slStr = gObtainEndStd(Format$(llGetTo, "m/d/yy"))
        gObtainMonthYear 0, slStr, ilMonth, ilYear           'get year  of effective date (to figure out the beginning of std year)
        slStr = "1/15/" & Trim$(str$(ilYear))                 'Jan of std year effective dat entered
        llGetFrom = gDateValue(gObtainStartStd(slStr))  'gather contracts from this date thru effective entered date
        ilYear = Val(RptSelCp!edcSelCTo.Text)           'year requested
        'Determine this years quarter span
        illoop = (igMonthOrQtr - 1) * 3 + 1             'determine starting month based on qtr entred
        slStr = Trim$(str$(illoop)) & "/15/" & Trim$(RptSelCp!edcSelCTo.Text)
        llDateS(1) = gDateValue(gObtainStartStd(slStr))
    End If
    slStr = Trim$(str$(RptSelCp!edcSelCFrom1.Text))
    ilInputQtr = Val(slStr)
    llDateS(2) = llDateS(1) + ((ilInputQtr) * 91) - 1         'end date of this reporting period requested
    slStartDate = Trim$(Format$(llDateS(1), "m/d/yy"))
    slEndDate = Trim$(Format$(llDateS(2), "m/d/yy"))
    slCntrType = gBuildCntTypes()
    slCntrStatus = "HOGN"              'only get holds and orders
    ilHOState = 2                  'get latest orders and revisions
    ilRet = gObtainCntrForDate(RptSelCp, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())

    ReDim tlCntList(0 To 0) As CPPCPMLIST           'list of Research totals for all lines, all cnt
    For llCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1                                           'loop while llCurrentRecd < llRecsRemaining
        'check for valid demo selection
        If (gSetCheck(RptSelCp!ckcAll)) Then
            llContrCode = tlChfAdvtExt(llCurrentRecd).lCode
            '6-4-04 implement single contract option for debugging
            If llSingleCntr > 0 And llSingleCntr <> tlChfAdvtExt(llCurrentRecd).lCntrNo Then
                ilFoundCnt = False
            Else
                 'Retrieve the contract, schedule lines and flights
                llContrCode = gPaceCntr(tlChfAdvtExt(llCurrentRecd).lCntrNo, llGetTo, hmCHF, tmChf)
                If llContrCode > 0 Then
                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCP, tgClfCP(), tgCffCP())
                    ilFoundCnt = True
                Else                        'didnt find a cntr within pacing effective date
                    ilFoundCnt = False
                End If
            End If
        Else
            '6-4-04 implement single contract option for debugging
            If llSingleCntr > 0 And llSingleCntr <> tlChfAdvtExt(llCurrentRecd).lCntrNo Then
                ilFoundCnt = False
            Else
                ilFoundCnt = False                             'assume nothing found until match in demo selection table
                For illoop = 0 To RptSelCp!lbcSelection(0).ListCount - 1 Step 1
                    If RptSelCp!lbcSelection(0).Selected(illoop) Then
                        slNameCode = tgRptSelDemoCodeCP(illoop).sKey
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tlChfAdvtExt(llCurrentRecd).iMnfDemo0 Then
                            llContrCode = tlChfAdvtExt(llCurrentRecd).lCode
                            'Retrieve the contract, schedule lines and flights
                            llContrCode = gPaceCntr(tlChfAdvtExt(llCurrentRecd).lCntrNo, llGetTo, hmCHF, tmChf)
                            If llContrCode > 0 Then
                                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCP, tgClfCP(), tgCffCP())
                                ilFoundCnt = True
                                Exit For
                            End If
                        End If
                    End If
                Next illoop
            End If
        End If
        'convert date entered to long
        gUnpackDateLong tgChfCP.iOHDDate(0), tgChfCP.iOHDDate(1), llTemp
        If tgChfCP.iMnfDemo(0) = 0 Or llTemp > llGetTo Then       'if demo doesnt exist or the date entered is later than requested
            ilFoundCnt = False
        End If

        'debug
        'If tgChfCP.lCntrNo = 22521 Or tgChfCP.lCntrNo = 22519 Then
        '    ilFoundCnt = True
       ' Else
        '    ilFoundCnt = False
        'End If

        If (ilFoundCnt And tgChfCP.iPctTrade <> 100) Then                                 'get a contract and test for printables,
                                                        'user input Entered and Active dates
            For ilClf = LBound(tgClfCP) To UBound(tgClfCP) - 1 Step 1
                ilTotLnSpts = 0                 'init total # spots per line
                llTotLnGross = 0                'init total dollars this line
                tmClf = tgClfCP(ilClf).ClfRec
                If tmClf.sType = "H" Or tmClf.sType = "S" Then      'only makes sense to do cpp cpm on standard or hidden lines
                    'obtain population and demo codes by schedule line
                    'ReDim ilWklyspotsl(1 To 52) As Integer       'sched lines weekly # spots
                    'ReDim llWklyAvgAud(1 To 52) As Long             'sched lines weekly avg aud
                    'ReDim llWklyRates(1 To 52) As Long           'sched lines weekly rates
                    'ReDim llWklyPopEst(1 To 52) As Long
                    ReDim ilWklyspotsl(0 To 51) As Long       'sched lines weekly # spots
                    ReDim llWklyAvgAud(0 To 51) As Long             'sched lines weekly avg aud
                    ReDim llWklyRates(0 To 51) As Long           'sched lines weekly rates
                    ReDim llWklyPopEst(0 To 51) As Long
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, tmClf.iDnfCode, 0, tgChfCP.iMnfDemo(0), llPop)
                    gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLineBase
                    gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llTemp
                    'Got a contract that has at least a line that spans the requested dates, now
                    'process only those lines within the rquested dates
                    'the line end date must be equal/after the earliest requested date and the
                    'line start date must be less than the requested latest date
                    If llTemp >= gDateValue(slStartDate) And llLineBase <= gDateValue(slEndDate) Then
                        If llLineBase < llDateS(1) Then
                            llLineBase = llDateS(1)         'start of line cant be any earlier than whats being requested
                        End If
                        If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                            llOvStartTime = 0
                            llOvEndTime = 0
                        Else
                            'override times exist
                            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                        End If

                        ilCff = tgClfCP(ilClf).iFirstCff
                        Do While ilCff <> -1
                            tmCff = tgCffCP(ilCff).CffRec
                            For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(illoop) = False
                            Next illoop

                            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                            llFltStart = gDateValue(slStr)
                            'backup start date to Monday
                            illoop = gWeekDayLong(llFltStart)
                            Do While illoop <> 0
                                llFltStart = llFltStart - 1
                                illoop = gWeekDayLong(llFltStart)
                            Loop
                            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                            llFltEnd = gDateValue(slStr)
                            '
                            'Loop thru the flight by week and build the number of spots for each week
                            '
                            'Start of flight cant be any earlier than what has been requested
                            If llFltStart < llDateS(1) Then
                                llFltStart = llDateS(1)
                            End If
                            If llFltEnd > llDateS(2) Then       'end date of flight cant be later than the period requested
                                llFltEnd = llDateS(2)
                            End If
                            For llDate2 = llFltStart To llFltEnd Step 7
                                If tmCff.sDyWk = "W" Then            'weekly
                                    ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                     For ilDay = 0 To 6 Step 1
                                        If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                                            If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        End If
                                     Next ilDay
                                Else                                        'daily
                                     If illoop + 6 < llFltEnd Then           'we have a whole week
                                        ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                        For ilDay = 0 To 6 Step 1
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next ilDay
                                     Else                                    'do partial week
                                        For llDate = llDate2 To llFltEnd Step 1
                                            ilDay = gWeekDayLong(llDate)
                                            ilSpots = ilSpots + tmCff.iDay(ilDay)
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next llDate
                                    End If
                                End If

                                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmClf.iDnfCode, tmClf.iVefCode, 0, tgChfCP.iMnfDemo(0), llDate2, llDate2, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                'Loop and build avg aud, spots, & spots per week
                                'ilUpperWk = UBound(ilWklyspotsl)
                                ilUpperWk = (llDate2 - llDateS(1)) / 7 + 1      'determine week index to start of requested period
                                ilWklyspotsl(ilUpperWk - 1) = ilSpots
                                'determine the # of 30" units
                                'ilMultiplyBy = 1
                                'If tmClf.iLen >= 30 Then
                                '    ilMultiplyBy = tmClf.iLen \ 30
                                'End If
                                'ilTotLnSpts = ilTotLnSpts + (ilSpots * ilMultiplyBy)
                                ilTotLnSpts = ilTotLnSpts + ilSpots
                                '1-25-19 implement gross net option
                                llRate = gGetGrossOrNetFromRate(tmCff.lActPrice, slGrossNet, tgChfCP.iAgfCode)
                                llWklyRates(ilUpperWk - 1) = llRate
                                llTotLnGross = llTotLnGross + (llRate * ilSpots)
'                                llWklyRates(ilUpperWk - 1) = tmCff.lActPrice
'                                llTotLnGross = llTotLnGross + (tmCff.lActPrice * ilSpots)
                                llWklyAvgAud(ilUpperWk - 1) = llAvgAud
                                llWklyPopEst(ilUpperWk - 1) = llPopEst
                            Next llDate2
                            ilCff = tgCffCP(ilCff).iNextCff               'get next flight record from mem
                        Loop                                            'while ilcff <> -1
                    End If                                              'line outside range of requested dates
                    'Finished all flights, calculat the lines research values
                    'Schedule line complete, get its avg aud data for the line
                    'Calculate Research by line for each quarter required.  Dont bother
                    'if any of these fields are zero
                    If ilTotLnSpts > 0 And llTotLnGross > 0 And llPop > 0 Then
                        'Create key as a string - used for sorting later
                        'Build demo code, Demo book Code, Vehicle code, advt code, sell office code, spot length & product
                        slStr = Trim$(str$(tgChfCP.iMnfDemo(0)))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        slKey = slStr & "|"

                        slStr = Trim$(str$(tmClf.iVefCode))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        slKey = slKey & slStr & "|"

                        slStr = Trim$(str$(tgChfCP.iAdfCode))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        slKey = slKey & slStr & "|"

                        For illoop = LBound(tlSlf) To UBound(tlSlf) - 1 Step 1
                            slStr = "00000"
                            If tgChfCP.iSlfCode(0) = tlSlf(illoop).iCode Then
                                slStr = Trim$(str$(tlSlf(illoop).iSofCode))
                                Do While Len(slStr) < 5
                                    slStr = "0" & slStr
                                Loop
                                Exit For
                            End If
                        Next illoop
                        slKey = slKey & slStr & "|"

                        slStr = Trim$(str$(tmClf.iLen))
                        Do While Len(slStr) < 3
                            slStr = "0" & slStr
                        Loop
                        slKey = slKey & slStr & "|" & tgChfCP.sProduct

                        'Create a new entry for each line.  They will be combined by key later for the total research values
                        llWinx = UBound(tlCntList)

                        'loop for max 4 quarters to get ratings, grimps, grps and cost for the quarter
                        For ilQtr = 1 To ilInputQtr
                            ilLoop3 = (ilQtr - 1) * 13 + 1
                            For illoop = 1 To 13
                                'llQtrSpots(ilLoop - 1) = ilWklyspotsl(ilLoop3 + ilLoop - 1)
                                'llQtrRates(ilLoop - 1) = llWklyRates(ilLoop3 + ilLoop - 1)
                                'llQtrAvgAud(ilLoop - 1) = llWklyAvgAud(ilLoop3 + ilLoop - 1)
                                'llQtrPopEst(ilLoop - 1) = llWklyPopEst(ilLoop3 + ilLoop - 1)
                                llQtrSpots(illoop - 1) = ilWklyspotsl(ilLoop3 + illoop - 2)
                                llQtrRates(illoop - 1) = llWklyRates(ilLoop3 + illoop - 2)
                                llQtrAvgAud(illoop - 1) = llWklyAvgAud(ilLoop3 + illoop - 2)
                                llQtrPopEst(illoop - 1) = llWklyPopEst(ilLoop3 + illoop - 2)
                            Next illoop
                            '10-30-14 default to use 1 place rating
                            'gAvgAudToLnResearch "1", False, llPop, llQtrPopEst(), llQtrSpots(), llQtrRates(), llQtrAvgAud(), llTotalCost, llTotalAvgAud, ilRtg(), ilTotalAvgRtg, llGrImp(), llTotalGrImp, llGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                            gAvgAudToLnResearch "1", False, llPop, llQtrPopEst(), llQtrSpots(), llQtrRates(), llQtrAvgAud(), dlTotalCost, llTotalAvgAud, ilRtg(), ilTotalAvgRtg, llGrImp(), llTotalGrImp, llGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                            'Build totals by line
                            tlCntList(llWinx).sKey = slKey
                            If tgSpf.sDemoEstAllowed = "Y" Then
                                tlCntList(llWinx).lPop(ilQtr) = llPopEst       '6-1-04 estimated pop by qtr or normal line pop also returned in llestpop  (was llPop)
                            Else
                                 tlCntList(llWinx).lPop(ilQtr) = llPop
                            End If
                            'tlCntList(llWinx).lCost(ilQtr) = llTotalCost
                            tlCntList(llWinx).lCost(ilQtr) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                            tlCntList(llWinx).lCPP(ilQtr) = llTotalCPP
                            tlCntList(llWinx).lCPM(ilQtr) = llTotalCPM
                            tlCntList(llWinx).lGrImp(ilQtr) = llTotalGrImp
                            tlCntList(llWinx).lGRP(ilQtr) = llTotalGRP
                            tlCntList(llWinx).iRtg(ilQtr) = ilTotalAvgRtg
                            tlCntList(llWinx).lAvgAud(ilQtr) = llTotalAvgAud
                        Next ilQtr

                        'Quarters obtained, do the yearly cpp & cpm
                        'Get the line research only for the # of qtrs requested
                        'ReDim Preserve ilWklyspotsl(1 To 13 * ilInputQtr) As Integer
                        'ReDim Preserve llWklyRates(1 To 13 * ilInputQtr) As Long
                        'ReDim Preserve llWklyAvgAud(1 To 13 * ilInputQtr) As Long
                        'ReDim Preserve llWklyPopEst(1 To 13 * ilInputQtr) As Long
                        ReDim Preserve ilWklyspotsl(0 To 13 * ilInputQtr - 1) As Long
                        ReDim Preserve llWklyRates(0 To 13 * ilInputQtr - 1) As Long
                        ReDim Preserve llWklyAvgAud(0 To 13 * ilInputQtr - 1) As Long
                        ReDim Preserve llWklyPopEst(0 To 13 * ilInputQtr - 1) As Long
                        '10-30-14 default to use 1 place rating regardless of agency flag
                        'gAvgAudToLnResearch "1", False, llPop, llWklyPopEst(), ilWklyspotsl(), llWklyRates(), llWklyAvgAud(), llTotalCost, llTotalAvgAud, ilRtg(), ilTotalAvgRtg, llGrImp(), llTotalGrImp, llGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                        gAvgAudToLnResearch "1", False, llPop, llWklyPopEst(), ilWklyspotsl(), llWklyRates(), llWklyAvgAud(), dlTotalCost, llTotalAvgAud, ilRtg(), ilTotalAvgRtg, llGrImp(), llTotalGrImp, llGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                        'tlCntList(llWinx).lCost(5) = llTotalCost
                        tlCntList(llWinx).lCost(5) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                        tlCntList(llWinx).lCPP(5) = llTotalCPP
                        tlCntList(llWinx).lCPM(5) = llTotalCPM
                        tlCntList(llWinx).lGrImp(5) = llTotalGrImp
                        tlCntList(llWinx).lGRP(5) = llTotalGRP
                        tlCntList(llWinx).iRtg(5) = ilTotalAvgRtg
                        tlCntList(llWinx).lAvgAud(5) = llTotalAvgAud
                        tlCntList(llWinx).iSpots = ilTotLnSpts          'total spots this line
                        'tlCntList(llWinx).lGross = llTotalCost          'total $ this line
                        tlCntList(llWinx).lGross = dlTotalCost          'total $ this line'TTP 10439 - Rerate 21,000,000
                        tlCntList(llWinx).iDnfCode = tmClf.iDnfCode     'demo book name
                        tlCntList(llWinx).lPop(5) = llPopEst      '6-1-04
                        ReDim Preserve tlCntList(0 To UBound(tlCntList) + 1) As CPPCPMLIST  '6-1-04
                    End If
                End If                  'tmclf.stype = H or tmclf.stype = S
            Next ilClf                  'get next line
        End If                          'ilfoundcnt = true
    Next llCurrentRecd                  'get another cnt

    'Sort the Array so that all like elements are together.  Then gather the like
    'elements to get the CPP or CPM
    If UBound(tlCntList) > 0 Then
        ArraySortTyp fnAV(tlCntList(), 0), UBound(tlCntList), 0, LenB(tlCntList(0)), 0, LenB(tlCntList(0).sKey), 0
    End If

    ilfirstTime = True
    ReDim tmAllCnt(0 To 0) As CPPCPMLIST
    If UBound(tlCntList) > 0 Then
        For llCurrentRecd = 0 To UBound(tlCntList) - 1 Step 1
            If ilfirstTime Then
                slKey = tlCntList(llCurrentRecd).sKey
                ilfirstTime = False
            End If
            If slKey = tlCntList(llCurrentRecd).sKey Then               'is this key same as the one saved
                 tmAllCnt(UBound(tmAllCnt)) = tlCntList(llCurrentRecd)
                 ReDim Preserve tmAllCnt(0 To UBound(tmAllCnt) + 1) As CPPCPMLIST
            Else
                mWriteCPPCPM slKey, ilEffDate()
                ReDim tmAllCnt(0 To 0) As CPPCPMLIST
                slKey = tlCntList(llCurrentRecd).sKey
                tmAllCnt(0) = tlCntList(llCurrentRecd)
                ReDim Preserve tmAllCnt(0 To 1) As CPPCPMLIST
            End If
        Next llCurrentRecd
        mWriteCPPCPM slKey, ilEffDate()       'write out last record
    End If
    'ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDpf)
    ilRet = btrClose(hmDef)
    ilRet = btrClose(hmRaf)
    ilRet = btrClose(hmSlf)

    btrDestroy hmRaf
    btrDestroy hmDef
    btrDestroy hmDpf
    btrDestroy hmSlf
    btrDestroy hmDrf
    btrDestroy hmGrf
    'btrDestroy hmUrf
    btrDestroy hmMnf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCHF

    Erase tgClfCP, tgCffCP, tlSlf
    Erase ilWklyspotsl, llWklyRates, llWklyAvgAud, llWklyPopEst
    Erase imAllRtg, lmAllGrimp, lmAllGrp, lmAllCost
    Erase tlCntList, tmAllCnt
    Erase llGrImp, llGRP, ilRtg
    Erase lmAllGrimp, lmAllGrp, lmAllCost, imAllRtg
End Sub
'
'
'                   mWriteCPPCPM - Obtain the Research values (CPP/CPM)
'                   for the quarters and year.  Write a CBF record to
'                   disk
'
'                   Created:  10/14/97
'
Sub mWriteCPPCPM(slKey As String, ilEffDate() As Integer)
Dim ilRet As Integer            'err return from parse rtn
Dim ilQtr As Integer            'loop on 5 quarters (1-4 = quarters, 5 = year)
Dim slCode As String            'parsing temp string
'3-16-10 integer exceeded, chg to long
Dim ilCurrentRecd As Integer
Dim llCurrentRecd As Long    'loop on unique entries for: demo cat, vehicle, advt, sell office, spot length & product
Dim ilUpper As Integer          '# of entries to send to ResearchTotal routine
Dim ilBookNameCode As Integer   'Book Name used for each entry.  If didfferent, population must be set to 0 for ReserachTotal rtn
Dim llPop As Long               'population used for the selected book (0 if different books across vehicles)
'Dim lRate As Long               'Avg Rate from ResearchTotal rtn
Dim dRate As Double             'Avg Rate from ResearchTotal rtn'TTP 10439 - Rerate 21,000,000
Dim iAvgRtg As Integer          'avg rating from ResearchTotal rtn
Dim lGrImp As Long              'total grimps from ResearchTotal rtn
Dim lGRP As Long                'total GRP from ResearchTotal rtn
Dim lCPP As Long                'total CPP from ResearchTotal rtn
Dim lCPM As Long                'total CPM from ResearchTotal rtn
Dim lAvgAud As Long             'total avg aud from ResearchTotal rtn
Dim llLnSpots As Long
    tmGrf = tmZeroGrf
    ilRet = gParseItem(slKey, 1, "|", slCode)       'parse outdemographic category
    tmGrf.iRdfCode = Val(slCode)
    ilRet = gParseItem(slKey, 2, "|", slCode)       'parse out vehicle code
    tmGrf.iVefCode = Val(slCode)
    ilRet = gParseItem(slKey, 3, "|", slCode)       'parse out advt code
    tmGrf.iAdfCode = Val(slCode)
    ilRet = gParseItem(slKey, 4, "|", slCode)       'parse out selling office code
    tmGrf.iSofCode = Val(slCode)
    ilRet = gParseItem(slKey, 5, "|", slCode)       'parse out spot length
    'tmGrf.iPerGenl(1) = Val(slCode)
    tmGrf.iPerGenl(0) = Val(slCode)
    ilRet = gParseItem(slKey, 6, "|", tmGrf.sGenDesc)       'Product Name
    ilBookNameCode = tmAllCnt(0).iDnfCode
    'llPop = tmAllCnt(0).lPop       '6-2-04


    ilUpper = UBound(tmAllCnt) - 1
    ReDim lmAllGrimp(0 To ilUpper) As Long
    ReDim lmAllGrp(0 To ilUpper) As Long
    ReDim lmAllCost(0 To ilUpper) As Long
    ReDim imAllRtg(0 To ilUpper) As Integer
    For ilQtr = 1 To 5                      'process one quarter at a time
        llPop = -1                          '6-2-04
        For llCurrentRecd = 0 To UBound(tmAllCnt) - 1 Step 1
            If ilQtr = 5 Then           'only accum the total spots & $ when the year is being processed
                'tmGrf.iPerGenl(2) = tmGrf.iPerGenl(2) + tmAllCnt(llCurrentRecd).iSpots
                tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + tmAllCnt(llCurrentRecd).iSpots
                tmGrf.lCode4 = tmGrf.lCode4 + (tmAllCnt(llCurrentRecd).lGross)
            End If
            lmAllGrimp(llCurrentRecd) = tmAllCnt(llCurrentRecd).lGrImp(ilQtr)
            lmAllGrp(llCurrentRecd) = tmAllCnt(llCurrentRecd).lGRP(ilQtr)
            lmAllCost(llCurrentRecd) = tmAllCnt(llCurrentRecd).lCost(ilQtr)
            imAllRtg(llCurrentRecd) = tmAllCnt(llCurrentRecd).iRtg(ilQtr)
            '6-2-04 determine which pop to use (0 or the number retrieved), not based on different books
            If llPop < 0 Then
                llPop = tmAllCnt(llCurrentRecd).lPop(ilQtr)
            ElseIf llPop <> tmAllCnt(llCurrentRecd).lPop(ilQtr) And llPop <> 0 Then
                llPop = 0               'varying populations
            End If
            'If ilBookNameCode <> tmAllCnt(ilCurrentRecd).iDnfCode Then      'if book name codes are different across vehicles, let Research
                                                                            'routine handle the population
            '    llPop = 0
            'End If
            'ilUpper = ilUpper + 1

        Next llCurrentRecd
        'Get Research values for the demo with or without remnants (build a record for each)
        'If ilUpper > 0 Then
            'gResearchTotals False, llPop, lmAllCost(), imAllRtg(), lmAllGrimp(), lmAllGrp(), lRate, iAvgRtg, lGrimp, lGRP, lCPP, lCPM
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", False, llPop, lmAllCost(), lmAllGrimp(), lmAllGrp(), llLnSpots, lRate, iAvgRtg, lGrImp, lGRP, lCPP, lCPM, lAvgAud
            gResearchTotals "1", False, llPop, lmAllCost(), lmAllGrimp(), lmAllGrp(), llLnSpots, dRate, iAvgRtg, lGrImp, lGRP, lCPP, lCPM, lAvgAud 'TTP 10439 - Rerate 21,000,000
            'create grf record for each demo gathered
            If RptSelCp!rbcSelCSelect(0).Value Then
                tmGrf.lDollars(ilQtr - 1) = lCPP
            Else
                tmGrf.lDollars(ilQtr - 1) = lCPM
            End If
        'End If
    Next ilQtr
    If lCPP <> 0 And lCPM <> 0 Then     'only write recd to disk to a CPP and CPM exist
        'tmGrf.iGenTime(0) = igNowTime(0)
        'tmGrf.iGenTime(1) = igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.iStartDate(0) = ilEffDate(0)
        tmGrf.iStartDate(1) = ilEffDate(1)
        tmGrf.iYear = igYear
        tmGrf.iCode2 = Val(RptSelCp!edcSelCFrom1.Text)      '#quarters
        If RptSelCp!rbcSelCSelect(0).Value Then            'CPP
            tmGrf.sDateType = "P"
        Else
            tmGrf.sDateType = "M"                          'CPM
        End If
        tmGrf.sBktType = "C"                                'assume corp
        If RptSelCp!rbcSelCInclude(1).Value Then
            tmGrf.sBktType = "S"
        End If
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    End If
End Sub
