Attribute VB_Name = "RPTCRPA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrpa.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCRPA.Bas
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
'Dim tmLRch() As RESEARCHLIST
Dim tmCbf As CBF                  'BR prepass file
Dim hmCbf As Integer
Dim imCbfRecLen As Integer        'BR record length
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmSlf As Integer            'Slsp file handle
'  Rating Book File
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length
'
' Copyright 2000 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' Programmer: D. Smith
' Date: 08/16/00
' Name: gSalesPricingAnalysisGen
' Purpose: Generate prepass file for Crystal - comparison of actual sales dollars
'          against ratecard dollars
'
'
Sub gSalesPricingAnalysisGen()
Dim ilRet As Integer
Dim ilCurrentRecd As Integer            'number of contracts processed so far
Dim ilLoop As Integer                   'temp loop variable
Dim ilLoop2 As Integer                  'temp loop variable
Dim llContrCode As Long                 'Contr ID to process
Dim slActStartDate As String            'Contract active start date
Dim slActEndDate As String              'contract active end date
Dim slEntStartDate As String            'Contract entered start date
Dim slEntEndDate As String              'contract entered end date
Dim ilClf As Integer                    'loop for lines
Dim ilCff As Integer                    'loop for flights
Dim slStr As String                     'temp string for conversions
Dim llDate As Long                      'temp serial date
Dim llDate2 As Long
Dim ilDay As Integer
Dim slCntrType As String                'valid contract types (per inq, direct respon, etc) to retrieve
Dim slCntrStatus As String              'valid contr status (working, complete, etc) to retrieve
Dim ilHOState As Integer                'which type of Holds Orders to retrieved (internally WCI)
Dim llFltStart As Long
Dim llFltEnd As Long
Dim ilSpots As Integer
ReDim ilInputDays(0 To 6) As Integer    'valid days of the week for audience retrieval
Dim llTemp As Long
Dim ilTotLnSpts As Integer              'sum of spots per line
Dim llTotLnGross As Long                'sum of gross $ per line
Dim slNameCode As String
Dim ilIdx As Integer
Dim ilFirstPass As Integer
Dim slTempDays As String
Dim slDysTms As String
Dim slDays As String
Dim ilShowOVDays As Integer
Dim ilShowOVTimes As Integer
Dim slStartTime As String
Dim slEndTime As String
Dim hlRdf As Integer                    'supports day parts
Dim tlRdf As RDF                        'supports day parts
Dim ilRdfRecLen As Integer              'supports day parts
Dim tlRdfSrchKey As INTKEY0             'local type to support day parts
Dim slSlsCode() As String               'array of selected salespeople
Dim ilSlsIdx As Integer                 'index into array of salespeople
Dim ilContinue As Integer               'used for test against user selected contract types
Dim slActStart As String                'used in building the date formula for Crystal
Dim slActEnd As String                  'used in building the date formula for Crystal
Dim slEntStart As String                'used in building the date formula for Crystal
Dim slEntEnd As String                  'used in building the date formula for Crystal
Dim ilInsertOK As Integer
Dim llTotalLineRCPrice As Long
Dim llTotalLineActPrice As Long
Dim llTotalLineRCSpots As Long

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    ReDim tgClfPA(0 To 0) As CLFLIST
    tgClfPA(0).iStatus = -1 'Not Used
    tgClfPA(0).lRecPos = 0
    tgClfPA(0).iFirstCff = -1
    ReDim tgCffPA(0 To 0) As CFFLIST
    tgCffPA(0).iStatus = -1 'Not Used
    tgCffPA(0).lRecPos = 0
    tgCffPA(0).iNextCff = -1
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
    imClfRecLen = Len(tgClfPA(0).ClfRec)
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
    imCffRecLen = Len(tgCffPA(0).CffRec)
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
    hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCbf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmDrf
        btrDestroy hmCbf
        btrDestroy hmUrf
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
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmDrf
        btrDestroy hmCbf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlRdf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hlRdf
        btrDestroy hmDrf
        btrDestroy hmCbf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    slCntrType = gBuildCntTypes()
    slCntrStatus = "HOGN"          'only get holds and orders
    ilHOState = 2                  'get latest orders and revisions
    'Test Use Only
    'slActStartDate = "8/1/00"
    'slActEndDate = "8/31/00"
    'slEntStartDate = "8/1/00"
    'slEntEndDate = "8/15/00"
    'User entered date values
'    slActStartDate = RptSelPA!edcSelCFrom.Text
'    slActEndDate = RptSelPA!edcSelCFrom1.Text
'    slEntStartDate = RptSelPA!edcSelCTo.Text
'    slEntEndDate = RptSelPA!edcSelCTo1.Text

'   12-16-19 change dates to use csi calendar control
    slActStartDate = RptSelPA!CSI_CalFrom.Text
    slActEndDate = RptSelPA!CSI_CalTo.Text
    slEntStartDate = RptSelPA!CSI_CalFrom1.Text
    slEntEndDate = RptSelPA!CSI_CalTo1.Text


    'gather all the salespeople that were selected into the slSlsCode array
    ilSlsIdx = 0
    ReDim slSlsCode(0 To 0)
    For ilLoop = 0 To RptSelPA!lbcSelection(0).ListCount - 1 Step 1
        If RptSelPA!lbcSelection(0).Selected(ilLoop) Then
            slNameCode = tgSalesperson(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slSlsCode(ilSlsIdx))
            ilSlsIdx = ilSlsIdx + 1
            ReDim Preserve slSlsCode(0 To ilSlsIdx)
        End If
    Next ilLoop
    ilRet = gCntrForActiveOHD(RptSelPA, slActStartDate, slActEndDate, slEntStartDate, slEntEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
    'Now loop throgh the found contracts that meet the date criteria
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        llTemp = tlChfAdvtExt(ilCurrentRecd).lCntrNo
    If tlChfAdvtExt(ilCurrentRecd).lCntrNo = 1219 Then
        ilRet = ilRet
    End If
        'loop through the 9 buckets - each contract has for potential for 9 salespeople
        'We decided not to show the splits so I set the loop to (0 to 0)
        'To see the splits set the loop to (0 to 9)
        For ilIdx = 0 To 0 Step 1
            'Loop testing the selected salespeople against a value in the given bucket above
            For ilSlsIdx = 0 To UBound(slSlsCode) - 1 Step 1
                If Val(slSlsCode(ilSlsIdx)) = tlChfAdvtExt(ilCurrentRecd).iSlfCode(ilIdx) Then
                    'Retrieve the contract, schedule lines and flights
                    llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfPA, tgClfPA(), tgCffPA())
                    'test to see if it meets the check boxes the user selceted and check trade
                    'first see what the user slected;
                    ilContinue = False
                    If ((RptSelPA!ckcSelC5(0).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "H") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(1).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "O") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(2).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "C") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(3).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "V") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(4).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "T") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(5).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "R") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(6).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "Q") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(7).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "S") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(8).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "M") Then
                        ilContinue = True
                    End If
                    If ((RptSelPA!ckcSelC5(9).Value = vbChecked) And tlChfAdvtExt(ilCurrentRecd).sType = "H") Then
                        ilContinue = True
                    End If
                    'We stop here if we didn't meet any of the above conditions - contract types
                    If ilContinue Then
                        If (tgChfPA.iPctTrade <> 100) Then    'get a contract and test for printables,
            
                            'now process looping schedule line - tgClfPA fills in upper and lower bounds
                            For ilClf = LBound(tgClfPA) To UBound(tgClfPA) - 1 Step 1
                                ilTotLnSpts = 0                 'init total # spots per line
                                llTotLnGross = 0                'init total dollars this line
                                tmClf = tgClfPA(ilClf).ClfRec
                                ilInsertOK = False
                                '3-19-14 determine cancel before start
                                gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slStr
                                llFltStart = gDateValue(slStr)
                                gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slStr
                                llFltEnd = gDateValue(slStr)
                                llTotalLineRCPrice = 0
                                llTotalLineActPrice = 0
                                llTotalLineRCSpots = 0
                                If (tmClf.sType = "H" Or tmClf.sType = "S") And (llFltEnd >= llFltStart) Then      'only makes sense to do cpp cpm on standard or hidden lines
                                    ilInsertOK = True
                                    ilCff = tgClfPA(ilClf).iFirstCff
                                    ilFirstPass = True
                                    Do While ilCff <> -1
                                        tmCff = tgCffPA(ilCff).CffRec
                                        For ilLoop2 = 0 To 6                 'init all days to not airing, setup for research results later
                                            ilInputDays(ilLoop2) = False
                                        Next ilLoop2
                                        gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                                        llFltStart = gDateValue(slStr)
                                        'backup start date to Monday
                                        ilLoop = gWeekDayLong(llFltStart)
                                        Do While ilLoop <> 0
                                            llFltStart = llFltStart - 1
                                            ilLoop = gWeekDayLong(llFltStart)
                                        Loop
                                        If ilFirstPass = True Then 'Get the first start date of the line
                                            tmCbf.iStartDate(0) = tmCff.iStartDate(0)
                                            tmCbf.iStartDate(1) = tmCff.iStartDate(1)
                                            ilFirstPass = False
                                        End If
                                        gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                                        llFltEnd = gDateValue(slStr)
                                        'Loop thru the flight by week and build the number of spots for each week
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
                             
                                            Else 'daily
                                                If ilLoop + 6 < llFltEnd Then     'we have a whole week
                                                    ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                                    For ilDay = 0 To 6 Step 1
                                                        If tmCff.iDay(ilDay) > 0 Then
                                                            ilInputDays(ilDay) = True
                                                        End If
                                                    Next ilDay
                                                Else 'do partial week
                                                    For llDate = llDate2 To llFltEnd Step 1
                                                        ilDay = gWeekDayLong(llDate)
                                                        ilSpots = ilSpots + tmCff.iDay(ilDay)
                                                        If tmCff.iDay(ilDay) > 0 Then
                                                            ilInputDays(ilDay) = True
                                                        End If
                                                    Next llDate
                                                End If
                                            End If
                                            ilTotLnSpts = ilTotLnSpts + ilSpots     '4-10-14 accum spots per flight for line.  this line was in wrong place, see instr below
                                            llTotalLineActPrice = llTotalLineActPrice + (ilSpots * tmCff.lActPrice)
                                            If tmCff.lPropPrice > 0 Then
                                                llTotalLineRCPrice = llTotalLineRCPrice + (ilSpots * tmCff.lPropPrice)
                                                llTotalLineRCSpots = llTotalLineRCSpots + ilSpots
                                            Else
                                                llTotalLineRCSpots = llTotalLineRCSpots
                                            End If
                                        Next llDate2
                                        'ilTotLnSpts = ilTotLnSpts + ilSpots
                                        ilCff = tgCffPA(ilCff).iNextCff         'get next flight record from mem
                                        'gDayNames returns the days of the week from the Flight Record: i.e. MTuWeThFr or SaSu, etc
                                        slTempDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slStr)
                                        slDysTms = ""

                                        'Retrieve the days this flight is to air, strip out the commas & blanks from the text string
                                        For ilLoop = 1 To Len(slTempDays) Step 1
                                            slDays = Mid$(slTempDays, ilLoop, 1)
                                            If slDays <> "" And slDays <> "," Then
                                                slDysTms = Trim$(slDysTms) & Trim$(slDays)
                                            End If
                                        Next ilLoop

                                        ilRdfRecLen = Len(tlRdf)
                                        tlRdfSrchKey.iCode = tmClf.iRdfCode
                                        ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tmCbf.sDysTms = "Missing DP"
                                        End If

                                        'read in RDF (Daypart record) from CLFrdfcode.
                                        'this for loop compares the days of the week defined in the daypart against the days of the week in the flight
                                        'If one of the CFF days of the week is different than that defined in the daypart, it sets a flag for testing later
                                        For ilLoop = 0 To 6 Step 1       'see if there are override days from the flights compared to DP
                                            'If tlRdf.sWkDays(7, ilLoop + 1) = "Y" Then       'is DP a valid day
                                            If tlRdf.sWkDays(6, ilLoop) = "Y" Then       'is DP a valid day
                                                If tmCff.iDay(ilLoop) = 0 Then
                                                    ilShowOVDays = True
                                                    Exit For
                                                Else
                                                    ilShowOVDays = False
                                                End If
                                            End If
                                        Next ilLoop

                                        'see if override times compared to DP
                                        'Similarly, the daypart times are compared against the line.  If theres a value in StartTime or End Time, it denotes
                                        'an override time; otherwise it should be zero and use just the times from the daypart record
                                        'If theres an override time, a Time Override flag is set for later testing
                                        ilShowOVTimes = False
                                        If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                                            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                                            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                                            ilShowOVTimes = True
                                        Else
                                            For ilLoop = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1
                                                If (tlRdf.iStartTime(0, ilLoop) <> 1) Or (tlRdf.iStartTime(1, ilLoop) <> 0) Then
                                                    gUnpackTime tlRdf.iStartTime(0, ilLoop), tlRdf.iStartTime(1, ilLoop), "A", "1", slStartTime
                                                    gUnpackTime tlRdf.iEndTime(0, ilLoop), tlRdf.iEndTime(1, ilLoop), "A", "1", slEndTime
                                                    Exit For
                                                End If
                                            Next ilLoop
                                        End If

                                        'Determine whether there was an override detected:  If so, use the days of the week from CFF that has been
                                        'retrieved and parsed and concatenate it with the times stored in the line
                                        If (ilShowOVDays Or ilShowOVTimes) Then
                                            tmCbf.sDysTms = slDysTms & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
                                        Else
                                            tmCbf.sDysTms = Trim$(tlRdf.sName)
                                        End If
                                    Loop                                            'while ilcff <> -1
                                    tmCbf.iEndDate(0) = tmCff.iEndDate(0)
                                    tmCbf.iEndDate(1) = tmCff.iEndDate(1)
                                End If                                              'line outside range of requested dates
                                If ilInsertOK = True Then
                                    'stuff the info into the CBF records
                                    tmCbf.iLen = False 'initial this field - it may not go through the code in the program
                                    tmCbf.iLen = ilShowOVDays
                                    'tmCbf.iGenTime(0) = igNowTime(0)
                                    'tmCbf.iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmCbf.lGenTime = lgNowTime
                                    tmCbf.iGenDate(0) = igNowDate(0)
                                    tmCbf.iGenDate(1) = igNowDate(1)
                                    tmCbf.iSlfCode = Val(slSlsCode(ilSlsIdx))
                                    tmCbf.lContrNo = tlChfAdvtExt(ilCurrentRecd).lCntrNo
                                    tmCbf.lLineNo = tmCff.iClfLine
                                    tmCbf.iVefCode = tmClf.iVefCode
                                    tmCbf.iAdfCode = tlChfAdvtExt(ilCurrentRecd).iAdfCode
                                    tmCbf.sProduct = tlChfAdvtExt(ilCurrentRecd).sProduct
                                    tmCbf.iMnfGroup = tlChfAdvtExt(ilCurrentRecd).iMnfDemo0
                                    tmCbf.lCurrModSpots = ilTotLnSpts

                                    'If the Rate or the Actual price is zero then omit it form the report
                                    'If (tmCff.lPropPrice <> 0) And (tmCff.lActPrice <> 0) Then
                                        
                                        'tmCbf.lRate = tmCff.lPropPrice
                                        If llTotalLineRCSpots > 0 Then
                                            tmCbf.lRate = (llTotalLineRCPrice) \ llTotalLineRCSpots                  'avg rate card price
                                        Else
                                            tmCbf.lRate = 0
                                        End If
                                        
                                        'tmCff.lActPrice = tmCff.lActPrice / 100   'drop the two lsd's - the cents
                                        If ilTotLnSpts > 0 Then
                                            'tmCbf.lMonth(3) = (llTotalLineActPrice / ilTotLnSpts) \ 100             'avg actual price
                                            tmCbf.lMonth(2) = (llTotalLineActPrice / ilTotLnSpts) \ 100             'avg actual price
                                        Else
                                            'tmCbf.lMonth(3) = 0
                                            tmCbf.lMonth(2) = 0
                                        End If
                                        ''tmCbf.lMonth(3) = tmCff.lActPrice
                                        
                                        ''tmCbf.lMonth(1) = (tmCff.lActPrice - tmCff.lPropPrice)
                                        'tmCbf.lMonth(1) = tmCbf.lMonth(3) - tmCbf.lRate                         'diff of actual price & rate card price
                                        tmCbf.lMonth(0) = tmCbf.lMonth(2) - tmCbf.lRate                         'diff of actual price & rate card price
                                        
                                       '' tmCbf.lMonth(2) = ilTotLnSpts * (tmCff.lActPrice - tmCff.lPropPrice)
                                        'tmCbf.lMonth(2) = ilTotLnSpts * (tmCbf.lMonth(3) - tmCbf.lRate)
                                        tmCbf.lMonth(1) = ilTotLnSpts * (tmCbf.lMonth(2) - tmCbf.lRate)
                                        ''If the Actual - Proposed equals zero then don't divide
                                        'If tmCbf.lMonth(1) = 0 Then
                                        '    tmCbf.lMonth(4) = 0
                                        'Else
                                        '    'If tmCff.lPropPrice <> 0 Then
                                        '    If tmCbf.lRate <> 0 Then
                                        '        'tmCbf.lMonth(4) = ((tmCbf.lMonth(1) * 100) / tmCff.lPropPrice)
                                        '        tmCbf.lMonth(4) = ((tmCbf.lMonth(1) * 100) / tmCbf.lRate)
                                        '    Else
                                        '        tmCbf.lMonth(4) = 0
                                        '    End If
                                        'End If
                                        
                                        If tmCbf.lMonth(0) = 0 Then
                                            tmCbf.lMonth(3) = 0
                                        Else
                                            'If tmCff.lPropPrice <> 0 Then
                                            If tmCbf.lRate <> 0 Then
                                                'tmCbf.lMonth(4) = ((tmCbf.lMonth(1) * 100) / tmCff.lPropPrice)
                                                tmCbf.lMonth(3) = ((tmCbf.lMonth(0) * 100) / tmCbf.lRate)
                                            Else
                                                tmCbf.lMonth(3) = 0
                                            End If
                                        End If
                                        'write the record out - its a keeper
                                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                    'End If
                                End If
                            Next ilClf  'get next line
                        End If
                    End If
                End If
            Next ilSlsIdx     'array of selected sales people
        Next ilIdx            'nine salespeople buckets on contract
    Next ilCurrentRecd        'get another cnt

    'If the user doesn't enter a given date we show "all dates" in the report banner
    If slActStartDate = "" Then
        slActStart = "all dates"
    Else
        slActStart = slActStartDate
    End If
    If slActEndDate = "" Then
        slActEnd = "all dates"
    Else
        slActEnd = slActEndDate
    End If
    If slEntStartDate = "" Then
        slEntStart = "all dates"
    Else
        slEntStart = slEntStartDate
    End If

    If slEntEndDate = "" Then
        slEntEnd = "all dates"
    Else
        slEntEnd = slEntEndDate
    End If

    If Not gSetFormula("Formula Dates Active", "'" & "Active " & slActStart & "-" & slActEnd & "'") Then
        Exit Sub
    End If
    If Not gSetFormula("Formula Dates Entered", "'" & "Entered " & slEntStart & "-" & slEntEnd & "'") Then
        Exit Sub
    End If

    'we are out of here!
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hlRdf)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmCbf)
End Sub
