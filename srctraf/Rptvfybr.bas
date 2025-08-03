Attribute VB_Name = "RPTVFYBR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfybr.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelBR.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Public tgRptSelBRAgencyCode() As SORTCODE
Public tgRptSelBRSalespersonCode() As SORTCODE
Public tgRptSelBRAdvertiserCode() As SORTCODE
Public tgRptSelBRNameCode() As SORTCODE
Public tgRptSelBRBudgetCode() As SORTCODE
Public tgRptSelBRDemoCode() As SORTCODE
'
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenReportBr                              *
'*                                                             *
'*             Created:6/16/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                             *
'*         Comments: Formula setups for Crystal                *
'*                                                             *
'*          Return : 0 =  either error in input, stay in       *
'*                   -1 = error in Crystal, return to          *
'*                        calling program                      *
''*                       failure of gSetformula or another    *
'*                    1 = Crystal successfully completed       *
'*                    2 = successful Bridge                    *
'***************************************************************
Function gCmcGenBr(ilListIndex As Integer, ilGenShiftKey As Integer) As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slTime As String
    Dim slUserID As String     '2-16-13 User ID to filter along with gendate and time
    gCmcGenBr = 0
    slSelection = ""
    gUnpackDate igNowDate(0), igNowDate(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    'gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
    'gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime    '10-20-01
    '9-14-09 do not use ignowtime which is time to nearest seconds;
    'use the time which obtained milliseconds (timegettime)
    slTime = Trim$(str$(lgNowTime))
    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    'slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    'slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & slTime          '9-14-09
    slSelection = slSelection & " And {CBF_Contract_BR.cbfGenTime} = " & slTime          '11-18-11
    
    '2-16-13 filter on gen date, time & urfcode
    slUserID = Trim$(str(tgUrf(0).iCode))
    slSelection = slSelection & " and {CBF_Contract_BR.cbfurfCode} = " & slUserID
    
'   Columns indicate the job # for the specific task(left column).  The numbers below each job# indicate the flag in
'   cbfExtra2Byte to filter records in the .rpt
'
'   Task                    igJobRptNo#1    igJobRptNo2     igJobRptNo3-->4     igJobRptNo4-->5     igJobRptNo5-->3
'                           Detail          NTR             ResearchSum         Billing Sum         CPM Line IDs
'                                                                               0 for filter is used as a record in main .rpt to link the subreports to
'   Detail/Sum,Combine
'     AT & NTR w/Resrch     0               4               <> 4,5,-1,8         0
'   Detail Only Combine
'     AT & NTR w/Resrch     0
'   Summ Only Combine
'     AT & NTR w/Resrch                     4              <> 4,5,-1,8         0
'

    '12-21-20   Add CPM to show all line IDs, CPM Research Summary, and CPM billing summary
'   CPM Line IDs will add
'   another job (#5)

    
'   Detail/Sum,Combine
'     AT & NTR wo/Resrch     0              4                                   0
'   Detail Only Combine
'     AT & NTR wo/Resrch     0
'   Summ Only Combine
'     AT & NTR wo/Resrch                    4                                   0
'


'   These versions do not have a billing summary for NTR items; they are only shown on combined
'   Detail/Sum,Separate
'     AT & NTR w/Resrch     0               4               <> 4,5,-1,8,9    0
'   Detail Only Separate
'     AT & NTR w/Resrch     0
'   Summ Only Separate
'     AT & NTR w/Resrch                     4               <> 4,5,-1,8,9    0

'   Detail/Sum,Separate
'     AT & NTR wo/Resrch     0              4                               0
'   Detail Only Separate
'     AT & NTR wo/Resrch     0
'   Summ Only Separate
'     AT & NTR wo/Resrch                    4                               0

'       CBF Record types :  cbfExtra2Byte = 0:  detail
'                                           2 = vehicle summary
'                                           3 = contract tots
'                                           4 = NTR
'                                           5 = sports (games)
'                                           6 = for Installment contracts only for those vehicles that
'                                               were not on the air time schedule (i.e. NTRs).  These
'                                               vehicles need to show on the Installment Summary version
'                                           -1 = Key record for Insertion Order
'                                           8 = NTR billing summary (non-installment)
'                                           9 = CPM detail IDs
'                                           10 = CPM vehicle summary (for Research page)
'                                           11 = CPM billing summary by vehicle

'        If (igJobRptNo = 1) Or (igJobRptNo = 3 And igDetSumBoth = 1) Then 'Detail or sumary pass & user requested summary
'       12-22-20 need to insert CPM job as # 3, others move up one #
        If (igJobRptNo = 1) Or (igJobRptNo = 4 And igDetSumBoth = 1) Then 'Detail or sumary pass & user requested summary

            If igJobRptNo = 1 Then              'for pass 1 of BR (Detail), filter out summary records
                slSelection = slSelection & " AND ({CBF_Contract_BR.cbfExtra2Byte} = 0)"
            Else
                'Both Research version and the billing summary come thru here
                If igBRSumZer Then              'its the billing summary
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = -1 or {CBF_Contract_BR.cbfExtra2Byte} = 0)  "        'send only the record so that the airtime and NTR subreports can link to it,
                                        '7-12-10 plus the detail records which contain the monthly summaries
                Else
 '                   slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} <> 4 and {CBF_Contract_BR.cbfExtra2Byte} <> 5 and {CBF_Contract_BR.cbfExtra2Byte} <> -1 and {CBF_Contract_BR.cbfExtra2Byte} <> 8 and {CBF_Contract_BR.cbfExtra2Byte} <> 9 and {CBF_Contract_BR.cbfExtra2Byte} <> 10)"        ' send  detail records & total research data reqd except ntr & sports comments & CPM (12-23-20)
                    'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0)" 'TTP 10591 - Proposal Snapshot Summary Report Displaying No Research
                    'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0 or {CBF_Contract_BR.cbfExtra2Byte} = 2)" 'TTP 10591 - Proposal Snapshot Summary Report Displaying No Research
                    slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0 or {CBF_Contract_BR.cbfExtra2Byte} = 2  or {CBF_Contract_BR.cbfExtra2Byte} = 3)" 'Fix 10591 - per Jason Email v81 TTP 10591 testing B1502 Fri 11/18/22 11:10 AM
                End If
            End If
        Else                                'BR summary
            If igJobRptNo = 2 Then              'NTR summary
                slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 4)"        'send only NTR records
            ElseIf igJobRptNo = 3 Then         'CPM line IDs
                slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 9)"        'send only CPM records
            Else
                If sgInclResearch <> "Y" Then                   'dont include research
                    'billing summary
                    slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0) "        'send only the record so that the airtime and NTR subreports can link to it
                Else
                    'Research is included, need to know which version of the summary will be printed
                    'Both Research version and the billing summary come thru here
                    If igBRSumZer Then              'its the billing summary
                        'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = -1 or {CBF_Contract_BR.cbfExtra2Byte} = 0)  "        'send only the record so that the airtime and NTR subreports can link to it,
                        '5-2-14 remove including cbfextra2byte = -1
                        'billing summary
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0) "        'send only the record so that the airtime and NTR subreports can link to it
                                   '7-12-10 plus the detail records which contain the monthly summaries
                    Else
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} <> 4 and {CBF_Contract_BR.cbfExtra2Byte} <> 5 and {CBF_Contract_BR.cbfExtra2Byte} <> -1 and {CBF_Contract_BR.cbfExtra2Byte} <> 8 and {CBF_Contract_BR.cbfExtra2Byte} <> 9 and {CBF_Contract_BR.cbfExtra2Byte} <> 10 and  {CBF_Contract_BR.cbfExtra2Byte} <> 11)"        'send  detail records & total research data reqd except ntr & sports comments & CPM (12-23-20)
                    End If
                End If
            End If
        End If
        
        'TTP 10537 - NTR Rates on Proposals and Contract
        'If sgInclResearch <> "Y" Or igJobRptNo = 3 Or igJobRptNo = 4 Then                   'exclude research, do form that may or maynot include rates.  if CPM version, may or may not need to show the rates
        If sgInclResearch <> "Y" Or igJobRptNo = 2 Or igJobRptNo = 3 Or igJobRptNo = 4 Then                    'exclude research, do form that may or maynot include rates.  if CPM version, may or may not need to show the rates
            If sgInclRates = "Y" Then            'include rates
                If Not gSetFormula("ShowRates", "'Y'") Then  'include rates with research (which includes weekly totals)
                    gCmcGenBr = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowRates", "'N'") Then
                    gCmcGenBr = -1
                    Exit Function
                End If
            End If
        End If
        If sgInclProof = "Y" Then                'proof?
            If Not gSetFormula("Proof", "'Y'") Then     'Requesting hidden lines
                gCmcGenBr = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("Proof", "'N'") Then     'Normal BR
                gCmcGenBr = -1
                Exit Function
            End If
        End If
                
        'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
        If bgSuppressNTRDetails = True Then
            If Not gSetFormula("SuppressNTRDetail", "true", False) Then
                'Not going to fail if the formula is not present to set
                'gCmcGenBr = -1
                'Exit Function
            End If
        Else
            If Not gSetFormula("SuppressNTRDetail", "false", False) Then
                'Not going to fail if the formula is not present to set
                'gCmcGenBr = -1
                'Exit Function
            End If
        End If

        'on proposals, show the agy comm and net, previously only showed the gross $
        If sgShowNetOnProps = "Y" Then       'show net amt on props
            If Not gSetFormula("UserWantsNet", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                gCmcGenBr = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UserWantsNet", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                gCmcGenBr = -1
                Exit Function
            End If
        End If
        
        If sgShowProdProt = "Y" Then       '8-25-15 show product protection
            If Not gSetFormula("ShowProdProt", "'Y'") Then
                gCmcGenBr = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowProdProt", "'N'") Then
                gCmcGenBr = -1
                Exit Function
            End If
        End If
      
        '4-27-12 all insertion orders and contracts to test for vehicle word wrap
        If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
            If Not gSetFormula("WordWrapVehicle", "'Y'") Then
                gCmcGenBr = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("WordWrapVehicle", "'N'") Then
                gCmcGenBr = -1
                Exit Function
            End If
        End If


        If igJobRptNo <> 1 Then     '2-13-04 make sure all summaries get the splits if requested
            '2-2-10  Summary, is the version to merge the NTR billing?
'            If igBRSumZer Or igBRSum Then           'for billing summary or research summary, there is option to show NTR totals with air time
            If igJobRptNo = 5 Or igJobRptNo = 4 Then                      'billing summary or research summary
                If sgInclNTRBillSummary = "Y" Then
                    If Not gSetFormula("ShowNTRSummary", "'Y'") Then
                        gCmcGenBr = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowNTRSummary", "'N'") Then
                        gCmcGenBr = -1
                        Exit Function
                    End If
                End If
                
            End If
            
            If sgInclSplits = "Y" Then           'show Slsp Commission Splits
                If Not gSetFormula("ShowSplits", "'Y'") Then  'show the slsp comm splits on summary
                    gCmcGenBr = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowSplits", "'N'") Then  'show the slsp comm splits on summary
                    gCmcGenBr = -1
                    Exit Function
                End If
            End If
            
            '10-31-19 show EDI agy client and adv product code.  show them on summary versions only
            If ((Asc(tgSaf(0).sFeatures6) And EDIAGYCODES) = EDIAGYCODES) Then    'using agy client code?
                If Not gSetFormula("ShowEDICodes", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                    gCmcGenBr = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowEDICodes", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                    gCmcGenBr = -1
                    Exit Function
                End If
            End If

        Else                'detail
            If ((Asc(tgSpf.sUsingFeatures8) And SHOWCMMTONDETAILPAGE) = SHOWCMMTONDETAILPAGE) Then      'yes, show comments on detail
                If Not gSetFormula("ShowComments", "'Y'") Then  'show comments:  other,system site, chg reason, cancellations reason)
                    gCmcGenBr = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowComments", "'N'") Then  'hide comments:  other,system site, chg reason, cancellations reason)
                    gCmcGenBr = -1
                    Exit Function
                End If
            End If
            
            '4-30-13  Show flight rates on packages (vs just show the total line
            If (Asc(tgSpf.sUsingFeatures10) And PKGLNRATEONBR) = PKGLNRATEONBR Then     'show flight rates with package lines
                If Not gSetFormula("ShowRateForPkgFlight", "'Y'") Then
                    gCmcGenBr = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowRateForPkgFlight", "'N'") Then
                    gCmcGenBr = -1
                    Exit Function
                End If
            End If
        End If
        
        '8410 - Exclude some special installment package or hidden records
        If ilListIndex = 0 And igJobRptNo = 1 Then
            slSelection = slSelection & " and ({CBF_Contract_BR.cbfLineType} <> 'X' And {CBF_Contract_BR.cbfLineType} <> 'Y')"
        End If
        If ilListIndex = 0 And igJobRptNo = 4 Then
            slSelection = slSelection & " and ({CBF_Contract_BR.cbfLineType} <> 'X')  "
        End If
        If ilListIndex = 0 And igJobRptNo = 5 Then
            slSelection = slSelection & " and ({CBF_Contract_BR.cbfLineType} <> 'X')  "
        End If
        If Not gSetSelection(slSelection) Then
            gCmcGenBr = -1
            Exit Function
        End If
    gCmcGenBr = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportBr                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportBr() As Integer
Dim blCustomLogo As Boolean

    If Not igUsingCrystal Then
        gGenReportBr = True
        Exit Function
    End If

    blCustomLogo = True                             'proposals/contracts use the sales source to determine the logo to print
    igBRSumZer = False                              'summary - monthly billing
    igBRSum = False                                 'summary with research
    If igJobRptNo = 1 Then                      'detail pass
        If sgInclResearch = "Y" Then        'If including research, show wide with everything if including rates
            If sgInclRates = "Y" Then             'with rates
                If Not gOpenPrtJob("BR.Rpt", , blCustomLogo) Then
                    gGenReportBr = False
                    Exit Function
                End If
            Else
                If Not gOpenPrtJob("BRNoRate.Rpt", , blCustomLogo) Then 'with research but without CPP/CPM, and any other rates
                    gGenReportBr = False
                    Exit Function
                End If
            End If
        Else                                    'exclude research, prices optional
            If Not gOpenPrtJob("BRZer.Rpt", , blCustomLogo) Then
                gGenReportBr = False
                Exit Function
            End If
        End If
'    ElseIf igJobRptNo = 3 Then                         'summary pass
        ElseIf igJobRptNo = 4 Then                      '12-22-20 job #3 added; others adjusted, summary pass

        If sgInclResearch = "Y" Then        'include research, assume to show prices
            If sgInclRates = "Y" Then      'with rates   & research
                igBRSum = True                                  'need to know which summary version is being processed
                If Not gOpenPrtJob("BRSum.Rpt", , blCustomLogo) Then
                    gGenReportBr = False
                    Exit Function
                End If
            Else
                igBRSum = True
'                If Not gOpenPrtJob("BRSumnor.Rpt", , blCustomLogo) Then     'Research without rates
                If Not gOpenPrtJob("Brsum.Rpt", , blCustomLogo) Then     'Research without rates
                    gGenReportBr = False
                    Exit Function
                End If
            End If
        ElseIf sgInclRates = "Y" Then
            igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
            If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomLogo) Then 'exclude research, prices optional
                gGenReportBr = False
                Exit Function
            End If
        Else
            igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
            If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomLogo) Then 'exclude research, prices optional
                gGenReportBr = False
                Exit Function
            End If
        End If
'    ElseIf igJobRptNo = 4 Then          'billing summary
     ElseIf igJobRptNo = 5 Then          '12-22-20 job #3 added, others adjusted
        igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
        If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomLogo) Then 'exclude research, prices optional
            gGenReportBr = False
            Exit Function
        End If
    ElseIf igJobRptNo = 2 Then          'NTR

        If Not gOpenPrtJob("BRNTR.Rpt", , blCustomLogo) Then 'exclude research, prices optional
            gGenReportBr = False
            Exit Function
        End If
     ElseIf igJobRptNo = 3 Then          '12-22-20 this CPM job added

        If Not gOpenPrtJob("BRCPM.Rpt", , blCustomLogo) Then 'exclude research, prices optional
            gGenReportBr = False
            Exit Function
        End If
    End If

    gGenReportBr = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReset                   *
'*                                                     *
'*             Created:1/31/96       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reset controls                 *
'*                                                     *
'*******************************************************
Sub mReset()
    igGenRpt = False
    RptSelBR!frcOutput.Enabled = igOutput
    RptSelBR!frcCopies.Enabled = igCopies
    'RptSelBR!frcWhen.Enabled = igWhen
    RptSelBR!frcFile.Enabled = igFile
    RptSelBR!frcOption.Enabled = igOption
    'RptSelBR!frcRptType.Enabled = igReportType
    Beep
End Sub
