Attribute VB_Name = "RPTVFYIN"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyin.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelIn.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelInAgencyCode() As SORTCODE
'Public sgRptSelInAgencyCodeTag As String
'Public tgRptSelInSalespersonCode() As SORTCODE
'Public sgRptSelInSalespersonCodeTag As String
'Public tgRptSelInAdvertiserCode() As SORTCODE
'Public sgRptSelInAdvertiserCodeTag As String
'Public tgRptSelInNameCode() As SORTCODE
'Public sgRptSelInNameCodeTag As String
'Public tgRptSelInBudgetCode() As SORTCODE
'Public sgRptSelInBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelInDemoCode() As SORTCODE
'Public sgRptSelInDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public sgRnfRptName As String * 3          'Report name from RNF file (L01, L02,... C01, C02,.. ..etc)
'Public igNoCodes As Integer
'Public igcodes() As Integer
'Public sgLogStartDate As String
'Public sgLogNoDays As String
'Public sgLogUserCode As String
'Public sgLogStartTime As String
'Public sgLogEndTime As String
'Public igZones As Integer                   'time zones  (0=all, 1=est, 2=cst, 3=mst, 4=pst)
'Public igRnfCode As Integer
'Public igInvoiceType As Integer     '0=invoice (no spots) - invport1.rpt, 1 = affidavit (no $) invaff.rpt
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
'Public igOutputTo As Integer        '0 = display , 1 = print
''Global spot types for Spots by Advt & spots by Date & Time
''bit selectivity for charged and different types of no charge spots
''bits defined right to left (0 to 9)
'Public Const SPOT_CHARGE = &H1         'charged
'Public Const SPOT_00 = &H2          '0.00
'Public Const SPOT_ADU = &H4         'ADU
'Public Const SPOT_BONUS = &H8       'bonus
'Public Const SPOT_EXTRA = &H10      'Extra
'Public Const SPOT_FILL = &H20       'Fill
'Public Const SPOT_NC = &H40         'no charge
'Public Const SPOT_MG = &H80         'mg
'Public Const SPOT_RECAP = &H100     'recapturable
'Public Const SPOT_SPINOFF = &H200   'spinoff
'Library calendar file- used to obtain post log date status
'
'******************************************************************
'*
'*      Procedure Name:gGenReportIn
'*
'*             Created:6/16/93       By:D. LeVine
'*            Modified:              By:
'*
'*         Comments: Formula setups for Crystal
'*
'*          Return : 0 =  either error in input, stay in
'*                   -1 = error in Crystal, return to
'*                        calling program
''*                       failure of gSetformula or another
'*                    1 = Crystal successfully completed
'*                    2 = successful Bridge
'       3-21-03 Send # of blank lines to place before & after logo
'               to adjust to fit in windowed envelope
'       01-18-07 Implement combined Air Time and NTR invoice
'*****************************************************************
Function gCmcGenIn(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    Dim tlCxf As CXF
    Dim hlCxf As Integer
    Dim ilCxfRecLen As Integer
    Dim tlSrchKey As LONGKEY0
    Dim ilBlanksBeforeLogo As Integer
    Dim ilBlanksAfterLogo As Integer
    Dim slExcludeInvoiceSelection As String 'Fix TTP 10826 / TTP 10813
'   ivrtypes:  1/30/21
'   when combining commercial and NTR
'    0=Detail Air Time
'    2=Subtotal
'    3=Air Time total
'    4=CPM Detail
'    5=CPM Total
'    7=NTR Detail
'    8=NTR Total
'    9=Combination contract Total
'
'   when not combining commercial and NTR
'   2 = a/t subtotal
'   3 = a/t contract total
    
    gCmcGenIn = 0
    slSelection = ""
    gUnpackDate igNowDate(0), igNowDate(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    slSelection = "{IVR_Invoice_Rpt.ivrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({IVR_Invoice_Rpt.ivrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    sgSelection = Trim$(slSelection)            '11-16-16
    sgSelectionToAdd = ""
    
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    'Emailed invoices are not printed/displayed. To print/display an invoice for an agency or direct advertiser that is set to use the PDF email feature, reprint the invoice without the Email checkbox checked on. 
    If bgSendSelevtivePDF = True And bsSelectedEmailInvoices <> "" Then
        'Fix TTP 10826 / TTP 10813
        'slSelection = slSelection & " And NOT({IVR_Invoice_Rpt.ivrInvNo} IN [" & bsSelectedEmailInvoices & "])"
        slExcludeInvoiceSelection = " And NOT({IVR_Invoice_Rpt.ivrInvNo} IN [" & bsSelectedEmailInvoices & "])"
    End If
    
    If igInvoiceType = 1 Or igInvoiceType = 4 Or igInvoiceType = 3 Then       '12-13-02  NTR, take all records
        'Fix TTP 10826 / TTP 10813
        'If Invoice!ckcArchive.Value = vbChecked Then
        If Invoice!ckcArchive.Value = vbChecked And (tgSpfx.iInvExpFeature And INVEXP_SELECTIVEEMAIL) <> INVEXP_SELECTIVEEMAIL Then
            slSelection = slSelection & " and ({IVR_Invoice_Rpt.ivrShowInvType} <> 5) "
            sgSelectionToAdd = ""
        End If
        '2-1-02 treat as aired and as ordered/update aired the same
        'If tgSpf.sInvAirOrder <> "O" Then           'no totals by market, filter the IVRTYPE = 3 total records,
            'otherwise totals by market - filter IVRTYPES 2  (subtotals by market) and IVRTYPE = 3 (invoice total) records
        '    slSelection = slSelection & "And {IVR_Invoice_Rpt.ivrType} = 3"
        'End If
        
    ElseIf igInvoiceType = 2 Then                   'affidavit only
        If igJobRptNo = 2 Then                      'summary pass, additional filter to exclude everything except detailspots
            slSelection = slSelection & "And {IVR_Invoice_Rpt.ivrType} = 0"
        End If
        sgSelection = slSelection               '11-16-16
    
    ElseIf igInvoiceType = 0 Then                'form #1
        If (Asc(tgSpf.sUsingFeatures5) And SUPPRESSTIMEFORM1) = SUPPRESSTIMEFORM1 Then      'suppress air time?
            If Not gSetFormula("ShowAirTime", "'N'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowAirTime", "'Y'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If
        If Not gSetFormula("SAF ISCI Form", "'" & tgSaf(0).sInvISCIForm & "'") Then
            gCmcGenIn = -1
            Exit Function
        End If
        
        '5-24-13 Sort by Payee or Sales source (output only). Only applies to form #1, with combined or not combined
        If tgSaf(0).sInvoiceSort = "S" Then            'sales source
            If Not gSetFormula("UseAsMajorSort", "'S'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else                                        'default sort by payee (inv #)
            If Not gSetFormula("UseAsMajorSort", "'P'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If
        
        '1-13-14 Audio type from line override
        If ((Asc(tgSaf(0).sFeatures1) And SHOWAUDIOTYPEONBR) = SHOWAUDIOTYPEONBR) Then      'yes, show audio type on lines
            If Not gSetFormula("ShowAudioTypeOption", "'Y'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowAudioTypeOption", "'N'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If
        
        If mSameInvFormula() <> 0 Then              '11-26-19, change way error is treated.  10-31-19 option to show edi agy client and adv product codes
            gCmcGenIn = -1
            Exit Function
        End If
        
        'Fix TTP 10826 / TTP 10813
        'If Invoice!ckcArchive.Value = vbChecked Then
        If Invoice!ckcArchive.Value = vbChecked And (tgSpfx.iInvExpFeature And INVEXP_SELECTIVEEMAIL) <> INVEXP_SELECTIVEEMAIL Then
            slSelection = slSelection & " and ({IVR_Invoice_Rpt.ivrShowInvType} <> 5) "
            sgSelectionToAdd = ""
        End If
    
    ElseIf igInvoiceType = 6 Then                       '01-18-07 combined air time & NTR
        If (Asc(tgSpf.sUsingFeatures5) And SUPPRESSTIMEFORM1) = SUPPRESSTIMEFORM1 Then      'suppress air time?
            If Not gSetFormula("ShowAirTime", "'N'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowAirTime", "'Y'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If
        If Not gSetFormula("SAF ISCI Form", "'" & tgSaf(0).sInvISCIForm & "'") Then
            gCmcGenIn = -1
            Exit Function
        End If
        
         '5-24-13 Sort by Payee or Sales source (output only).  Only applies to form #1, with combined or not combined
        If tgSaf(0).sInvoiceSort = "S" Then            'sales source
            If Not gSetFormula("UseAsMajorSort", "'S'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else                                        'default sort by payee (inv #)
            If Not gSetFormula("UseAsMajorSort", "'P'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If
        
        '1-13-14 Audio type from line override
        If ((Asc(tgSaf(0).sFeatures1) And SHOWAUDIOTYPEONBR) = SHOWAUDIOTYPEONBR) Then      'yes, show audio type on lines
            If Not gSetFormula("ShowAudioTypeOption", "'Y'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowAudioTypeOption", "'N'") Then
                gCmcGenIn = -1
                Exit Function
            End If
        End If

        If mSameInvFormula() <> 0 Then            '10-31-19 option to show edi agy client and adv product codes
            gCmcGenIn = -1
            Exit Function
        End If
        
        'slSelection = "{IMR_Invoice_Main_Rpt.imrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = "{IVR_Invoice_Rpt.ivrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        'slSelection = slSelection & " And Round({IMR_Invoice_Main_Rpt.imrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        slSelection = slSelection & " And Round({IVR_Invoice_Rpt.ivrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
'        slSelection = slSelection & " and ( ({IVR_Invoice_Rpt.ivrType} = 3  or  {IVR_Invoice_Rpt.ivrType} = 5 or {IVR_Invoice_Rpt.ivrType} = 6) )"
'       1/30/21 ntr detail and total record types changed to  7 & 8 (from 4 & 5), add cpm detail and total records (types 4 & 5)
        'slSelection = slSelection & " and ( ({IVR_Invoice_Rpt.ivrType} = 3  or  {IVR_Invoice_Rpt.ivrType} = 5 or {IVR_Invoice_Rpt.ivrType} = 9 or {IVR_Invoice_Rpt.ivrType} = 8) )"
        'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
        slSelection = slSelection & " and ( ({IVR_Invoice_Rpt.ivrType} = 3  or  {IVR_Invoice_Rpt.ivrType} = 5 or {IVR_Invoice_Rpt.ivrType} = 7 or {IVR_Invoice_Rpt.ivrType} = 8 or {IVR_Invoice_Rpt.ivrType} = 9 ) )"
        
        'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
        'Emailed invoices are not printed/displayed. To print/display an invoice for an agency or direct advertiser that is set to use the PDF email feature, reprint the invoice without the Email checkbox checked on. 
        If bgSendSelevtivePDF = True And bsSelectedEmailInvoices <> "" Then
            'Add the Inclusion of Invoices
            'Fix TTP 10826 / TTP 10813
            'slSelection = slSelection & " And NOT({IVR_Invoice_Rpt.ivrInvNo} IN [" & bsSelectedEmailInvoices & "])"
            slExcludeInvoiceSelection = " And NOT({IVR_Invoice_Rpt.ivrInvNo} IN [" & bsSelectedEmailInvoices & "])"
        End If
            
        sgSelection = slSelection       'save original selection without the EDI and PDF invoices
        '11-16-16 if archiving with finals, include everything; but exclude the edi and pdf invoices for the printed run
        
        'Fix TTP 10826 / TTP 10813
        'If Invoice!ckcArchive.Value = vbChecked Then
        If Invoice!ckcArchive.Value = vbChecked And (tgSpfx.iInvExpFeature And INVEXP_SELECTIVEEMAIL) <> INVEXP_SELECTIVEEMAIL Then
            'slSelection = slSelection & " and ({IMR_Invoice_Main_Rpt.imrShowInvType} <> 5)"
            slSelection = slSelection & " and ({IVR_Invoice_Rpt.ivrShowInvType} <> 5)"
            'sgSelectionToAdd = " and ({IMR_Invoice_Main_Rpt.imrShowInvType} <> 5)"
            sgSelectionToAdd = " and ({IVR_Invoice_Rpt.ivrShowInvType} <> 5)"
        End If
        
    ElseIf igInvoiceType = 7 Or igInvoiceType = 5 Then        '12-21-16
        '12-21-16 if archiving with finals, include everything; but exclude the edi and pdf invoices for the printed run
        If igInvoiceType = 7 Then
            slSelection = slSelection & " And ({IVR_Invoice_Rpt.ivrType} = 0 or {IVR_Invoice_Rpt.ivrType} = 3)"
        End If

        sgSelection = slSelection
        'Fix TTP 10826 / TTP 10813
        'If Invoice!ckcArchive.Value = vbChecked Then
        If Invoice!ckcArchive.Value = vbChecked And (tgSpfx.iInvExpFeature And INVEXP_SELECTIVEEMAIL) <> INVEXP_SELECTIVEEMAIL Then
            slSelection = slSelection & " and ({IVR_Invoice_Rpt.ivrShowInvType} <> 5) "
            sgSelectionToAdd = ""
        End If
    End If
    

    sgSetSelectionForFinals(UBound(sgSetSelectionForFinals)) = Trim$(slSelection) + Trim$(sgSelectionToAdd)       'excluding type 5 (edi and/or agency pdf)
    ReDim Preserve sgSetSelectionForFinals(0 To UBound(sgSetSelectionForFinals) + 1) As String
    
    sgSetSelectionForAll(UBound(sgSetSelectionForAll)) = Trim$(sgSelection)         'basic gen date and time filter
    ReDim Preserve sgSetSelectionForAll(0 To UBound(sgSetSelectionForAll) + 1) As String
    
    'sgSelection = Trim$(slSelection)
    'Fix TTP 10826 / TTP 10813 - Exclude Invoice when Viewing
    'If Not gSetSelection(slSelection) Then
    If Not gSetSelection(slSelection + slExcludeInvoiceSelection) Then
        gCmcGenIn = -1
        Exit Function
    End If

    '2-1-01 Get the Disclaimer for invoice form from Spf and send via a formula
    '1-22-07 rep has a new type
'    If igInvoiceType = 1 Or igInvoiceType = 5 Then       'type 2 are the affidavits, 3 = Portrait (inv & aff, but disclaimer
'    '12-19-12 Remove using a formula for the SiteDisclaimer text.  Use the pointer that is stored in IVR.  Formula causes error when there is a c/r
'                                    'has special testing and changes based on Advertisers)
'        If tgSpf.lBCxfDisclaimer > 0 Then
'            'open CXF
'            hlCxf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'            ilRet = btrOpen(hlCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'
'            ilCxfRecLen = Len(tlCxf)
'            tlSrchKey.lCode = tgSpf.lBCxfDisclaimer
'            ilRet = btrGetEqual(hlCxf, tlCxf, ilCxfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'            If ilRet = BTRV_ERR_NONE Then
'                'send to crystal
'                'slStr = Mid$(tlCxf.sComment, 1, tlCxf.iStrLen)
'                slStr = gStripChr0(tlCxf.sComment)
'                If Not gSetFormula("Disclaimer", "'" & slStr & "'") Then
'                    gCmcGenIn = -1
'                    Exit Function
'                End If
'
'            End If
'            ilRet = btrClose(hlCxf)
'            btrDestroy hlCxf
'        End If
'        If tgSpf.lBCxfDisclaimer = 0 Or ilRet <> BTRV_ERR_NONE Then
'            'send to crystal (blanks)
'            If Not gSetFormula("Disclaimer", "  ") Then
'                gCmcGenIn = -1
'                Exit Function
'            End If
'        End If
'    End If

    '3-21-03 Send blanks to show in header to align to fit in windowed envelope
    If tgSpf.sExport = "0" Or tgSpf.sExport = "N" Or tgSpf.sExport = "Y" Or tgSpf.sExport = "" Then
        ilBlanksBeforeLogo = 0
    Else
        ilBlanksBeforeLogo = Val(tgSpf.sExport)
    End If
    If Trim$(tgSpf.sImport) = "0" Or Trim$(tgSpf.sImport) = "N" Or Trim$(tgSpf.sImport) = "Y" Or Trim$(tgSpf.sImport) = "" Then
        ilBlanksAfterLogo = 0
    Else
        ilBlanksAfterLogo = Val(tgSpf.sImport)
    End If
    If Not gSetFormula("BlanksBeforeLogo", ilBlanksBeforeLogo) Then
        gCmcGenIn = -1
        Exit Function
    End If
    If Not gSetFormula("BlanksAfterLogo", ilBlanksAfterLogo) Then
        gCmcGenIn = -1
        Exit Function
    End If

    'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
    If igInvoiceType = 4 Or igInvoiceType = 6 Then 'NTR only or Combined
        If Invoice.rbcType(INVGEN_Reprint).Value And Invoice.ckcType(INVTYPE_NTR) = vbChecked Then
            If Invoice.ckcSuppressNTRDetails.Value = vbChecked Then
                If Not gSetFormula("SuppressNTRDetail", "true", False) Then
                    'Not going to fail if the formula is not present to set
                End If
            Else
                If Not gSetFormula("SuppressNTRDetail", "false", False) Then
                    'Not going to fail if the formula is not present to set
                End If
            End If
        End If
    End If
    
    '4-26-12  test for vehicle name word wrap (for forms only)
    If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
        If Not gSetFormula("WordWrapVehicle", "'Y'") Then
            gCmcGenIn = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("WordWrapVehicle", "'N'") Then
            gCmcGenIn = -1
            Exit Function
        End If
    End If


    gCmcGenIn = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportIn                      *
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
Function gGenReportIn() As Integer
Dim blCustomizeLogo As Boolean

    If Not igUsingCrystal Then
        gGenReportIn = True
        Exit Function
    End If

    If igInvoiceType = 2 Then      '2 = affidavit of performance, no special logos; its also a regular report from list
        blCustomizeLogo = False
    Else
        blCustomizeLogo = True          'all invoices will execute customized logos
    End If
    If igInvoiceType = 0 Then               'ordered,aired,reconciled form (previous bridge)
        If Not gOpenPrtJob("Invoice.Rpt", , blCustomizeLogo) Then
            gGenReportIn = False
            Exit Function
        End If
    ElseIf igInvoiceType = 1 Then                   'As ordered invoice (no spots)
        '2-1-02 treat as aired & as ordered/update aired the same
        'If tgSpf.sInvAirOrder = "O" And tgSpf.sBLaserForm = "2" Then
            'this form shows no aired spots
            If Not gOpenPrtJob("InvPort1.Rpt", , blCustomizeLogo) Then
                gGenReportIn = False
                Exit Function
            End If
        'Else
            'This form shows aired spots
        '    If Not gOpenPrtJob("Inv1Air.Rpt") Then
        '        gGenReportIn = False
        '        Exit Function
        '    End If
        'End If
    ElseIf igInvoiceType = 2 Then               'affidavit:  2 passes. 1st = list of spots,
                                                '2nd is summary of spots by invoice
        If igJobRptNo = 1 Then                      'detail pass
            If Not gOpenPrtJob("InvAff.Rpt", , blCustomizeLogo) Then
                gGenReportIn = False
                Exit Function
            End If
        Else                                     'summary pass (spot length counts by vehicle)
            'if Show Ordered,UpdateAired & using separate invoice and affidavit forms, print affidavit
            'summary with market subtotals
            If tgSpf.sInvAirOrder = "O" And tgSpf.sBLaserForm = "2" Then
                If Not gOpenPrtJob("InvAfsm2.Rpt", , blCustomizeLogo) Then     'show market subtotals
                    gGenReportIn = False
                    Exit Function
                End If
            Else
                If Not gOpenPrtJob("InvAffsm.Rpt", , blCustomizeLogo) Then
                    gGenReportIn = False
                    Exit Function
                End If
            End If
        End If
    ElseIf igInvoiceType = 3 Then               'portrait form, as aired billing: combined inv/affidavit
        If Not gOpenPrtJob("InvPort3.Rpt", , blCustomizeLogo) Then 'currently not called (Show invoice as ordered with aired spots)
            gGenReportIn = False
            Exit Function
        End If
    ElseIf igInvoiceType = 4 Then               'NTR
        If Not gOpenPrtJob("Inv_NTR.Rpt", , blCustomizeLogo) Then 'currently not called (Show invoice as ordered with aired spots)
            gGenReportIn = False
            Exit Function
        End If
    ElseIf igInvoiceType = 5 Then               '11-2-06 REP, show ordered and aired spots & $
        If Not gOpenPrtJob("InvPort1A.rpt", , blCustomizeLogo) Then
            gGenReportIn = False
            Exit Function
        End If
    ElseIf igInvoiceType = 6 Then           '1-18-07 combined Air Time and NTR
        If Not gOpenPrtJob("InvCombine.rpt", , blCustomizeLogo) Then
            gGenReportIn = False
            Exit Function
        End If
    ElseIf igInvoiceType = 7 Then           '3-8-12 3-col Aired
        If Not gOpenPrtJob("Inv3ColAired.rpt", , blCustomizeLogo) Then
            gGenReportIn = False
            Exit Function
        End If
    Else                                    'unused
        If igJobRptNo = 1 Then                      'detail pass
            If Not gOpenPrtJob("InvPort2.Rpt", , blCustomizeLogo) Then
                gGenReportIn = False
                Exit Function
            End If
        End If
    End If
    gGenReportIn = True
End Function

Public Function mSameInvFormula() As Integer
        mSameInvFormula = 0
        '10-31-19 show EDI agy client and adv product code.  show them on summary versions only
        If ((Asc(tgSaf(0).sFeatures6) And EDIAGYCODES) = EDIAGYCODES) Then    'using agy client code?
            If Not gSetFormula("ShowEDICodes", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                mSameInvFormula = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowEDICodes", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                mSameInvFormula = -1
                Exit Function
            End If
        End If
    Exit Function
End Function


