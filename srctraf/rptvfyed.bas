Attribute VB_Name = "rptvfyed"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelED.Bas
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
'Public tgRptSelAgencyCode() As SORTCODE
'Public sgRptSelAgencyCodeTag As String
'Public tgRptSelSalespersonCode() As SORTCODE
'Public sgRptSelSalespersonCodeTag As String
'Public tgRptSelAdvertiserCode() As SORTCODE
'Public sgRptSelAdvertiserCodeTag As String
'Public tgRptSelNameCode() As SORTCODE
'Public sgRptSelNameCodeTag As String
'Public tgRptSelBudgetCode() As SORTCODE
'Public sgRptSelBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelDemoCode() As SORTCODE
'Public sgRptSelDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
''Global tgLnSdfExtSort() As SDFEXTSORT
''Global tgLnSdfExt() As SDFEXT
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
Public lgTotal_ISRRecs As Long      '3-28-02    record count created in ISR

Type RVF_VEF
    iVefCode As Integer
    iPkLineNo As Integer            '0 if conventional
    lTotalOrd(1 To 14) As Long      'total $ ordered from package lines
    lTotalGross(1 To 14) As Long    'Total dollars within billing period (cash only), 0=prior to year, 1-12 = requested 12 months, 13 = future
    lTotalVefDollars As Long        'Total Vehicle Dollars (sdf)
    lTotalVefBilledDollars As Long  'Total Vehicle Billed Dollars (phf/rvf)
End Type
Type RVFSORT
    sKey As String * 10
    tlRvfRec As RVF
End Type
Type STARTENDINX                'table containing either RVF/PHF or Spots that are
                                'sorted by contract # or line # which has the starting
                                'and ending index of the records to process so the entire
                                'array does not have to be searched for every contract
    iProcessed As Integer       '0 = contract not proc. flag, 1 = contract processed
                                'remains a zero if the contract has expired but RVF/PHF exist
    lLineNo As Long
    iStartInx As Integer
    iEndInx As Integer
End Type
Type SDFSORTLIST            'table of SDF entries sorted by   Line  # (5 characters)
    sKey As String * 5      'XXXXX
    tSdf As SDF
End Type
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmcGen                         *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*
'*          Return : 0 =  either error in input, stay in
'*                   -1 = error in Crystal, return to
'*                        calling program
''*                       failure of gSetformula or another
'*                    1 = Crystal successfully completed
'*                    2 = successful Bridge
'*******************************************************
Function gCmcGenEd(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
    Dim slSelection As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilYear As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilValue As Integer
        gCmcGenEd = 0
        slStr = RptSelED!edcSelCFrom.Text
        ilYear = mVerifyYear(slStr)
        If ilYear = 0 Then
            mReset
            RptSelED!edcSelCFrom.SetFocus                 'invalid year
            gCmcGenEd = False
            Exit Function
        End If
        slStr = RptSelED!edcQtr.Text
        ilRet = mVerifyInt(slStr, 1, 4)
        If ilRet = -1 Then
            mReset
            RptSelED!edcQtr.SetFocus                 'invalid qtr
            gCmcGenEd = False
            Exit Function
        End If
        'following formulas sent to Crystal to monthly headings
        slStr = ""
        ilValue = Val(RptSelED!edcQtr.Text)
        If ilValue = 1 Then
            slStr = "1st Quarter "
        ElseIf ilValue = 2 Then
            slStr = "2nd Quarter "
        ElseIf ilValue = 3 Then
            slStr = "3rd Quarter "
        Else
            slStr = "4th Quarter "
        End If
        slStr = slStr & RptSelED!edcSelCFrom.Text
        If Not gSetFormula("QtrHeader", "'" & slStr & "'") Then
            gCmcGenEd = -1
            Exit Function
        End If
        If Not gSetFormula("StartingMonth", ((ilValue - 1) * 3 + 1)) Then         'pass starting month of the starting std qtr for report column headings
            gCmcGenEd = -1
            Exit Function
        End If

        '3-27-02 option for totals by contract or advt
        If RptSelED!rbcTotalsBy(0).Value Then           'totals by contract
            If Not gSetFormula("TotalsBy", "'" & "C" & "'") Then
                gCmcGenEd = -1
                Exit Function
            End If
        Else                                            'totals by advt
            If Not gSetFormula("TotalsBy", "'" & "A" & "'") Then
                gCmcGenEd = -1
                Exit Function
            End If
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenEd = -1
            Exit Function
        End If

    gCmcGenEd = 1
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
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
Function gGenReportED() As Integer
    If RptSelED!rbcEarnYr(0).Value Then       'show only the current year version
        If RptSelED!rbcSelC4(2).Value Or RptSelED!rbcSelC4(3) Then  '4-4-02 Consolidated gross or net without participant splits
            If Not gOpenPrtJob("PEYrCon.Rpt") Then
                gGenReportED = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("ProdEarn.Rpt") Then
                gGenReportED = False
                Exit Function
            End If
        End If
    Else
        If RptSelED!rbcSelC4(2).Value Or RptSelED!rbcSelC4(3) Then  '4-4-02 Consolidated gross or net
            If Not gOpenPrtJob("PEBalCon.Rpt") Then  'show the entire contract totals version without particpant splits
                gGenReportED = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("ProdBal.Rpt") Then  'show the entire contract totals version
                gGenReportED = False
                Exit Function
            End If
        End If
    End If
    gGenReportED = True
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
    RptSelED!frcOutput.Enabled = igOutput
    RptSelED!frcCopies.Enabled = igCopies
    'RptSelED!frcWhen.Enabled = igWhen
    RptSelED!frcFile.Enabled = igFile
    RptSelED!frcOption.Enabled = igOption
    'RptSelED!frcRptType.Enabled = igReportType
    Beep
End Sub

'
'               mVerifyInt - verify input.  Value must be between two arguments provided
'               <input>     slStr - user input
'                           ilLowInt - lowest value allowed
'                           ilHiInt - highest value allowed
'               <output>    Return - converted integer
'                                    -1 if invalid
Function mVerifyInt(slStr As String, ilLowInt As Integer, ilHiInt As Integer) As Integer
Dim ilInput As Integer
    mVerifyInt = 0
    ilInput = Val(slStr)
    If (ilInput < ilLowInt) Or (ilInput > ilHiInt) Then
        mVerifyInt = -1
    Else
        mVerifyInt = ilInput
    End If
End Function
'
'               mVerifyYear - Verify that the year entered is valid
'                             If not 4 digit year, add 1900 or 2000
'                             Valid year must be > than 1950 and < 2050
'                             Input - string containing input
'                             Output - Integer containing year else 0
'
Function mVerifyYear(slStr As String) As Integer
Dim ilInput As Integer
    mVerifyYear = 0
    If IsNumeric(slStr) Then
        ilInput = Val(slStr)
        If ilInput < 100 Then           'only 2 digit year input ie.  96, 95,
            If ilInput < 50 Then        'adjust for year 1900 or 2000
                ilInput = 2000 + ilInput
            Else
                ilInput = 1900 + ilInput
            End If
        End If
        If (ilInput < 1950) Or (ilInput > 2050) Then
            mVerifyYear = 0
        Else
            mVerifyYear = ilInput
        End If
    End If
End Function
