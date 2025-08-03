Attribute VB_Name = "RptvfyRR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyrr.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptvfyRR.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
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
'Public tgRptSelRRvertiserCode() As SORTCODE
'Public sgRptSelRRvertiserCodeTag As String
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
'Public tgBookName() As SORTCODE
'Public sgBookNameTag As String
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
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
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenRR(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    Dim slSelection As String
    ReDim ilDate(0 To 1) As Integer
    Dim llDate As Long
    Dim llEDate As Long
    Dim ilDays As Integer
    Dim slInclTypeStatus As String   'Date: 4/10/2020 exclude contract types
    Dim slExclTypeStatus As String   'Date: 4/10/2020 exclude contract types
    
    gCmcGenRR = 0
'    slDate = RptSelRR!edcSelCFrom.Text
    slDate = RptSelRR!CSI_CalFrom.Text          '12-16-19 change to use csi calendar control
    If Not gValidDate(slDate) Then
        mReset
        RptSelRR!CSI_CalFrom.SetFocus
        Exit Function
    End If
    slDate = RptSelRR!edcSelCTo.Text        '# Days
    ilRet = gVerifyInt(slDate, 1, 35)
    If ilRet < 1 Then
        mReset
        RptSelRR!edcSelCTo.SetFocus
        Exit Function
    End If
    ilDays = Val(slDate)
    If ilDays > 35 Then     'can only run this report for a month (calendar or 5 week std bdcst)
        MsgBox "Maximum 35 days allowed, please reduce date span ", vbOKOnly
        mReset
        RptSelRR!edcSelCTo.SetFocus
        Exit Function
    End If

    'Active Dates entered
 '   slDate = RptSelRR!edcSelCFrom.Text
    slDate = RptSelRR!CSI_CalFrom.Text          '12-16-19 change to use csi calendar control
    gPackDate slDate, ilDate(0), ilDate(1)
    gUnpackDateLong ilDate(0), ilDate(1), llDate
    If slDate = "" Then
        slStr = ""
    Else
        slStr = Format$(llDate, "m/d/yy")
    End If
    llEDate = llDate + ilDays - 1        'calc end date


    If slStr = "" And slDate = "" Then                          'selective contracts, no dates enterd
        slStr = "All dates for selective contracts"
    Else
        slStr = slStr & " - " & Format$(llEDate, "m/d/yy")
    End If

    If Not gSetFormula("ActiveDates", "'" & slStr & "'") Then
        gCmcGenRR = -1
        Exit Function
    End If
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenRR = -1
        Exit Function
    End If

    If RptSelRR!rbcBook(0).Value Then       'closest book
        If Not gSetFormula("Book", "'Use Closest book to airing'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    ElseIf RptSelRR!rbcBook(1).Value Then
        If Not gSetFormula("Book", "'Use vehicle default book'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("Book", "'Use schedule line book'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    End If

    If RptSelRR!rbcSortBy(0).Value Then     'sort by ADvt?
        If Not gSetFormula("Sortby", "'A'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    ElseIf RptSelRR!rbcSortBy(1).Value Then 'sort by Vehicle
        If Not gSetFormula("Sortby", "'V'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    ElseIf RptSelRR!rbcSortBy(2).Value Then
        If Not gSetFormula("Sortby", "'P'") Then    'prod category
            gCmcGenRR = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("Sortby", "'B'") Then    'business category
            gCmcGenRR = -1
            Exit Function
        End If
    End If
    
    If RptSelRR!rbcGrossNet(0).Value Then       '2-28-20
        If Not gSetFormula("GrossNet", "'G'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("GrossNet", "'N'") Then
            gCmcGenRR = -1
            Exit Function
        End If
    End If
    
    'Date: 4/10/2020 display include/exclude contract types (CVTRQSM) and status (HOGN)
    slInclTypeStatus = "": slExclTypeStatus = ""
    If RptSelRR!ckcSelC5(7).Value = vbChecked Then
        slInclTypeStatus = "Holds"
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Holds"
        End If
    End If
    If RptSelRR!ckcSelC5(8).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "Orders"
        Else
            slInclTypeStatus = slInclTypeStatus & ", Orders"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Orders"
        Else
            slExclTypeStatus = slExclTypeStatus & ", Orders"
        End If
    End If
    If RptSelRR!ckcSelC5(0).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "Std"
        Else
            slInclTypeStatus = slInclTypeStatus & ", Std"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Std"
        Else
            slExclTypeStatus = slExclTypeStatus & ", Std"
        End If
    End If
    If RptSelRR!ckcSelC5(1).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "Resv"
        Else
            slInclTypeStatus = slInclTypeStatus & ", Resv"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Resv"
        Else
            slExclTypeStatus = slExclTypeStatus & ", Resv"
        End If
    End If
    If RptSelRR!ckcSelC5(2).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "Rem"
        Else
            slInclTypeStatus = slInclTypeStatus & ", Rem"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Rem"
        Else
            slExclTypeStatus = slExclTypeStatus & ", Rem"
        End If
    End If
    If RptSelRR!ckcSelC5(3).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "DR"
        Else
            slInclTypeStatus = slInclTypeStatus & ", DR"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "DR"
        Else
            slExclTypeStatus = slExclTypeStatus & ", DR"
        End If
    End If
    If RptSelRR!ckcSelC5(4).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "PI"
        Else
            slInclTypeStatus = slInclTypeStatus & ", PI"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "PI"
        Else
            slExclTypeStatus = slExclTypeStatus & ", PI"
        End If
    End If
    If RptSelRR!ckcSelC5(5).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "PSA"
        Else
            slInclTypeStatus = slInclTypeStatus & ", PSA"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "PSA"
        Else
            slExclTypeStatus = slExclTypeStatus & ", PSA"
        End If
    End If
    If RptSelRR!ckcSelC5(6).Value = vbChecked Then
        If slInclTypeStatus = "" Then
            slInclTypeStatus = "Promo"
        Else
            slInclTypeStatus = slInclTypeStatus & ", Promo"
        End If
    Else
        If slExclTypeStatus = "" Then
            slExclTypeStatus = "Promo"
        Else
            slExclTypeStatus = slExclTypeStatus & ", Promo"
        End If
    End If
    
    If Not gSetFormula("IncludeTypes", "'" & slInclTypeStatus & "'") Then
        gCmcGenRR = -1
        Exit Function
    End If
    If Not gSetFormula("ExcludeTypes", "'" & slExclTypeStatus & "'") Then
        gCmcGenRR = -1
        Exit Function
    End If
    
    gCmcGenRR = 1         'ok
    Exit Function
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
    RptSelRR!frcOutput.Enabled = igOutput
    RptSelRR!frcCopies.Enabled = igCopies
    'RptSelRR!frcWhen.Enabled = igWhen
    RptSelRR!frcFile.Enabled = igFile
    RptSelRR!frcOption.Enabled = igOption
    'RptSelRR!frcRptType.Enabled = igReportType
    Beep
End Sub
