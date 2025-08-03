Attribute VB_Name = "RPTVFYRA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyra.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRA.Bas
'
' Release: 4.5
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
'Public tgBookName() As SORTCODE
'Public sgBookNameTag As String
'Public tgNamedAvail() As SORTCODE
'Public sgNamedAvailTag As String
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public imRcfCode As Integer     '1-24-01
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
'                                                      *
'       8/18/99 Eliminate the prepass.  Print directly *
'               from ASF                               *
'*                                                     *
'*******************************************************
Function gCmcGenRA(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim slTimeStamp As String
    Dim slTime As String
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilMajorSet As Integer
    Dim slSelection As String
    Dim slCode As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slOr As String

    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String

    If RptSelRA!rbcSel(0).Value Then    'by vehicles (vs package)
        ilLoop = RptSelRA!cbcSet1.ListIndex 'get the type of sorting for output
        ilMajorSet = gFindVehGroupInx(ilLoop, tgVehicleSets1())

        gCmcGenRA = 0

        If Not gSetFormula("Group", str$(ilMajorSet)) Then
            gCmcGenRA = -1
            Exit Function
        End If

        slOr = ""
        If Not RptSelRA!ckcAll.Value = vbChecked Then
            For ilLoop = 0 To RptSelRA!lbcSelection(0).ListCount - 1 Step 1
                If RptSelRA!lbcSelection(0).Selected(ilLoop) Then
                    slNameCode = tgCSVNameCode(ilLoop).sKey 'RptSelRA!tgCSVNameCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    slSelection = slSelection & slOr & "{VEF_Vehicles.vefCode} = " & Trim$(slCode)
                    slOr = " Or "
                End If
            Next ilLoop
        End If
    Else                                'package vehicles summary
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear

        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenRA = -1
            Exit Function
        End If

        If RptSelRA!rbcTotals(0).Value Then          'DP
            If Not gSetFormula("ByDPorVeh", "'D'") Then
                gCmcGenRA = -1
                Exit Function
            End If
        Else                                'vehicle subtotals
            If Not gSetFormula("ByDPorVeh", "'V'") Then
                gCmcGenRA = -1
                Exit Function
            End If
        End If
        slTimeStamp = gFileDateTime(sgDBPath & "asf.btr")
        'Send the date & time stamp of last updated
        '1-20-09 use the date stamp stored in ASF because of REP/NET feature;
        'net may be receiving from different reps so file is not deleted and recreated
        If Not gSetFormula("DateStamp", "'" & slTimeStamp & "'") Then
            gCmcGenRA = -1
            Exit Function
        End If

    End If
    slTimeStamp = gFileDateTime(sgDBPath & "asf.btr")
    'Send the date & time stamp of last updated
    '1-20-09 use the date stamp stored in ASF because of REP/NET feature;
    'net may be receiving from different reps so file is not deleted and recreated
'    If Not gSetFormula("DateStamp", "'" & slTimeStamp & "'") Then
'        gCmcGenRA = -1
'        Exit Function
'    End If
    If Not gSetSelection(slSelection) Then
        gCmcGenRA = -1
        Exit Function
    End If
    gCmcGenRA = 1         'ok
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
    RptSelRA!frcOutput.Enabled = igOutput
    RptSelRA!frcCopies.Enabled = igCopies
    'RptSelRA!frcWhen.Enabled = igWhen
    RptSelRA!frcFile.Enabled = igFile
    RptSelRA!frcOption.Enabled = igOption
    'RptSelRA!frcRptType.Enabled = igReportType
    Beep
End Sub
