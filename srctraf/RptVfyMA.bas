Attribute VB_Name = "RptVfyMA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptVfyMA.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelMA.Bas
'
' Release: 4.5
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

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
Function gCmcGenMA(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    Dim slNameCode As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slReserve As String
    Dim ilVGSort As Integer
    Dim slSort As String * 1
    Dim slCodeForSort As String * 6
    Dim ilListIndex As Integer
    Dim ilSort1 As Integer
    Dim ilSort2 As Integer
    Dim ilSort3 As Integer
    Dim ilNoWeeks As Integer
    Dim slWeekStart As String
    Dim llWeekStart As Long
    Dim slWeekEnd As String
    Dim llWeekEnd As Long
    
        gCmcGenMA = 0
   
        slExclude = ""
        slInclude = ""
        
        gIncludeExcludeCkc RptSelMA!ckcProposals(0), slInclude, slExclude, "Working"
        gIncludeExcludeCkc RptSelMA!ckcProposals(1), slInclude, slExclude, "Complete"
        gIncludeExcludeCkc RptSelMA!ckcProposals(2), slInclude, slExclude, "Unapproved"

        gIncludeExcludeCkc RptSelMA!ckcCType(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelMA!ckcCType(1), slInclude, slExclude, "Orders"

        gIncludeExcludeCkc RptSelMA!ckcCType(3), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelMA!ckcCType(4), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelMA!ckcCType(5), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelMA!ckcCType(6), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelMA!ckcCType(7), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelMA!ckcCType(8), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelMA!ckcCType(9), slInclude, slExclude, "Promo"

        gIncludeExcludeCkc RptSelMA!ckcCType(10), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelMA!ckcCType(2), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelMA!ckcCType(11), slInclude, slExclude, "Non-Polit"

        If Len(slInclude) > 0 Then
            slInclude = "Include: " & slInclude
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenMA = -1
                Exit Function
            End If
        End If
        If Len(slExclude) > 0 Then
            slExclude = "Exclude: " & slExclude
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenMA = -1
                Exit Function
            End If
        End If

        ilListIndex = RptSelMA!cbcSet1.ListIndex
        ilVGSort = gFindVehGroupInx(ilListIndex, tgVehicleSets1())

        If Not gSetFormula("VGSort", ilVGSort) Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        'N = none; A = advt/prod,cnt; C= Advt/Prod, cnt,veh; S = slsp; V = Vehicle, advt/prod; G = vehicle group
        slCodeForSort = "NSG"
        ilSort1 = RptSelMA!cbcSort1.ListIndex
        ilSort2 = RptSelMA!cbcSort2.ListIndex
        ilSort3 = RptSelMA!cbcSort3.ListIndex
        
        'tell crystal what the sort parameters are
        slSort = Mid(slCodeForSort, ilSort1 + 1)
        If Not gSetFormula("SortField1", "'" & slSort & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        slSort = Mid(slCodeForSort, ilSort2 + 1)
        If Not gSetFormula("SortField2", "'" & slSort & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        'N = none; A = advt/prod,cnt; C= Advt/Prod, cnt,veh; S = slsp; V = Vehicle, advt/prod; G = vehicle group
        '
        slCodeForSort = "ACV"
      
        'tell crystal what the sort parameters are
        'Listindex returned is relative to 0
        slSort = Mid(slCodeForSort, ilSort3 + 1)
        If Not gSetFormula("SortField3", "'" & slSort & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        'each group has an option to skip to a new page
        slSort = "Y"
        If RptSelMA!ckcSkipSort1.Value = vbUnchecked Then
            slSort = "N"
        End If
        If Not gSetFormula("Sort1NewPage", "'" & slSort & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        slSort = "Y"
        If RptSelMA!ckcSkipSort2.Value = vbUnchecked Then
            slSort = "N"
        End If
        If Not gSetFormula("Sort2NewPage", "'" & slSort & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        
        slWeekStart = RptSelMA!calStartDate.Text
        llWeekStart = gDateValue(slWeekStart)
    
        slWeekStart = Format$(llWeekStart, "m/d/yy")
        ilNoWeeks = Val(RptSelMA!edcNoWeeks.Text)
        llWeekEnd = llWeekStart + ((ilNoWeeks - 1) * 7) + 6
        slWeekEnd = Format$(llWeekEnd, "m/d/yy")
        slStr = "Active Dates " & slWeekStart & "-" & slWeekEnd
        If Not gSetFormula("ActiveDates", "'" & slStr & "'") Then
            gCmcGenMA = -1
            Exit Function
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenMA = -1
            Exit Function
        End If
    gCmcGenMA = 1         'ok
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
    RptSelMA!frcOutput.Enabled = igOutput
    RptSelMA!frcCopies.Enabled = igCopies
    'RptSelMA!frcWhen.Enabled = igWhen
    RptSelMA!frcFile.Enabled = igFile
    RptSelMA!frcOption.Enabled = igOption
    'RptSelMA!frcRptType.Enabled = igReportType
    Beep
End Sub
'
'               gGenReport - Validity check and open Crystal report
'
Public Function gGenReportMA() As Integer
Dim slStr As String
Dim ilRet As Integer
Dim ilSort3 As Integer
Dim slCodeForSort As String
Dim slSort As String * 1

        gGenReportMA = True
        slStr = RptSelMA!calStartDate.Text
        If Not gValidDate(slStr) Then
            mReset
            RptSelMA!calStartDate.SetFocus
            gGenReportMA = False
            Exit Function
        End If
    
        slStr = RptSelMA!edcNoWeeks.Text                  'edit # weeks
    
        ilRet = gVerifyInt(slStr, 1, 53)
        If ilRet = -1 Then
            mReset
            RptSelMA!edcNoWeeks.SetFocus                 'invalid # weeks
            gGenReportMA = False
            Exit Function
        End If
        
        ilSort3 = RptSelMA!cbcSort3.ListIndex
        'N = none; A = advt/prod,cnt; C= Advt/Prod, cnt,veh; S = slsp; V = Vehicle, advt/prod; G = vehicle group
        slCodeForSort = "ACV"
        slSort = Mid(slCodeForSort, ilSort3 + 1)
        
        If slSort = "A" Then
            If Not gOpenPrtJob("MarginAcqCnt.rpt") Then
                gGenReportMA = False
                Exit Function
            End If
        ElseIf slSort = "C" Then
            If Not gOpenPrtJob("MarginAcqCntVeh.rpt") Then
                gGenReportMA = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("MarginAcqVeh.rpt") Then
                gGenReportMA = False
                Exit Function
            End If
        End If
        
        
End Function
