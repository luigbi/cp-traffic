Attribute VB_Name = "RPTVFYRevEvt"


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyRevEvt.Bas  Revenue by Event
'
Option Explicit
Option Compare Text

'If adding or changing order of sort/selection list boxes, change these constants and also
'see rptcrRevEvt for any further tests.
Const SORT_ADVT = 1
Const SORT_TITLE1 = 2
Const SORT_TITLE2 = 3
Const SORT_SUBT1 = 4
Const SORT_SUBT2 = 5
Const SORT_VEHICLE = 6


'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportRevEvt                *
'*      Spot Business Booked
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenRevEvt() As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
Dim slDateFrom As String
Dim slDateTo As String
Dim slDate As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
Dim slTime As String
Dim slSelection As String
Dim slInclude As String
Dim slExclude As String
Dim llDate As Long
Dim ilSort As Integer
Dim slUseVG As String * 1
Dim slSortCode As String * 1
Dim ilPerStartDate(0 To 1) As Integer
Dim ilSaveMonth As Integer
Dim ilDay As Integer
Dim ilYear As Integer
Dim slMonthInYear As String * 36


        gCmcGenRevEvt = 0
                        
        ilSort = (RptSelRevEvt!cbcSort1.ListIndex) + 1      '0 will indicate none for other 2 sorts
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort1", "'" & slSortCode & "'") Then
            gCmcGenRevEvt = -1
            Exit Function
        End If

        ilSort = RptSelRevEvt!cbcSort2.ListIndex
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort2", "'" & slSortCode & "'") Then
            gCmcGenRevEvt = -1
            Exit Function
        End If
        
        ilSort = RptSelRevEvt!cbcSort3.ListIndex
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort3", "'" & slSortCode & "'") Then
            gCmcGenRevEvt = -1
            Exit Function
        End If
               
        If RptSelRevEvt!ckcSkip1.Value = vbChecked Then
            If Not gSetFormula("SkipSort1", "'Y'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort1", "'N'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        End If

        If RptSelRevEvt!ckcSkip2.Value = vbChecked Then
            If Not gSetFormula("SkipSort2", "'Y'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort2", "'N'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        End If
        
        If RptSelRevEvt!ckcSkip3.Value = vbChecked Then
            If Not gSetFormula("SkipSort3", "'Y'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort3", "'N'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        End If
      
        slDateFrom = RptSelRevEvt!calStart.Text
        If Not gSetFormula("StartDate", "'" & slDateFrom & "'") Then
            gCmcGenRevEvt = -1
            Exit Function
        End If
        slDateTo = RptSelRevEvt!calEnd.Text
        If Not gSetFormula("EndDate", "'" & slDateTo & "'") Then
            gCmcGenRevEvt = -1
            Exit Function
        End If
        
        'airtime/ntr notation
        If RptSelRevEvt!rbcAirTimeNTR(0).Value = True Then
            If Not gSetFormula("AirTimeNTRRequested", "'Incl: Air Time Only'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        ElseIf RptSelRevEvt!rbcAirTimeNTR(1).Value = True Then
            If Not gSetFormula("AirTimeNTRRequested", "'Incl:  NTR Only'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("AirTimeNTRRequested", "'Incl:  Air Time & NTR'") Then
                gCmcGenRevEvt = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
        If Not gSetSelection(slSelection) Then
            gCmcGenRevEvt = 0
            Exit Function
        End If
            
        gCmcGenRevEvt = 1         'ok
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportRevEvt                *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Validity checking of input &   *
'                               Open the Crystal report                *
'
'*******************************************************
Function gGenReportRevEvt() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer

        If RptSelRevEvt!ckcSummaryOnly = vbChecked Then
            If Not gOpenPrtJob("RevByEventSum.Rpt") Then
                gGenReportRevEvt = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("RevByEvent.Rpt") Then
                gGenReportRevEvt = False
                Exit Function
            End If
        End If
       

    gGenReportRevEvt = True
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
    RptSelRevEvt!frcOutput.Enabled = igOutput
    RptSelRevEvt!frcCopies.Enabled = igCopies
    'RptSelRevEvt!frcWhen.Enabled = igWhen
    RptSelRevEvt!frcFile.Enabled = igFile
    RptSelRevEvt!frcOption.Enabled = igOption
    'RptSelRevEvt!frcRptType.Enabled = igReportType
    Beep
End Sub
'
'                   mConvertIndexToCode - convert the index number of sort selection
'                   to a alpha code to send to crystal
'                   <input> index to selection
'                   <return>  1 char code indicating the sort selected
Private Function mConvertIndexToCode(ilIndex As Integer) As String
Dim slChar As String * 1
        
        slChar = "N"            'assume NONE selected
        If ilIndex = SORT_ADVT Then
            slChar = "A"
        ElseIf ilIndex = SORT_TITLE1 Then
            slChar = "1"
        ElseIf ilIndex = SORT_TITLE2 Then
            slChar = "2"
        ElseIf ilIndex = SORT_SUBT1 Then
            slChar = "S"
        ElseIf ilIndex = SORT_SUBT2 Then
            slChar = "U"
        ElseIf ilIndex = SORT_VEHICLE Then
            slChar = "V"
        End If
        mConvertIndexToCode = slChar
        Exit Function
End Function
