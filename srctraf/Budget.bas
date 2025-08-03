Attribute VB_Name = "BUDGETSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budget.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Budget.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
Type BVFREC
    sKey As String * 110 'Sales Office; Vehicle Name
    SOffice As String * 40
    sMktRank As String * 5
    sVehicle As String * 50
    sVehSort As String * 5
    tBvf As BVF
    iStatus As Integer
    lRecPos As Long
    iSaveIndex As Integer
End Type
Type BSFREC
    sKey As String * 40 'Salesperson
    tBsf As BSF
    iStatus As Integer
    lRecPos As Long
    iSaveIndex As Integer
End Type
Type BDUSERVEH
    iCode As Integer
    sName As String * 40
    sState As String * 1
    iSort As Integer
End Type
Type SALEOFFICE
    iCode As Integer
    sName As String * 40
    sMktRank As String * 5
    sState As String * 1
End Type
Type BUDAVAILS
    tRdf As RDF 'Base dayparts only
    iSellout As Integer
    lRifIndex As Long
    i30Avails As Long 'Avail count for 30's
    lRate As Long   'Rate from Rate Card
End Type
Type RCFINFO
    iRcfCode As Integer
    iYear As Integer
    iSellout As Integer
End Type
'In RateCardSubs
'Type PDGROUPS
'    iYear As Integer
'    iStartWkNo As Integer
'    iNoWks As Integer
'    iTrueNoWks As Integer
'    iFltNo As Integer
'    sStartDate As String
'    sEndDate As String
'End Type
Type ADVTTOTALS
    sKey As String * 46
    iAdfCode As Integer
    sTotal As String * 10
    sIndex As String * 6
    lTotal(0 To 4) As Long  'Year Totals
    iPtAdvtTotals As Integer    'Required to find AdvtValues after sort
    iFirstValue As Integer
End Type
Type ADVTVALUES
    iPtAdvtTotals As Integer
    iVefCode As Integer
    iSofCode As Integer
    'lDollars(0 To 4, 1 To 53)  As Long
    lDollars(0 To 4, 0 To 53)  As Long      'Index zero ignored with Dollars
    iNextValue As Integer
End Type
Type BUDRESEARCH
    iVefCode As Integer
    iMnfDemo As Integer
    sDemoName As String * 8
    lRateAud As Long
    lCPPCPM As Long
    lAvails As Long
    iPctSellout As Integer
    lDollars As Long
End Type
Type MNTHINFO
    sName As String * 9
    iStartWkNo As Integer
    iEndWkNo As Integer
End Type
Public sgBDShow() As String * 40 'Values shown in budget area (1=Name; 2-6=Dollars)
Public igLBBDShow As Integer
Public igBDView As Integer
Public tgBvfRec() As BVFREC  'Budget by Office
Public igLBBvfRec As Integer
Public tgBsfRec() As BSFREC  'Budget by Salesperson
Public igLBBsfRec As Integer
Public igMode As Integer     '0=New; 1=Old
Public igBDReturn As Integer    '0=Cancelled
Public lgTotal As Long
Public igBudgetType As Integer  '0=Budget; 1=Actuals
'Public igCurNoWks As Integer
'Public igNewNoWks As Integer
Public igModelMnfBudget As Integer
Public igModelYear As Integer
Public sgBudgetName As String
Public igDirect As Integer
Public igNewMnfBudget As Integer
Public igNewYear As Integer
Public sgBAName As String
Public sgBvfVehName As String
Public sgBvfOffName As String
Public tgSalesOfficeCode() As SORTCODE


Public tgBudUserVehicle() As SORTCODE
Public sgBudUserVehicleTag As String


'*******************************************************
'*                                                     *
'*      Procedure Name:mStraightLinePrediction         *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Straight Line Prediction of    *
'*                      Dollars                        *
'*                                                     *
'*******************************************************
Function mStraightLinePrediction(llDollar() As Long) As Long
    Dim llTDollar As Long
    Dim llTCount As Long
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim flSX As Single
    Dim flB As Single
    Dim flA As Single
    Dim flT As Single
    Dim flST2 As Single
    ilCount = 0
    llTDollar = 0
    For ilLoop = LBound(llDollar) To UBound(llDollar) Step 1
        ilCount = ilCount + 1
        llTCount = llTCount + ilCount
        llTDollar = llTDollar + llDollar(ilLoop)
    Next ilLoop
    flSX = CSng(llTCount) / CSng(ilCount)
    ilCount = 0
    For ilLoop = LBound(llDollar) To UBound(llDollar) Step 1
        ilCount = ilCount + 1
        flT = CSng(ilCount) - flSX
        flST2 = flST2 + flT * flT
        flA = flA + flT * CSng(llDollar(ilLoop))
    Next ilLoop
    flA = flA / flST2
    flB = (CSng(llTDollar) - CSng(llTCount) * flA) / CSng(ilCount)
    'Y = AX + B
    mStraightLinePrediction = CLng(CSng((ilCount + 1)) * flA + flB)
End Function

    

