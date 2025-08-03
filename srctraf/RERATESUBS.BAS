Attribute VB_Name = "RERATESUBS"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ReRateSubs.BAS
'
' Release: 1.0
'

Public Type BOOKBYLINECNTR
    lChfCode As Long
    lCntrNo As Long
    sSelected As String * 1
    iFirst As Integer
End Type
Public tgBookByLineCntr() As BOOKBYLINECNTR

Public Type BOOKBYLINEASSIGNED
    lClfCode As Long
    iLineNo As Integer
    sType As String * 1
    iPkLineNo As Integer
    iLen As Integer
    iPurchaseDnfCode As Integer '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors (Match DNFcode integer type)
    iReRateDnfCode As Integer   '1/25/22 - JW - TTP 10385 - ""
    iVefCode As Integer
    iRdfCode As Integer
    lChfCode As Long
    lStartDate As Long
    lEndDate As Long
    iMGCount As Integer
    iOutsideCount As Integer
    iBonusCount As Integer
    lColor As Long
    sDaypartName As String * 20
    sAssignMethod As String * 1 'L=by LIne; B=by Book
    iNext As Integer
End Type
Public tgBookByLineAssigned() As BOOKBYLINEASSIGNED
    
Public Type BOOKINFO
    iDnfCode As Integer
    sName As String * 30
    lBookDate As Long
    lFirst As Long
End Type
Public tgBookInfo() As BOOKINFO

Public Type BOOKVEHICLE
    iVefCode As Integer
    lNext As Long
End Type
Public tgBookVehicle() As BOOKVEHICLE
    
Public igReRateAllowedLengths() As Integer
Public lgReRateStartDate As Long
Public lgReRateEndDate As Long

Public igExcludeDnfCode() As Integer
'Used when assigning by Closest Book

Public Type ReRateHeaderInfo
    Agency As String
    Advertiser As String
    Product As String
    OrderNo As String
    PurchaseRevision As String
    ReRateRevision As String
    Demo As String
    ContractPopulation As String
    ReRatePopulation As String
    Period As String
End Type

Public Function gBinarySearchExcludeBook(ilDnfCode As Integer) As Integer

    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer

    ilMin = LBound(igExcludeDnfCode)
    ilMax = UBound(igExcludeDnfCode) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilDnfCode = igExcludeDnfCode(ilMiddle) Then
            'found the match
            gBinarySearchExcludeBook = ilMiddle
            Exit Function
        ElseIf ilDnfCode < igExcludeDnfCode(ilMiddle) Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchExcludeBook = -1
    
End Function

