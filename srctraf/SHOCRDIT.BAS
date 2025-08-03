Attribute VB_Name = "ShoCrditSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SHOCRDIT.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpBkCpy.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ShoCrdit subs and functions
Option Explicit
Option Compare Text


Public igCollCallFrom As Integer '0=Basic 10; 1=ShowCredit
Public sgCollCallName As String  'If igCollCallFrom 1, advertiser or agency name

Type ADVTCREDIT
    iCode As Integer    'Internal code number for advertiser
    sName As String * 30  'Name
    sAddrID As String * 9
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    sCurrAR As String * 6   'Current A/R (xx,xxx,xxx.xx)
    sUnbilled As String * 6 'Unbilled amount (xx,xxx,xxx.xx)
    sCrdApp As String * 1   'Direct Only: Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)
    sBillAgyDir As String * 1   'A=Bill Agency; D=Bill Advertiser
End Type
Public Const ADVTCREDITEXTPK As String = "IB30B9BLB6B6BB"
Type AGYCREDIT
    iCode As Integer    'Internal code number for agency
    sName As String * 40  'Name
    sCityID As String * 5   'City ID
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    sCurrAR As String * 6   'Current A/R (xx,xxx,xxx.xx)
    sUnbilled As String * 6 'Unbilled amount (xx,xxx,xxx.xx)
    sCrdApp As String * 1   'Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)
End Type
Public Const AGYCREDITEXTPK As String = "IB40B5BLB6B6B"

Type SHOCREDITSORT
    sKey As String * 100
    iType As Integer     '0=Advertiser; 1=Agency
    sName As String * 40  'Name
    sAddrID As String * 9
    sCityID As String * 5   'City ID
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    lDollars As Long
    sCrdApp As String * 1   'Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)
    sBillAgyDir As String * 1   'A=Bill Agency; D=Bill Advertiser
    iCode As Integer
End Type
