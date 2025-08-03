Attribute VB_Name = "ProgammaticBuySystemSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of LiveLog.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AgencySystem.Bas
'
' Release: 1.0
'
' Description:
Option Explicit
Option Compare Text


'Global Comm Server Vars
Global sgCommServerIP As String
Global sgCommServerPort As String
Global sgCustomerName As String
Global sgSalesBrochure As String
Global sgProgrammaticCHF As CHF
Global bgShutDown As Boolean
Global lgIdleTime As Long

Global igAgfCode As Integer
Global sgAdfname As String
Global sgAdfAbbrv As String
Global lgAdfCode As Long
Global sgProdName As String
Global lgPrfCode As Long
Global igSlfCode As Integer

Public Type TOKENINFO
    sDate As String * 10
    sTime As String * 10
    iAgfCode As Integer
    sAgency As String * 40
    sBuyer As String * 20
    sEMail As String * 80
    sPhone As String * 20
    sInUseStatus As String * 1
    iRcfCode As Integer
    sCommToken As String * 40
    iTokenIdx As Integer
End Type
Public tgTokenInfo() As TOKENINFO

Global sgLLName As String
Global sgStudioID As String
Global sgSignOnMsg As String
Global sgReplaceMsg As String
Global igAddAvtCode As Integer
Global sgAddTime As String
Global lgAddCopyCode As Long
Global igHelpScreen As Integer '1=Vehicle selection screen; 2= Log screen
Public igEventAdfCode() As Integer 'Event Advertisers

Public sgClientName As String

Public Type ADFINFO
    sName As String * 30
    sProd As String * 35
    sDemo As String * 6
    iAdfAgfCode As Integer
    iCode As Integer
End Type
Public tgAdfInfo() As ADFINFO

