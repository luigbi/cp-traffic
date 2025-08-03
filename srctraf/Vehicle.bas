Attribute VB_Name = "VEHICLESUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Vehicle.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Vehicle.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Global variables for vehicle
Option Explicit
'Public igShowHelpMsg As Integer
Public igVefCodeModel As Integer '0=VefCode to model from
Public igVehReturn As Integer   '0=Cancel; 1=Model
Public igVehMode As Integer     '1=Change; 0=New
                                    'CALLNONE=secondary
                                    'CALLSOURCE----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public igVpfType As Integer '0=All except Package; 1= Package; 2 = Sports only

Public igVehSelType As Integer  '1=Set Default Book Names
Public igVehSelCode As Long     'igVehSelType = 1: Code = dnfCode


Public Sub gInitVef(tlVef As VEF)
    'Create vehicle record

    tlVef.iCode = 0
    tlVef.sName = ""
    tlVef.sAddr(0) = ""
    tlVef.sAddr(1) = ""
    tlVef.sAddr(2) = ""
    tlVef.sPhone = ""
    tlVef.sFax = ""
    tlVef.sUnused1 = ""
    tlVef.sDialPos = ""
    tlVef.lPvfCode = 0
    tlVef.iReallDnfCode = 0
    tlVef.sUpdateRVF(0) = ""
    tlVef.sUpdateRVF(1) = ""
    tlVef.sUpdateRVF(2) = ""
    tlVef.sUpdateRVF(3) = ""
    tlVef.sUpdateRVF(4) = ""
    tlVef.sUpdateRVF(5) = ""
    tlVef.sUpdateRVF(6) = ""
    tlVef.sUpdateRVF(7) = ""
    tlVef.iCombineVefCode = 0
    tlVef.iMnfHubCode = 0
    tlVef.iTrfCode = 0
    tlVef.sType = ""
    tlVef.sCodeStn = ""
    tlVef.iVefCode = 0
    tlVef.iOwnerMnfCode = 0
    tlVef.iProdPct(0) = 0
    tlVef.iProdPct(1) = 0
    tlVef.iProdPct(2) = 0
    tlVef.iProdPct(3) = 0
    tlVef.iProdPct(4) = 0
    tlVef.iProdPct(5) = 0
    tlVef.iProdPct(6) = 0
    tlVef.iProdPct(7) = 0
    tlVef.sState = "A"
    tlVef.iMnfGroup(0) = 0
    tlVef.iMnfGroup(1) = 0
    tlVef.iMnfGroup(2) = 0
    tlVef.iMnfGroup(3) = 0
    tlVef.iMnfGroup(4) = 0
    tlVef.iMnfGroup(5) = 0
    tlVef.iMnfGroup(6) = 0
    tlVef.iMnfGroup(7) = 0
    tlVef.iSort = 0
    tlVef.iDnfCode = 0
    tlVef.iMnfDemo = 0
    tlVef.iMnfSSCode(0) = 0
    tlVef.iMnfSSCode(1) = 0
    tlVef.iMnfSSCode(2) = 0
    tlVef.iMnfSSCode(3) = 0
    tlVef.iMnfSSCode(4) = 0
    tlVef.iMnfSSCode(5) = 0
    tlVef.iMnfSSCode(6) = 0
    tlVef.iMnfSSCode(7) = 0
    tlVef.sExportRAB = "N"
    tlVef.lVsfCode = 0
    tlVef.lRateAud = 0
    tlVef.lCPPCPM = 0
    tlVef.lYearAvails = 0
    tlVef.iPctSellout = 0
    tlVef.iMnfVehGp2 = 0
    tlVef.iMnfVehGp3Mkt = 0
    tlVef.iMnfVehGp4Fmt = 0
    tlVef.iMnfVehGp5Rsch = 0
    tlVef.iMnfVehGp6Sub = 0
    tlVef.iNrfCode = 0
    tlVef.iSSMnfCode = 0
    tlVef.sStdPrice = ""
    tlVef.sStdInvTime = ""
    tlVef.sStdAlter = ""
    tlVef.iStdIndex = 0
    tlVef.sStdAlterName = ""
    tlVef.iRemoteID = 0
    tlVef.iAutoCode = 0
    tlVef.sExtUpdateRvf(0) = ""
    tlVef.sExtUpdateRvf(1) = ""
    tlVef.sExtUpdateRvf(2) = ""
    tlVef.sExtUpdateRvf(3) = ""
    tlVef.sExtUpdateRvf(4) = ""
    tlVef.sExtUpdateRvf(5) = ""
    tlVef.sExtUpdateRvf(6) = ""
    tlVef.sExtUpdateRvf(7) = ""
    tlVef.sStdSelCriteria = ""
    tlVef.sStdOverrideFlag = ""
    tlVef.sContact = ""
    
End Sub
