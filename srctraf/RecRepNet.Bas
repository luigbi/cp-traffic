Attribute VB_Name = "RecNetRep"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RecRepNet.Bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  tgServerNRF                   sgNetDBPath                   sgRepDBPath               *
'*  sgXFerDBID                                                                            *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  GIFKEY1                       MIF                           MIFKEY1                   *
'*  MIFKEY2                       QIF                           QIFKEY1                   *
'*                                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: NetRep.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for the Rep-Net Link interface
Option Explicit

Public sgRepDBID As String
Public sgNetDBID As String

Public tgNrf As NRF
Public tgServerNRF() As NRF
Public sgNetDBPath As String
Public sgRepDBPath As String
Public sgXFerDBID As String

'******************************************************************************
' GIF_Gather_Info Record Definition
'
'******************************************************************************
Type GIF
    lCode                 As Long            ' Auto-Increment
    sNRProcessed          As String * 1      ' Net-Rep processed this reocrd (Y
                                             ' or N)
    lGenDate              As Long            ' Date that is stored within the
                                             ' transferred text file
    lGenTime              As Long            ' Time that is stored within the
                                             ' transferred text file
    sType                 As String * 1      ' Kind of record to be transmitted
                                             ' (C=Contract/Lines; S=Posted
                                             ' Spots; M=Message; A=6 month
                                             ' avails)
    lCntrNo               As Long            ' Contract number if gifType = C
    iVefCode              As Integer         ' Veficle reference code.  gifType
                                             ' = S and A
    iStartDate(0 To 1)    As Integer         ' Start date.  gifType = S
    iEndDate(0 To 1)      As Integer         ' End Date.  gifType = S
    sMessageRepDBID As String * 5
    sMessageNetDBID As String * 5
    lMessageCefCode       As Long            ' Message reference. gifType = M
    iProcessedDate(0 To 1) As Integer        ' Processed Date
    iProcessedTime(0 To 1) As Integer        ' Processed Time
    sUnused               As String * 20     ' Unused
End Type


'Type GIFKEY0
'    lCode                 As Long
'End Type

Type GIFKEY1 'VBC NR
    sNRProcessed          As String * 1 'VBC NR
    lGenDate              As Long 'VBC NR
    lGenTime              As Long 'VBC NR
End Type 'VBC NR


'******************************************************************************
' MIF_Map_Info Record Definition
'
'******************************************************************************
Type MIF 'VBC NR
    lCode                 As Long            ' Auto-Increment 'VBC NR
    sType                 As String * 1      ' Kind of record: A=Advertiser; 'VBC NR
                                             ' G=Agency; D=Daypart; V=Vehicle
    sRepDBID              As String * 5      ' Rep database ID 'VBC NR
    sNetDBID              As String * 5      ' Network database ID 'VBC NR
    lNetDBCode            As Long            ' Network item Code 'VBC NR
    lRepDBCode            As Long            ' Rep system Item code 'VBC NR
    sUnused               As String * 10     ' Unused 'VBC NR
End Type 'VBC NR


'Type MIFKEY0
'    lCode                 As Long
'End Type

Type MIFKEY1 'VBC NR
    sType                 As String * 1 'VBC NR
    sNetDBID              As String * 5 'VBC NR
    lNetDBCode            As Long 'VBC NR
End Type 'VBC NR

Type MIFKEY2 'VBC NR
    sType                 As String * 1 'VBC NR
    sRepDBID              As String * 5 'VBC NR
    lRepDBCode            As Long 'VBC NR
End Type 'VBC NR


'******************************************************************************
' NRF_Network_Rep_Info Record Definition
'
'******************************************************************************
Type NRF
    iCode                 As Integer         ' Auto-Increment
    sType                 As String * 1      ' Type: N=Network; R=Rep
    sName                 As String * 40     ' Name of the Network or Rep
    sDBID                 As String * 5      ' Database ID of Network or Rep
    sInsertionType        As String * 1      ' Insertion type.  nrfType = N or
                                             ' Rep system. C=CSI Interface;
                                             ' M=Manual; N=Neither
    sSlspFirstName        As String * 20     ' Generic Salesperson First Name.
                                             ' nrfType = R
    sSlspLastName         As String * 20     ' Generic Salesperson Last name.
                                             ' nrfType = R
    sEMail                As String * 80     ' E-Maill address.  nrfType = R or
                                             ' N.  Rep defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sPhoneNo              As String * 25     ' Phone number.  nrfType = R or N.
                                             ' Rep defined on Network system.
                                             ' Network defined on Rep system
    sContactName          As String * 40     ' Contact Name.  nrfType = R or N.
                                             ' Rep defined on Network system.
                                             ' Network defined on Rep system
    sFTPIsOn              As String * 1      ' FTPIsOn.  nrfType = N.  Network
                                             ' defined on Network system.
                                             ' Network defined on Rep system.
                                             ' Y=Yes; N=No
    sFTPAddress           As String * 20     ' FTP address.  nrfType = N.
                                             ' Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    iFTPPort              As Integer         ' Port number.  nrfType = N.
                                             ' Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sFTPImportDir         As String * 120    ' FTP Import Directory.  nrfType =
                                             ' N.  Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sFTPExportDir         As String * 120    ' FTP export directory.  nrfType =
                                             ' N.  Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sFTPUserID            As String * 20     ' FTP User ID.  nrfType = N.
                                             ' Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sFTPUserPW            As String * 20     ' FTP User password.  nrfType = N.
                                             ' Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sIISRootURL           As String * 120    ' IIS Root URL.  nrfType = N.
                                             ' Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sIISRegSection        As String * 80     ' IIS Register Section.  nrfType =
                                             ' N.  Network defined on Network
                                             ' system.  Network defined on Rep
                                             ' system
    sTransmitPrice        As String * 1      ' Transmit Price from Rep to Net (Y/N).  Test for N.
    sUnused               As String * 29     ' Unused
End Type


'Type NRFKEY0
'    iCode                 As Integer
'End Type


'******************************************************************************
' QIF_Queue_Info Record Definition
'
'******************************************************************************
Type QIF 'VBC NR
    lCode                 As Long            ' Auto-Increment 'VBC NR
    sNRProcessed          As String * 1      ' Net-Rep Processed this record (Y 'VBC NR
                                             ' or N)
    lGenDate              As Long            ' Date that is stored within the 'VBC NR
                                             ' transferred text file
    lGenTime              As Long            ' Time that is stored within the 'VBC NR
                                             ' transferred text file
    sType                 As String * 1      ' Type: T= Transmit to SQL Server; 'VBC NR
                                             ' S=Save to Pervasive
    sSubType              As String * 1      ' Kind of record to be transmitted 'VBC NR
                                             ' or saved (C=Contract/Lines;
                                             ' S=Posted Spots; M=Messages; A=6
                                             ' month avails)
    sFileName1            As String * 80     ' File name. Sample 'VBC NR
                                             ' aaa_dbid_CntrNo_GenDate_GenTime.c
                                             ' sv; aaa_dbid_GenDate_GenTime.csv.
                                             ' aaa=Rep or Net; dbid=RepDBID or
                                             ' NetDBID
    sFileName2            As String * 80     ' File name. Sample 'VBC NR
                                             ' aaa_dbid_CntrNo_GenDate_GenTime.c
                                             ' sv; aaa_dbid_GenDate_GenTime.csv.
                                             ' aaa=Rep or Net; dbid=RepDBID or
                                             ' NetDBID
    sFileName3            As String * 80     ' File name. Sample 'VBC NR
                                             ' aaa_dbid_CntrNo_GenDate_GenTime.c
                                             ' sv; aaa_dbid_GenDate_GenTime.csv.
                                             ' aaa=Rep or Net; dbid=RepDBID or
                                             ' NetDBID
    lCntrNo               As Long            ' Contract number (qifType = T and qifSubtype = C) 'VBC NR
    iNrfCode              As Integer         ' NRF reference used to know which 'VBC NR
                                             ' server is to receive the text
    iProcessedDate(0 To 1) As Integer        ' Date processed 'VBC NR
    iProcessedTime(0 To 1) As Integer        ' Processed time 'VBC NR
    sUnused               As String * 20     ' Unused 'VBC NR
End Type 'VBC NR


'Type QIFKEY0
'    lCode                 As Long
'End Type

Type QIFKEY1 'VBC NR
    sType                 As String * 1 'VBC NR
    sNRProcessed          As String * 1 'VBC NR
    lGenDate              As Long 'VBC NR
    lGenTime              As Long 'VBC NR
End Type 'VBC NR

