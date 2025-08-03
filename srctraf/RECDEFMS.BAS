Attribute VB_Name = "RECDEFMS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RECDEFMS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  MNFKEY3                                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RecDefMS.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions M- Z
Option Explicit



'********************************************************
'
'Media code file definition
'
'*********************************************************

'Mcf record layout
Type MCF
    iCode As Integer        'Internal code number for Media code
    sName As String * 6     'Identification name
    sPrefix As String * 6   'Prefix on log
    sReuse As String * 1    'Reuse number (Y/N)
    sScript As String * 1   'Scripted allowed (Y/N)
    'sTempMedia As String * 1    'Temporary Media #'s
    sSortCart As String * 1    'D=Last Used date; C=Cart #
    sCartDisp As String * 1 'Cart Disposition (N=N/A; S=Save; P=Purge; A=Ask after Expired)
    '2/2/12: Replaced mcfTapeDisp with mcfSuppressOnExport. Use default of N
    'sTapeDisp As String * 1 'Tape Disposition (N=N/A; R=Return; D=Destroy; A=Ask after Expired)
    sSuppressOnExport As String * 1 'Suppress on Automation Export (Y=Yes, N=No).  Test for Y
    sAssignNo(0 To 1) As String * 5    'Start number if reusing
    'iGroupNo As Integer     'Group number
    iVefCode As Integer     'Vehicle code if assigning media codes to vehicles
End Type

'Mcf key record layout- use INTKEY0
'Type MCFKEY0
'    iCode As Integer
'End Type

'******************************************************************************
' MEF_Media_Extra Record Definition
'
'******************************************************************************
Type MEF
    iCode                 As Integer         ' Auto Increment
    iMcfCode              As Integer         ' Media Definition reference code
    sPrefix               As String * 6      ' iPump Media prefix for Cart #.
    sSuffix               As String * 6      ' iPump Media suffix for Cart #.
    sEventType            As String * 2      ' iPump Event Type (PL or ID)
    sNetworkID            As String * 2      ' iPump Network ID (KL or A1)
    sNameSpace            As String * 60     ' iPump Folder location that the
                                             ' audio will be stored within.
    sUnused               As String * 10     ' Unused
End Type


'Mef key record layout- use INTKEY0
'Type MEFKEY0
'    iCode                 As Integer
'End Type

Type MEFKEY1
    iMcfCode              As Integer
End Type


'******************************************************************************
' MGF_Media_Sold_Game Record Definition
'
'******************************************************************************
Type MGF
    lCode                 As Long            ' Auto Increment field
    lMsfCode              As Long            ' MSF Reference (Media Sold).
    lIsfCode              As Long            ' Inventory Specification Reference
    iGameNo               As Integer         ' Game Number
    iNoUnits              As Integer         ' Number of units
    lRate                 As Long            ' Purchase Rate (xxxxx.xx).
    lCost                 As Long            ' Cost (xxxxx.xx)
    sBilled               As String * 1      ' Billed (Y/N)
    sUnused               As String * 10     ' Unused
End Type


'Type MGFKEY0-Use LONGKEY0
'    lCode                 As Long
'End Type

Type MGFKEY1
    lMsfCode              As Long
    iGameNo               As Integer
End Type

Type MGFLIST
    MgfRec As MGF
    iStatus As Integer
    iNextMgf As Integer
End Type

'********************************************************
'
'Multi-Name file definition
'
'********************************************************

'Multi file record layout
Type MNF
    iCode As Integer    'Internal code number for multi-name
    sType                 As String * 1      ' A=Announcer;B=BusCat;C=ProductPro
                                             ' tection;D=Demo;E=Genre;F=Research
                                             ' ;G=SalesRegion;H=VehGroup;I=NTR
                                             ' Type;J=Terms;K=
                                             ' Seg;L=Lang;M=MissedReason;N=FeedT
                                             ' ype;O=Compet;P=Sport;R=RevSet;S=Sal
                                             ' eSource;T=SaleTeam;V=InvSort;X=Ex
                                             ' clus.;Y=TransType;Z=SportTeam
    sName                 As String * 50     ' Name
    sRPU As String * 5  'Rate as money (Pack Decimal Number, 2 places after dec point)
    sUnitType As String * 6 'Unit type with "per" removed
    sSSComm As String * 4   'Commission (Pack Decimal Number, 4 places after dec point)
    iMerge As Integer   'Merge code number
    iGroupNo As Integer 'If "A": Group number or if "S": sales origin (1=Local; 2=Regional; 3=National)
    sCodeStn As String * 5  '"C" and "R" codes assigned by station
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode  Note:in the code, but no longer used
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode  Note:in the code, but no longer used
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)  Note:in the code, but no longer used
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)  Note:in the code, but no longer used
    sUnitsPer As String * 15    'NTR Type (I) = Units Per ------
    lCost As Long       'Type = "I" (NTR):  Acquisition Cost (xxxxxx.xx)
    sUnused As String * 20
End Type

'Multi file key record layout- use INTKEY0
'Type MNFKEY0
'    iCode As Integer
'End Type

'Type MNFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

Type MNFKEY3 'VBC NR
    sType As String * 1 'VBC NR
End Type 'VBC NR

'Sales source code extract record
Type MNFEXT
    iCode As Integer    'Internal code number for multi-name
    sName As String * 20  'Name
    iGroupNo As Integer 'If "A": Group number or if "S": sales origin (1=Local; 2=Regional; 3=National)
End Type
Public Const MNFEXTPK As String = "IB20I"

Type MNFCOMPEXT
    iCode As Integer    'Internal code number for multi-name
    sName As String * 20  'Name
    sUnitType As String * 6 'Unit type with "per" removed
End Type
Public Const MNFCOMPEXTPK As String = "IB20B6"

'******************************************************************************
' MSF_Media_Sold Record Definition
'
'******************************************************************************
Type MSF
    lCode                 As Long            ' Autoincrement Reference
    lChfCode              As Long            ' Contract Reference Code
    lghfcode              As Long            ' Game Header Reference code
    iVefCode              As Integer         ' Vehicle Reference code
    iIhfCode              As Integer         ' Type/Item header reference (get
                                             ' itf and iif from header)
    iInvGameNo            As Integer         ' Media Sold Game number that match
                                             ' Teams, Dates andTimes with this
                                             ' image
    lCxfCode              As Long            ' Comment
    sUnused               As String * 20     ' Unused
End Type


'Type MSFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type MSFKEY1
    iVefCode              As Integer
End Type

Type MSFKEY2
    lChfCode              As Long
End Type

Type MSFLIST
    MsfRec As MSF
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    iFirstMgf As Integer
    iCxfIndex As Integer    'Index for comment within sgMsfCntrCxf
End Type


'*********************************************************
'
'   NIF - Network Inventory
'
'*********************************************************
Type NIF
    lCode As Long                   'internal code (autoinc)
    iVefCode As Integer             'vehicle code
    iYear As Integer                'year (i.e. 2005)
    sInvWkYear As String * 1        'Inv by wk (W) or Yr (Y)
    sAllowRollover As String * 1    'allow unused inventory to rollover (Y/N)
    'lInvCount(1 To 53) As Long      'time in minutes/seconds
    lInvCount(0 To 52) As Long      'time in minutes/seconds
    iEnterDate(0 To 1) As Integer   'date entered
    iUrfCode As Integer             'user entered inventory
    sUnused As String * 20          'unused
End Type

'NIF key record layout- use INTKEY0
Type NIFKEY1
    iVefCode As Integer    'Vehicle code
    iYear As Integer       'Inventory year
End Type
' Dan M 8-21-08 for adding NTRs to pacing reports
Type NTRPacing
    lSBFTotal As Long
    lSbfDate As Long
    iVefCode As Integer
    bGarbage As Boolean
    slNTRAgyCommFlag As String * 1         '3-28-19
End Type

'******************************************************************************
' PAF_Program_Air_Info Record Definition
'
'******************************************************************************
Type PAF
    lCode                 As Long            ' Program Air Information Reference
                                             ' Code
    iVefCode              As Integer         ' Vehicle Reference Code
    sName                 As String * 20     ' Program Name
    iStartDate(0 To 1)    As Integer         ' Program definition start date
    iEndDate(0 To 1)      As Integer         ' Program definition end date. TFN
                                             ' = 12/31/2069
    sMo                   As String * 1      ' Air on Monday (Y/N).
    sTu                   As String * 1      ' Air on Tuesday (Y/N)
    sWe                   As String * 1      ' Air on Wednesday (Y/N).
    sTh                   As String * 1      ' Air on Thursday (Y/N)
    sFr                   As String * 1      ' Air on Friday (Y/N)
    sSa                   As String * 1      ' Air on Saturday (Y/N)
    sSu                   As String * 1      ' Air on Sunday
    iStartTime(0 To 1)    As Integer         ' Program definition start time
    iEndTime(0 To 1)      As Integer         ' Program definition end time
    iUrfCode              As Integer         ' User Refernce Code
    sUnused               As String * 20     ' Unused
End Type


Type PAFKEY0
    lCode                 As Long
End Type

Type PAFKEY1
    iVefCode              As Integer
    iEndDate(0 To 1)      As Integer
End Type

Type PAFKEY2
    iVefCode              As Integer
    iStartDate(0 To 1)    As Integer
End Type


'******************************************************************************
' PDF_Invoices_PDF Record Definition
'
'******************************************************************************
Type PDF
    lCode                 As Long            ' PDF file internal reference code
    iAdfCode              As Integer         ' Direct Advertiser internal
                                             ' reference code
    iAgfCode              As Integer         ' Agency internal reference code
    sName                 As String * 40     ' Advertiser/Agency Name of person
                                             ' to recieve the PDF from Invoicing
    sPhone                As String * 20     ' Advertiser/Agency Phone Number of
                                             ' person to recieve the PDF from
                                             ' Invoicing
    sEMailAddress         As String * 80     ' Advertiser/Agency E-Mail address
                                             ' of person to recieve the PDF from
                                             ' Invoicing
    sUnused               As String * 10     ' Unused
End Type


'Pdf key record layout- use LONGKEY0
'Type PDFKEY0
'    lCode                 As Long
'End Type

'Pdf key record layout- use INTKEY0
'Type PDFKEY1
'    iAdfCode              As Integer
'End Type

'Pdf key record layout- use INTKEY0
'Type PDFKEY2
'    iAgfCode              As Integer
'End Type

'********************************************************
'
'Past history Receivable file definition- This file is the same as RVF
'
'*********************************************************

'Phf record layout- using RVF
'Type PHF
'    lCode As Long    'Internal code number (AutoInc)
'    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
'    iAdfCode As Integer 'Advertiser code number
'    lPrfCode As Long    'Product code
'    iSlfCode As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
'    lCntrNo As Long     'Contract Number
'    lInvNo As Long      'Invoice number
'    lRefInvNo As Long   'Reference invoice number
'    iAirVefCode As Integer 'Vehicle Code of airing vehicle
'    lUnused2 As Long    'was Check number
'    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
'    sTranType As String * 2   '
'    sAction As String * 1   '
'    sGross As String * 6 'Gross amount (xx,xxx,xxx.xx)
'    sNet As String * 6   'Net amount (xx,xxx,xxx.xx)
'    iAgePeriod As Integer   'Aging period
'    iAgingYear As Integer   'Aging year
'    sCashTrade As String * 1    'C=Cash; T=Trade
'    iPurgeDate(0 To 1) As Integer  'Purge Date or zero if not purged
'    iUrfCode As Integer 'Last user who modified receivable
'    iBillVefCode as Integer    'Vehicle code of billing vehicle
'    iPkLineNo As Integer       'Package line no used to combine transactions
'    iInvDate(0 To 1) As Integer  'Invoice Date
'    iDateEntrd(0 To 1) As Integer  'Created Date
'    iRemoteID As Integer        'Remote ID (Note: Unique Cntr# is Cntr#-RemoteID)
'    iMnfGroup As Integer   'Participant MnfCode, set to zero except if sales source is ask
'    iTax1 as long          'sales tax #1 12-27-01
'    iTax2 as long          'sales tax #2 12-27-01
'    iMnfItem As Integer    'NTR Type (only set for TranType of IM) 9/10/02
'    sInCollect As String * 1    'In Collection (Y/N), 9/27/02
'    lCefCode As Long        'Transaction Comment (for On Account)
'    lSbfCode As Long        'Pointer to SBF if REP or NTR invoicing (added 12/13/02)
'    lAcquisitionCost as Long    'Acquisition Cost (xxxxxx.xx)
'    iBacklogTrfCode As Integer  'Tax table code for backlog NTR
'    sType As String * 1         'Type of record: Blank=Aired and NTR dollars for contracts w/o Installment;
'                                'I=Installment amount from contract; A=Air and NTR amount for contract with installment
'    sInvoiceUndone As String * 1    'Invoice Undone (N/Y).  Test for Y
'    iPnfBuyer as Integer
'    lGsfCode              As Long            ' Sport Event Reference code
'    sCheckNo              As String * 10     ' Check Number. replaced Long with string
'    lPcfCode              As Long            ' Podcast Ad Server contract reference
'                                             ' code
'    sUnused               As String * 6      ' Unused
'End Type

'Phf key record layout- using Rvf
'Type PHFKEY0
'    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
'End Type

'Phf key record layout- using Rvf
'Type PHFKEY1
'    iAdfCode As Integer 'Advertiser code number
'End Type

'Phf key record layout- use LONGKEY0
'Type PHFKEY2
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Phf Key record layout
'Type PHFKEY3
'    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
'End Type

'Phf key record layout
'Type PHFKEY4
'    lCntrNo As Long     'Contract Number
'    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
'End Type

'Phf key record layout
'Type PHFKEY5
'    lInvNo As Long      'Invoice number
'End Type


'Past history Receivable extract record- using Rvf
'Type PHFEXT
'    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
'    iAdfCode As Integer 'Advertiser code number
'End Type

'Past history Receivable Transaction extract record- using Rvf
'Type PHFTRANEXT
'    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
'    iAdfCode As Integer 'Advertiser code number
'    lInvNo As Long      'Invoice number
'    iAirVefCode As Integer 'Vehicle Code of airing vehicle
'    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
'    sTranType As String * 2   '
'    sCashTrade As String * 1    'C=Cash; T=Trade
'    iBillVefCode as Integer    'Vehicle code of billing vehicle
'    iPkLineNo As Integer       'Package line no used to combine transactions
'End Type


'********************************************************
'
'Projection file definition
'
'*********************************************************

'Pjf record layout
Type PJF
    lCode As Long    'Internal code number (AutoInc)
    iSlfCode As Integer    'Salesperson code
    iSofCode As Integer    'Selling Office code
    iYear As Integer
    iMnfBus As Integer      'Mnf Potential Business Category
    iEffDate(0 To 1) As Integer   'Effective Date (entered date)
    lCxfChgR As Long        'Comment code number for change reason
    iAdfCode As Integer 'Advertiser code number
    lPrfCode As Long 'Product Code
    iMnfDemo As Integer     'Mnf Demo Code Number
    iVefCode As Integer 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    sNewRet As String * 1   'R=Return; N=New
    lChfCode As Long
    lGross(0 To 53) As Long   'Gross budget amount (xxxxxxx.xx)
    iRolloverDate(0 To 1) As Integer    'Roll over date
    iEffTime(0 To 1) As Integer 'Effective Time (entered time)
    iUrfCode As Integer     'Last user to change record
    iRemoteID As Integer        'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    lAutoCode As Long        'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iSourceID As Integer        'Remote User ID (avoid resending to sender)
    sUnused As String * 20
End Type

'Pjf key record layout- use INTKEY0
Type PJFKEY0
    iSlfCode As Integer    'Salesperson code
    iRolloverDate(0 To 1) As Integer   'Effective Date (entered date)
End Type

'Pjf key record layout- use LONGKEY0
'Type PJFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Type PJFKEY2- use LONGKEY0
'    lChfCode As Long
'End Type


'********************************************************
'
'Personnel file definition
'
'*********************************************************

'Pnf record layout
Type PNF
    iCode As Integer    'Autoincrement
    iAgfCode As Integer 'Agency Code (zero if defined for direct advertiser)
    iAdfCode As Integer 'Advertiser code (zero if not direct advertiser)
    sType As String * 1 'B=Buyer; P=Payable Clerk
    sName As String * 40    'Name of Buyer or Payable Clerk
    sTitle As String * 20   'Title
    sPhone As String * 14   'Phone number and extension
    sFax As String * 10 'Fax number
    lCxfCode As Long    'Comment
    sState As String * 1    'A=Active; D=Dormant
    iAddedDate(0 To 1) As Integer   'Date record added (used in ClnFile if both adfcode and agfcode are
                                    'zero and date in past, then the record is removed
                                    'as it was not associated correctly with either adf or agf
    iRemoteID As Integer        'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    iAutoCode As Integer        'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iSourceID As Integer    'Remote User ID (avoid resending to sender)
    lEMailCefCode As Long   'E-Mail reference pointer
    sCntrAddr(0 To 2) As String * 40    'Contract address to be used instead of advertiser contract address if defined
    sBillAddr(0 To 2) As String * 40    'Bill address to be used instead of advertiser bill address if defined
    sUnused As String * 20
End Type

'PNf key record layout- use INTKEY0
'Type PNFKEY0
'    iCode As Integer
'End Type

'Pnf key record layout
Type PNFKEY1
    iAgfCode As Integer 'Agency Code (zero if defined for direct advertiser)
End Type

'Pnf key record layout
Type PNFKEY2
    iAdfCode As Integer 'Advertiser Code (zero if defined for agency)
End Type

'Type PNFKEY3- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type



'********************************************************
'
'Product Names file definition
'
'*********************************************************

'Prf record layout
Type PRF
    lCode As Long           'Autoincrement number
    iAdfCode As Integer     'Advertiser code number
    sName As String * 35    'Product name
    iMnfComp(0 To 1) As Integer 'Competitive code
    iMnfExcl(0 To 1) As Integer 'Program Exclusions code
    iPnfBuyer As Integer    'Personnel Buyer Code number
    sCppCpm As String * 1    'P=CPP; M=CPM; N=N/A (was sPriceType)
    iMnfDemo(0 To 3) As Integer 'Mnf Demo copy number
    lTarget(0 To 3) As Long  'CPP or CPM target (xxxxx.xx)
    lLastCPP(0 To 3) As Long  'Last CPP (xxxxx.xx)
    lLastCPM(0 To 3) As Long  'Last CPM (xxx.xx)
    sState As String * 1        'A=Active; D=Dormant
    iUrfCode As Integer     'Last user who modified product names
    iMnfBus As Integer  'Business Category
    iRemoteID As Integer        'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    lAutoCode As Long        'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iSourceID As Integer    'Remote User ID (avoid resending to sender)
    sUnused As String * 20
End Type

'Prf key record layout- use LONGKEY0
'Type PRFKEY0
'    lCode As Long
'End Type

'Prf key record layout
Type PRFKEY1
    iAdfCode As Integer     'Advertiser code number
End Type

'Type PRFKEY2- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type


'********************************************************
'
'Standard Package Vehicles file definition
'
'*********************************************************

'Pvf record layout
Type PVF
    lCode As Long        'Internal code number for user Option
    sName As String * 20    'Combo name
    lLkPvfCode As Long      'Link to Next in Chain of VSF records (0=No More)
    iVefCode(0 To 24) As Integer 'Vehicle code
    iRdfCode(0 To 24) As Integer 'Rate Card daypart code
    iNoSpot(0 To 24) As Integer   'Number of spots
    iPctRate(0 To 24) As Integer '% Rate Split (for vefStdPrice = 2) (xxx.xx)
    sType As String * 1         'Package Type. Blank or A = AirTime Only; S=Podcast Only; M=Mix AirTime and Podcast buy; C=Podcast Ad Server buy only
    sACT1LineupCode       As String * 11     ' Default Pkg Veh ACT1 Lineup Code
    sACT1StoredTime       As String * 1      ' Default Pkg Veh ACT1 Stored Time
                                             ' (T or Blank). Test for T
    sACT1StoredSpots      As String * 1      ' Default Pkg Veh ACT1 Stored Spots
                                             ' (S or Blank). Test for S
    sACT1StoreClearPct    As String * 1      ' Default Pkg Veh ACT1 Clearance
                                             ' Percent (C or Blank). Test for C
    sACT1DaypartFilter    As String * 1      ' Default Pkg Veh ACT1 Daypart
                                             ' Filter (F Blank). Test for F
    sUnused               As String * 5      ' Unused
End Type

'Pvf key record layout- use LONGKEY0
'Type PVFKEY0
'    lCode As Long
'End Type

Type PVFKEY1
    sType As String * 1         'Package Type. Blank or A = AirTime Only; S=Podcast Spots Only; M=Mix AirTime and Podcast Spot buy; C=Podcast CPM buy only
End Type

'********************************************************
'
'Rate Card file definition
'
'*********************************************************

'Rcf record layout
Type RCF
    iCode As Integer    'Internal code number for Vehicle
    sName As String * 4 'Rate Card #
    iVefCode As Integer 'Vehicle Code
    iStartDate(0 To 1) As Integer  'Start Date of Rate Card
    iEndDate(0 To 1) As Integer  'End Date of Rate Card or zero if not superseded
    iGridsUsed As Integer   '# of grid levels in use
    iTodayGrid As Integer   'Current Grid level in use
    iBaseLen As Integer 'Base length
    lRound As Long      'Round Rates to nearest $ (xxx.xx)
    sUseBB As String * 1    'D=Dollar; I=Index; N=N/A
    lBBValue As Long    '$ Or % for single billboard (xxxxxxxx.xx or xx.x)
    sUseFP As String * 1    'D=Dollar; I=Indext; N=N/A
    lFPValue As Long    'Fixed price surcharge or % (xxxxxxx.xx or xx.x)
    sUseNP As String * 1    'D=Dollar; P=Percent; N=N/A
    lNPValue As Long    'non-preemptive $ or % (xxxxxxx.xx or xx.x)
    sUseLen As String * 1  'A=Actual $; D=Delta dollar; I=Index; N=N/A
    'iLen(1 To 10) As Integer 'Spot length for length price factor
    iLen(0 To 9) As Integer 'Spot length for length price factor
    'lValue(1 To 10) As Long    'Spot length factor-sLenType indicates type of value (xxxxxxx.xx)
    lValue(0 To 9) As Long    'Spot length factor-sLenType indicates type of value (xxxxxxx.xx)
    sUseSpot As String * 1  'A=Actual $; D=Delta dollar; I=Index; N=N/A
    'iSpotMin(1 To 4) As Integer 'Min Spot frequency for discount
    iSpotMin(0 To 3) As Integer 'Min Spot frequency for discount
    'iSpotMax(1 To 4) As Integer 'Max Spot Frequency for discount
    iSpotMax(0 To 3) As Integer 'Max Spot Frequency for discount
    'lSpotVal(1 To 4) As Long  'Spot frequence discount (xxxxxxx.xx)
    lSpotVal(0 To 3) As Long  'Spot frequence discount (xxxxxxx.xx)
    sUseWeek As String * 1  'A=Actual $; D=Delta dollar; I=Index; N=N/A
    'iWkMin(1 To 4) As Integer   'Min week frequency for discount
    iWkMin(0 To 3) As Integer   'Min week frequency for discount
    'iWkMax(1 To 4) As Integer   'Max week frequency for discount
    iWkMax(0 To 3) As Integer   'Max week frequency for discount
    'lWkVal(1 To 4) As Long    'Week frequency discount (xxxxxxx.xx)
    lWkVal(0 To 3) As Long    'Week frequency discount (xxxxxxx.xx)
    sUseDay As String * 1  'A=Actual $; D=Delta dollar; I=Index; N=N/A
    'sDay(1 To 10, 1 To 7) As String * 1 '
    sDay(0 To 9, 0 To 6) As String * 1 '
    'lDyRate(1 To 10) As Long      'Day rate (xxxxxxx.xx)
    lDyRate(0 To 9) As Long      'Day rate (xxxxxxx.xx)
    sUseHour As String * 1  'A=Actual $; D=Delta dollar; I=Index; N=N/A
    'sHrDay(1 To 10, 1 To 7) As String * 1 '
    sHrDay(0 To 9, 0 To 6) As String * 1 '
    'iHrStartTime(0 To 1, 1 To 10) As Integer
    iHrStartTime(0 To 1, 0 To 9) As Integer
    'iHrEndTime(0 To 1, 1 To 10) As Integer
    iHrEndTime(0 To 1, 0 To 9) As Integer
    'lHrRate(1 To 10) As Long      'Day rate (xxxxxxx.xx)
    lHrRate(0 To 9) As Long      'Day rate (xxxxxxx.xx)
    iFltNo(0 To 53) As Integer      'Number of weeks within index (1 used as start)
    'lGridIndex(1 To 12) As Long   'Grid Index values (x.xxxx)
    lGridIndex(0 To 11) As Long   'Grid Index values (x.xxxx)
    iYear As Integer            'Year rate card to be used
    iUrfCode As Integer     'Last user who modified rate card
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    sUse1stPos As String * 1        'D=Dollar; I=Index; N=N/A
    l1stPosValue As Long            '1st position $ or % (xxxxxxx.xx or xx.x)
    sUseSoloAvail As String * 1        'D=Dollar; I=Index; N=N/A
    lSoloAvailValue As Long            'Solo Avail $ or % (xxxxxxx.xx or xx.x)
    sUsePrefDT As String * 1        'D=Dollar; I=Index; N=N/A
    lPrefDTValue As Long            'Preferred Days/Times $ or % (xxxxxxx.xx or xx.x)
    sUnused As String * 5
End Type

'Rcf key record layout- use INTKEY0
'Type RCFKEY0
'    iCode As Integer
'End Type

'Type RCFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

'Rate Card extract record


'********************************************************
'
'Daypart file definition
'
'*********************************************************

'Rdf record layout
Type RDF
    iCode As Integer    'Internal code number for Rate Card program
    sName As String * 20    'Program/Time Name
    sBase As String * 1 'Base daypart (Y=Yes, N=No)
    sReport As String * 1 'Include as Daypart on reports (Y=Yes, N=No)
    sInOut As String * 1    'N=N/A; I=Book into avail specified; O=Exclude avail specified
    ianfCode As Integer     'Avail name code if sInOut = I or O
    sTimeOver As String * 1 'Y=OK to override times specified; N=Times can't to override
    sState As String * 1    'A=Active; D=Drmant
    'Start Time, End Time , Week day and % spot only valid if NO library specified
    'The values are stored in a shifted format- high number indicies filled first
    'i.e. If two times defined, the first time will be stored in index 6 and the second time will
    '     be stored as index 7.  The first five indicies will be set to zero and blank.
    'iStartTime(0 To 1, 1 To 7) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iStartTime(0 To 1, 0 To 6) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        'If Hund Sec = 1 and all other times = 0, then time not defined
                        'In Basic the left most dimension expands first
    'iEndTime(0 To 1, 1 To 7) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iEndTime(0 To 1, 0 To 6) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    'sWkDays(1 To 7, 1 To 7)  As String * 1  'Dimension 1: Time; Dimension 2:Days;
    sWkDays(0 To 6, 0 To 6)  As String * 1  'Dimension 1: Time; Dimension 2:Days;
                                            'Dimension 2: Index 1= Monday, 2= tuesday,...
                                            'Weekday flag: Y=day allowed; N=Day disallowed
    'iSpotPct(1 To 7) As Integer  'Percentage of spots (xxx.xx)
    iSpotPct(0 To 6) As Integer  'Percentage of spots (xxx.xx)
    sContrRef As String * 1   'Y=Referenced by contract; N=Not Referenced
    iLtfCode(0 To 2) As Integer    'Library code (set is automatically, 0=manual)
    iSortCode As Integer    'Sort code number
    iRemoteID As Integer        'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    iAutoCode As Integer        'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iSourceID As Integer        'Remote User ID (avoid resending to sender)
    sOverridePrtRules As String * 1 'Override Print Rules: D=Include Daypart Name; A=Include Avail Name; N= Meither name included
    iDPGroupMnfCode       As Integer         ' Daypart Group mnf internal
                                             ' reference code. When defined,
                                             ' Demo Bar daypart totals by group
                                             ' name
    sUnused As String * 17
End Type

'Rdf key record layout- use INTKEY0
'Type RDFKEY0
'    iCode As Integer
'End Type

'Type RDFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

'Rate Card Dayparts extract record


'******************************************************************************
' RHF_Reconcile_Hist Record Definition
'
'******************************************************************************
Type RHF
    iCode                 As Integer         ' Auto Increment field
    sSourceType           As String * 1      ' Source Type (R=Reconcile; U=Undo
                                             ' Invoicing-Rollback)
    sRRP                  As String * 1      ' Reconcile period (S=Standard
                                             ' broadcast; C=Calendar;
                                             ' F=Fiscal-Corporate)
    iRPRP(0 To 1)         As Integer         ' End of Previous Reconcile period
    iRCRP(0 To 1)         As Integer         ' End of Current reconcile period
    iRNRP(0 To 1)         As Integer         ' End of next reconcile period
    sRB                   As String * 6      ' Reconcile Balance, End of
                                             ' Previous reconcile period
    iUrfCode              As Integer         ' User reference code
    iDateEntered(0 To 1)  As Integer         ' Date entered
    sUnused               As String * 10     ' Unused
End Type


'Rhf key record layout- use INTKEY0
'Type RHFKEY0
'    iCode                 As Integer
'End Type

'********************************************************
'
'Rate Card Item file definition
'
'*********************************************************

'Rif record layout
Type RIF
    lCode As Long    'Internal code number (AutoInc)
    iRcfCode As Integer 'Rate Card # code
    iVefCode As Integer 'Vehicle Code
    iRdfCode As Integer
    iYear As Integer    'Year
    'lRateIndex As Long    'Rate Index (xx.xxx)
    sBase As String * 1     'Base Daypart (Y or N)
    sRpt As String * 1      'Show on Reports (Y or N)
    iSort As Integer        'Sort code within vehicle
    iROSPct As Integer   'ROS % (xxx.xx)
    lRate(0 To 53) As Long    'Rate for each flight index (xxxxxxx.xx)
    iRateChgDate(0 To 1) As Integer 'Rate Changed date used for alerts
    iRateChgTime(0 To 1) As Integer 'Rate Change Time used for alerts
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    lAutoCode As Long    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    lAcquisitionCost As Long 'Default schedule line acquisition cost if using the Paperwork Distribution system
    lPodCPM               As Long            ' Podcast CPM values (xxxx.xx)
    sUnused               As String * 12
End Type

'Rif key record layout- use INTKEY0
'Type RIFKEY0
'    iRcfCode As Integer 'Rate Card # code
'End Type

'Rif key record layout- use LONGKEY0
'Type RIFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Type RIFKEY2- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type


'********************************************************
'
'Record Lock file definition
'
'*********************************************************

Type RLF
    lCode As Long                   'Autoincrement field
    iUrfCode As Integer             'User that locked record
    sType    As String * 1      ' C=Contract;S=SSF;I=Invoice;M=Copy
                                ' ;R=Collect;Y=Purge
    sSubType As String * 1      ' For Type = C: C=Altering;
                                ' ""=Scheduling; Type=S:C=Contract
                                ' Scheduling
                                ' Task;P=Posting;""=Spots; Type =I:
                                ' U=Undo;""=Final;
                                ' Type=M:""=Inventory;Type=R:R=Reco
                                ' ncile;Z=Zero-Purge;
                                ' Type=Y:1=Purge;2=Set Copy Dates
                                    'for Type = S: SSFCode
    lRecCode              As Long            ' for Type = C: ChfCode; for Type =
                                             ' S: SSFCode
    iEnteredDate(0 To 1) As Integer 'Entered date
    iEnteredTime(0 To 1) As Integer
    sUnused As String * 10
End Type

'Rlf key record layout- use LONGKEY0
'Type RLFKEY0
'    lCode As Integer
'End Type

Type RLFKEY1
    sType As String * 1             'C=Contract; S=SSF
    lRecCode As Long                'for Type = C: ChfCode
                                    'for Type = S: SSFCode
End Type

Type RLFVIEW
    sKey As String * 20
    tRlf As RLF
End Type



'********************************************************
'
'Receivable file definition
'
'*********************************************************

'Rvf record layout.  PHF must match RVF layout exactly; RVR must match the record layout in the first part of the record design
Type RVF
    lCode As Long    'Internal code number (AutoInc)
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
    lPrfCode As Long    'Product code
    iSlfCode As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    lCntrNo As Long     'Contract Number
    lInvNo As Long      'Invoice number
    lRefInvNo As Long   'Reference invoice number
    iAirVefCode As Integer 'Vehicle Code of Airing vehicle
    lUnused2 As Long    'was Check number
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
    sTranType             As String * 2      ' IN=Invoice, AN=Invoice
                                             ' Adjustment, PI=Payment to
                                             ' Invoice, PO=Payment On Account,
                                             ' WB=Write-Off Bad Debt,
                                             ' WV=Write-Off Variance. W+ some
                                             ' kind of write off
    sAction As String * 1   '
    sGross As String * 6 'Gross amount (xx,xxx,xxx.xx)
    sNet As String * 6   'Net amount (xx,xxx,xxx.xx)
    iAgePeriod As Integer   'Aging period
    iAgingYear As Integer   'Aging year
    sCashTrade As String * 1    'C=Cash; T=Trade; M=Merchandising; P=Promotion
    iPurgeDate(0 To 1) As Integer  'Purge Date or zero if not purged
    iUrfCode As Integer 'Last user who modified receivable
    iBillVefCode As Integer    'Vehicle code of billing vehicle
    iPkLineNo As Integer       'Package line no used to combine transactions
    iInvDate(0 To 1) As Integer  'Invoice Date
    iDateEntrd(0 To 1) As Integer  'Created Date
    iRemoteID As Integer        'Remote ID (Note: Unique Cntr# is Cntr#-RemoteID)
    iMnfGroup As Integer   'Participant MnfCode, set to zero except if sales source is ask
    lTax1 As Long          'sales tax #1 12-27-01
    lTax2 As Long          'sales tax #2 12-27-01
    iMnfItem As Integer    'NTR Type (only set for TranType of IM) 9/10/02
    sInCollect As String * 1    'In Collection (Y/N), 9/27/02
    lCefCode As Long        'Transaction Comment (for On Account)
    lSbfCode As Long        'Pointer to SBF if REP or NTR invoicing (added 12/13/02)
    lAcquisitionCost As Long    'Acquisition Cost (xxxxxx.xx)
    iBacklogTrfCode As Integer  'Tax table code for backlog NTR
    sType As String * 1         'Type of record: Blank=Aired and NTR dollars for contracts w/o Installment;
                                'I=Installment amount from contract; A=Air and NTR amount for contract with installment
    sInvoiceUndone As String * 1    'Invoice Undone (N/Y).  Test for Y
    iPnfBuyer As Integer        'Buyer reference
    lGsfCode              As Long            ' Sport Event Reference code
    sCheckNo              As String * 10     ' Check Number. replaced Long with string
    lPcfCode              As Long            ' Podcast Ad Server contract reference
                                             ' code
    sUnused               As String * 6      ' Unused
End Type

'Rvf key record layout
Type RVFKEY0
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
End Type

'Rvf key record layout
Type RVFKEY1
    iAdfCode As Integer 'Advertiser code number
End Type

'Rvf key record layout- use LONGKEY0
'Type RVFKEY2
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Rvf Key record layout
Type RVFKEY3
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
End Type

'Rvf key record layout
Type RVFKEY4
    lCntrNo As Long     'Contract Number
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
End Type

'Rvf key record layout
Type RVFKEY5
    lInvNo As Long      'Invoice number
End Type

'Receivable extract record
Type RVFEXT
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
End Type
Public Const RVFEXTPK As String = "II"

'Receivable Transaction extract record
Type RVFTRANEXT
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
    lInvNo As Long      'Invoice number
    iAirVefCode As Integer 'Vehicle Code of Airing vehicle
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
    sTranType As String * 2   '
    sCashTrade As String * 1    'C=Cash; T=Trade
    iBillVefCode As Integer    'Vehicle code of billing vehicle
    iPkLineNo As Integer       'Package line no used to combine transactions
    iMnfItem As Integer         '7-7-06 NTR type
    iMnfGroup As Integer        '2/7/07: Participant
    sType As String * 1         'Type of record: Blank=Aired and NTR dollars for contracts w/o Installment;
                                'I=Installment amount from contract; A=Air and NTR amount for contract with installment
    lCntrNo As Long     'Contract Number
End Type
Public Const RVFTRANEXTPK As String = "IILIIIB2BIIIIBL"

Type RVFUNPURGE
    lCode As Long
    sTranType As String * 2   '
    sCashTrade As String * 1    'C=Cash; T=Trade
    sNet As String * 6   'Net amount (xx,xxx,xxx.xx)
    lTax1 As Long          'sales tax #1 12-27-01
    lTax2 As Long          'sales tax #2 12-27-01
End Type
Public Const RVFUNPURGEPK As String = "LB2BB6LL"

'******************************************************************************
' RWF_Rep_By_Week Record Definition
'
'******************************************************************************
Type RWF
    lCode                 As Long            ' Auto Code
    lSbfCode              As Long            ' SBF Code reference
    'iWkNoSpots(1 To 5)    As Integer         ' Week 1-5 number of spots
    'iWkNoCarried(1 To 5)  As Integer         ' Number of spots airing as MG/Outsides into this week from another week.
    'iWkOrderSpotNo(1 To 5)    As Integer     ' Week 1-5 number of ordered spots
    iWkNoSpots(0 To 4)    As Integer         ' Week 1-5 number of spots
    iWkNoCarried(0 To 4)  As Integer         ' Number of spots airing as MG/Outsides into this week from another week.
    iWkOrderSpotNo(0 To 4)    As Integer     ' Week 1-5 number of ordered spots
    sUnused               As String * 20     ' Unused
End Type


'Type RWFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type RWFKEY1
    lSbfCode              As Long
End Type


'********************************************************
'
'Special Bill file definition
'
'*********************************************************

'Sbf record layout
Type SBF
    lCode As Long    'Internal code number (AutoInc)
    lChfCode As Long            'Contract code
    iDate(0 To 1) As Integer    'Bill Date of line
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    sTranType As String * 1     'Transaction type: I=Item bill; F=Installment; T=Import Invoice
    iBillVefCode As Integer     'Bill Vehicle Code (Ordered)
    iMnfItem As Integer         'Item billing type code [I]
    iNoItems As Integer         'Number of items [I] or # of Spots [T]
    lGross As Long              'Amount/item [I] or Gross Amount {F or T](xxxxxxx.xx)
    iCalCarryBonus As Integer   'Calendar bonus carry over for this month into next month(time posting only) [T]  9/10/02
    sUnused1 As String * 4      'Unused (changed from UnitName to Unused 12/13/02 because UnitsPer replaced it)
    iPrintInvDate(0 To 1) As Integer    'Print Rep Invoice Date (added 12/13/02)
    sDescr As String * 80       'Description
    sAgyComm As String * 1      'Agency commissionable (Y/N)- commission from agency file [I]
    sSlsComm As String * 1      'Salesperson commissionable (Y/N) % from item bill field [I]
    'sSlsTax As String * 1       'Sales taxable (Y/N)- % from site table [I]
    sUnused2 As String * 1      'Unused
    sBilled As String * 1       'Billed Flag: Y=Yes; N=No
    sCashTrade As String * 1    'Cash (C) or Trade (T) [T]
    iAirVefCode As Integer      'Airing Vehicle Code [T]
    iAirNoSpots As Integer      'Airing number of spots [T]
    iBonusNoSpots As Integer    'Bonus number of spots [T]
    'lTax1 As Long               'Tax 1 amount (xxxxxxx.xx) [T]
    'lTax2 As Long               'Tax 2 amounr (xxxxxxx.xx) [T]
    iTrfCode As Integer         'Tax rate reference
    sUnused3 As String * 6      'Unused
    iImportDate(0 To 1) As Integer  'Date imported [T]
    iExportDate(0 To 1) As Integer  'Date exported [T]
    lRefInvNo As Long           'Refernced Invoice number [T]
    iCombineID As Integer       'Can to used to recreate tgRvfByVef [T]
    iUrfCode As Integer         'Last user who modified billing
    lOGross As Long             'Ordered Gross [T]
    iCommPct As Integer         'Agency Commission % [T]
    iRecDate(0 To 1) As Integer 'Received Date [T]  9/10/02
    iPostDate(0 To 1) As Integer 'Posted Date [T]  9/10/02
    sPostStatus As String * 1   'R=Received; P=Posted [T]  9/10/02
    iCalCarryOver As Integer    'Calendar carry over for this month into next month [T]  9/10/02
    lSpotPrice As Long          'Spot Price from CFF [T]  Used to match lines  9/10/02
    iMissCarryOver As Integer   'Missed carry over total from previous months into this month [T]  9/10/02
    sInserted As String * 1     'Manually Inserted (Y/N) [T] 12/27/02
    sBarterPaid As String * 1   'Barter Paid (Y/N)
    lAcquisitionCost As Long    'Acquisition Cost (xxxxxx.xx)
    iSpotLen As Integer         'Spot Length if Rep Posting (T) and Barter used
    iIhfCode As Integer         'Inventory Header reference code
    iLineNo As Integer          'Line number (T and Rep Posting by Time)
End Type

'Sbf key record layout
Type SBFKEY0
    lChfCode As Long            'Contract code
    iDate(0 To 1) As Integer    'Bill Date of line
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    sTranType As String * 1     'Transaction type: I=Item bill; F=Installment; T=Import Invoice
End Type

'Sbf key record layout- use LONGKEY0
'Type SBFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Sbf key record layout
Type SBFKEY2
    sTranType As String * 1     'Transaction type: I=Item bill; F=Installment; T=Import Invoice
    iDate(0 To 1) As Integer    'Bill Date of line
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
End Type

'Sbf key record layout
Type SBFKEY3
    sTranType As String * 1     'Transaction type: I=Item bill; F=Installment; T=Import Invoice
    iPostDate(0 To 1) As Integer 'Posted Date [T]  9/10/02
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
End Type

Type SBFKEY4
    iBillVefCode          As Integer
    sTranType             As String * 1
    iPostDate(0 To 1)     As Integer
End Type

Type SBFLIST
    SbfRec As SBF
    lRecPos As Long
    iStatus As Integer  '0=New; 1=old and retain, -1= New but not used
End Type

'6-9-14 moved again with sbf recd definition
Type SBFTypes                   '9-30-02
    iNTR As Integer             'include "I" SBFTypes
    iInstallment As Integer     'include "F" SBFTypes
    iImport As Integer          'include "T" SBFTypes
End Type


'********************************************************
'
'Spot detail file definition (psf must have the same image)
'
'*********************************************************

'Sdf record layout
Type SDF
    lCode As Long           'Spot Reference Number
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    iAdfCode As Integer     'Advertiser code number
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled; O=Outside contract limits (not a MG);
                                'H=Hidden (like cancel but does not show on invoices);
                                'For G and O, SMF contains specifications
    iMnfMissed As Integer   'Missed reason or Rotation #. Note -1 indicates ?
    sTracer As String * 1   'M=Mouse move, O=On demand & mouse moved, P=Created in post log, C=Line Change
                            'N=N/A, D=on Demand & created in post log, P=Library change caused spot to be moved
                            '1=Mouse move in between todays date and last log date, 2=Mouse move in post log, 3=Line change inbetween todays date and last log date
    sAffChg As String * 1   'Post Log:T=Time change, C=Copy change, B=Time and copy changed, blank=no change, A=Added
    sPtType As String * 1   'Copy code type (0=none assigned; 1=cif code; 2= ccfcode; 3=tzf code)
    lCopyCode As Long        'Copy inventory code
    iLen As Integer         'Spot length
    sPriceType As String * 1 'L=Use line Price; P=Post Log N/C old values => T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
                            'For SpotType = X; PriceType = "B"
    sSpotType As String * 1 'Spot type: A=Contract spot; O=Open BB; C=Close BB; L=Cmml Promo;
                            'N=Feed; S=PSA; M=Promo; D=Donut; B=Bookend; Y=BB only; E=Any BB;
                            'T=Remnant (if spfSchdRemnant = "N"); Q=per Inquire
                            'X=Extra Spot
    sBill As String * 1     'Billed status (Y= Yes; N=No)
    iRotNo As Integer       'Rotation #. Note -1 indicates ?
    lSmfCode As Long        'SMF Code Number
    iUrfCode As Integer     'Last user who modified spot
    lFsfCode As Long
    iGameNo As Integer      'Game number
    sXCrossMidnight       As String * 1      ' Is spot time across midnight (Y/N
                                             ' 0r Blank).  If so (Y), then air
                                             ' date is sdfDate plus 1
    sWasMG                As String * 1      ' Missed spot was once a MG (Y/N or
                                             ' Blank).  Previous MG information
                                             ' stored into sxf_Spot_Extension.
                                             ' Test for Y
    sFromWorkArea         As String * 1      ' Missed spot was left in the work
                                             ' area when exited the Spot screen
                                             ' (Y/N or blank).  Information
                                             ' retained in sxf_Spot_Extension so
                                             ' that the work area can be
                                             ' rebuilt. Test for Y
    sUnused               As String * 8      ' Unused
End Type

'Sdf key record layout
Type SDFKEY0
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    lFsfCode As Long
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready?, U=Unscheduled,..
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

'Sdf key record layout
Type SDFKEY1
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready?, U=Unscheduled,..
End Type

Type SDFKEY2
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready?, U=Unscheduled,..
    iAdfCode As Integer     'Advertiser code number
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

'Sdf key record layout- use LONGKEY0
'Type SDFKEY3
'    lCode As Long
'End Type

Type SDFKEY4
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    lChfCode As Long        'Contract code
End Type

'Sdf key record layout- use LONGKEY0
'Type SDFKEY5
'    lChfCode As Long
'End Type

Type SDFKEY6
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    iGameNo As Integer      'Game number
End Type

Type SDFKEY7
    iAdfCode              As Integer
    iDate(0 To 1)         As Integer
End Type

Type SDFMDEXT
    sKey As String * 50
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    lFsfCode As Long        'Feed code
    iAdfCode As Integer     'Advertiser code number
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iMnfMissed As Integer   'Missed reason
    iLen As Integer         'Spot length
    lCntrNo As Long         'Contract number
    sCntrType As String * 1 'Contract type
    sAdvtAbbr As String * 7 'Advertiser abbr name
    lRecPos As Long         'Record position
    sShow As String * 120
    sPrice As String * 12
    sAdvtProd As String * 30    'Advt Abbr and Product
    iRotNo As Integer
    sPtType As String * 1
    lCopyCode As Long
    lchfcxfCode As Long     'Other Header comment if showing on spot screen (-1=Not set)
    lchfcxfInt As Long      'Internal comment
    lClfCxfCode As Long     'Line comment if showing on Spot screen (-1=Not Set)
    iLineInfo As Integer    'See EVTINFO
    iNoMissed As Integer    'Number of missed (for Contract display)
    iNoOrdered As Integer   'Number of Ordered
    iMissedMG As Integer    'Missed MG (True or False)
    iGameNo As Integer
    sSplitNetwork As String 'Split Network buy: "N"=No; "Y"=Yes
    sLiveCopy As String * 1
    sNetRegionAbbr As String * 5
    sCITFlag As String * 1  'C=Cart; I=ISCI; T=creative Title
    sCopyCIT As String * 30    'Cart # or ISCI or Creative title.  sCITFlag indicates which value
    iSelected As Integer    '0=Not selected; 1=Selected. Used when moving between work area and missed area
    lSxfCode As Long
    lSdfCode As Long
End Type

'Build within gObtainCntrSpots (PopSubs.Bas) and mObtainCntrToInvoice (Invoice.Frm)
Type SDFEXT
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
    sTracer As String * 1   'M=Mouse move, O=On demand & mouse moved, C=Created in post log,
    iLen As Integer         'Spot length
    sPriceType As String * 1 'L=Line; P=Post Log N/C  was T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    sSpotType As String * 1 'Spot type: A=Contract spot; O=Open BB; C=Close BB; L=Cmml Promo;
                            'N=Feed; S=PSA; M=Promo; D=Donut; B=Bookend; Y=BB only; E=Any BB;
                            'T=Remnant (if spfSchdRemnant = "N"); Q=per Inquire
    sBill As String * 1     'Billed status (Y= Yes; N=No)
    iAdfCode As Integer     'Advertiser code
    iGameNo As Integer      'Game Number
    sXCrossMidnight       As String * 1      ' Is spot time across midnight (Y/N
                                             ' 0r Blank).  If so (Y), then air
                                             ' date is sdfDate plus 1
    lCode As Long           'Sdf Code
    lRecPos As Long         'Record position
    lMdDate    As Long     'Missed date for MG's (used in gCntrDisp)
    iStatus As Integer     'Status bits:0=Date error if 1; 1=Time error if 1
                           'Overlayed with Original vehicle if call to gObtainCntrSpot filed ilAllSchdInSelVef is set to True
                           'It is only set to True in Spots when showing Graphs
End Type
Public Const SDFEXTPK As String = "ILIIIIIBBIBBBIIBL" 'LLI"
Public Const SDFEXTPK_RPT As String = "ILIIIIIBBIBBBIIBLLL" 'LI"  7-15-02 used for Advt spot count report.  LRecPos is overlayed with the smfcode,
                                                        'and iLen is overlayed with the advt code
                                                        '8-6-04 fsfcode is overlayed with Missed date (long)
Type SDFEXTSORT
    sKey As String * 70 'Line #|Line Vehicle|Date|Time
    'iSdfExtIndex As Integer    'Replaced with long to handle more then 32000 records
    lSdfExtIndex As Long
End Type


'********************************************************
'
'Short Title file definition
'
'*********************************************************

'Sif record layout
Type SIF
    lCode As Long           'Autoincrement number
    iAdfCode As Integer     'Advertiser code number
    sName As String * 15    'Short Title
    sState As String * 1        'A=Active; D=Dormant
    iUrfCode As Integer     'Last user who modified product names
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    lAutoCode As Long       'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    sUnused As String * 10
End Type

'Sif key record layout- use LONGKEY0
'Type SIFKEY0
'    lCode As Long
'End Type

'Sif key record layout
Type SIFKEY1
    iAdfCode As Integer     'Advertiser code number
End Type


'********************************************************
'
'Salesperson file definition
'
'*********************************************************

'Slf record layout
Type SLF
    iCode As Integer        'Internal code number for salesperson
    sFirstName As String * 20   'Identification name
    sLastName As String * 20    'Last name
    iSofCode As Integer     'Sales office code
    sPhone As String * 14   'Phone number/extension
    sFax As String * 10     'Fax number
    iMnfSlsTeam As Integer  'Sales team code
    'sSS As String * 9       'Social security
    'sGLAcct As String * 10  'G/L Account #
    'lDrawAmt As Long  'Draw amount xxxxx.xx
    'sCYBud(1 To 12) As String * 5   'Current year budget (xxxxxxxxx)
    'sLYBud(1 To 12) As String * 5   'Last year budget (xxxxxxxxx)
    'sCYAct(1 To 12) As String * 5   'Current year actuals (xxxxxxx.xx)
    'sLYAct(1 To 12) As String * 5   'Last years actual (xxxxxxx.xx)
    'sCYComm(1 To 12) As String * 5  'Current year commission earned (xxxxxxx.xx)
    'sLYComm(1 To 12) As String * 5   'Last years earned commission (xxxxxxx.xx)
    iMerge As Integer       'Merge code number
    sJobTitle As String * 1     'S=Salesperson; P=Planner; M=Sales Manager; N=Negotiator
    sState As String * 1        'A=Active; D=Dormant
    iUrfCode As Integer     'Last user code number who altered this file
    sCodeStn As String * 5  'Station salesperson code
    lSalesGoal As Long      'Sales Goal (xxxxxxxx)
    iUnderComm As Integer      'Commission paid if under sales Goal (xx.xx)
    iOverComm As Integer       'Commission paid if over sales goal (xx.xx)
    iRemUnderComm As Integer      'Commission paid if under sales Goal (xx.xx)
    iRemOverComm As Integer       'Commission paid if over sales goal (xx.xx)
    lStartCommPaid As Long  'Start Commission Paid (2 places)
    lStartSales As Long     'Start Sales (Gross or Net, 2 places)
    iStartCommDate(0 To 1) As Integer   'Start commission Date
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    'iGroupNo As Integer             'Group number used if salesperson does not have a user defined
    'sUnused As String * 18
    iNewClientComm As Integer   'New Client business Commission (xx.xx)
    iIncClientComm As Integer    'Increase Client business commission (xx.xx)
    sUnused As String * 16          '12-11-06 remove groupno, and give back to unused
End Type

'Slf key record layout- use INTKEY0
'Type SLFKEY0
'    iCode As Integer
'End Type

'Type SLFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

'Salesperson extract record
Type SLFEXT
    iCode As Integer        'Internal code number for Vehicle
    sFirstName As String * 20   'Identification name
    sLastName As String * 20    'Last name
    sState As String * 1        'A=Active; D=Dormant
End Type
Public Const SLFEXTPK As String = "IB20B20B"

'********************************************************
'
'Spot Makegood Specification file definition
'
'*********************************************************

'Smf record layout
Type SMF
    lCode As Long           'Autoincrement Code
    lSdfCode As Long        'SDF Code number
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    iMissedDate(0 To 1) As Integer    'Missed Date of spot
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iMissedTime(0 To 1) As Integer 'Missed Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iStartDate(0 To 1) As Integer    'Start Date of spot
    iEndDate(0 To 1) As Integer    'End Date of spot
    iActualDate(0 To 1) As Integer 'Actual Date of spot scheduled
    iStartTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iEndTime(0 To 1) As Integer 'End Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iActualTime(0 To 1) As Integer 'Actual Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    'sWkDays(1 To 7)  As String * 1  'Weekday flag: Y=day allowed; N=Day disallowed
    sWkDays(0 To 6)  As String * 1  'Weekday flag: Y=day allowed; N=Day disallowed
    sMGSource As String * 1     'S=Created by MG Spec, M=Created by mouse move outside contract
                                'A=Mouse Move to another vehicle
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready to sch MG, U=Unscheduled MG,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled, H=Hidden, O=Outside
    iUrfCode As Integer
    iOrigSchVef As Integer      'Original vehicle spot booked into
    sPtType As String * 1   'Only used for Bill As Order to retain Missed Copy Assignment
                            'Copy code type (0=none assigned; 1=cif code; 2= ccfcode; 3=tzf code)
    lCopyCode As Long        'Copy inventory code
    iRotNo As Integer       'Rotation #. Note -1 indicates ?
    lMtfCode As Long    'MG Tracking File code (when M for N MG genertaed, otherwise zero)
    lFsfCode As Long
    iGameNo As Integer      'Original game number
    sUnused As String * 10
End Type

'Smf key record layout
Type SMFKEY0
    lChfCode As Long        'Contract code
    iLineNo As Integer      'Line number
    lFsfCode As Long
    iMissedDate(0 To 1) As Integer    'Schedule or missed Date of spot
End Type

'Smf key record layout- use LONGKEY0
'Type SMFKEY1
'    lCode As Long
'End Type

'Smf key record layout- use LONGKEY0
'Type SMFKEY2
'    lSdfCode As Long
'End Type

Type SMFKEY3
    iOrigSchVef As Integer      'Original vehicle spot booked into
    iGameNo As Integer      'Original game number
End Type

Type SMFKEY4
    iMissedDate(0 To 1) As Integer    'Missed Date of spot
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    lChfCode As Long        'Contract code
End Type

Type SMFKEY5                       '4-5-10
    iOrigSchVef As Integer      'Original vehicle spot booked into
    iMissedDate(0 To 1) As Integer    'Missed Date of spot
End Type

'********************************************************
'
'Sales Office File Definition
'
'********************************************************

'Sales Office record layout
Type SOF
    iCode As Integer        'Internal code number for Sales office
    iMnfSSCode As Integer   'Reference to multi-name: Sales source
    iMnfRegion As Integer   'Reference to multi-name: Sales regions
    sName As String * 20    'Name
    iMktRank As Integer     'Market Rank
    sAddr(0 To 2) As String * 25   'Address
    sPhone As String * 14   'Phone plus extension
    sFax As String * 10     'Fax number
    'iMerge As Integer       'Merge code number
    sState As String * 1    'A=Active; D=Dormant
    sUnused1 As String * 20  'Unused
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode
End Type

'Sales Office key record layout- use INTKEY0
'Type SOFKEY0
'    iCode As Integer
'End Type

'Sales office and sales source code extract record
Type SOFMNFEXT
    iCode As Integer
    iMnfSSCode As Integer
    sSOFName As String * 20
End Type
Public Const SOFMNFEXTPK As String = "IIB20"

'********************************************************
'
'Site Option file definition
'
'*********************************************************

'Spf record layout
Type SPF
    iCode As Integer    'Internal code number for site Option
    iUrfGCode As Integer    'Last user to change the site Option file
    sGClient As String * 25 'Client name
    sGAddr(0 To 2) As String * 25   'Client address
    sGTBar As String * 1  'C=show client name in title bar; U=show user name in title bar
    sSSellNet As String * 1  'Y=using selling to airing links; N= not using links
    sSDelNet As String * 1  'Y=using selling to delivery links; N= not using links
    'iSDay As Integer        'Unused'# of days to retain Hold & reseravtion spots
    'replaced iSDay with following bytes 8-23-01
    sInvVehSel As String * 1    'Y=Allow selective vehicles in invoice
    sSchdPromo As String * 1       'Y=Schd Promo; N=Manual Schd
    sSBrStdQt As String * 1  'Contracts/Propsals by Standard Quarters(Y/N)
    sHideGhostSptScr As String * 1  'Hide Ghost Spots on Spot Screen (Y/N) was Using Reservation Contracts Types(Y/N)
    sBActDayCompl As String * 1  'Active 'Day Complete' Testing Types(Y/N)
    'sSUseRem As String * 1  'Unused-'Using Remnant Contracts Types(Y/N)
    'sSUseDR As String * 1  'Unused-'Using Direct Response Contracts (Y/N)
    'sSUsePI As String * 1  'Unused-'Using per Inquire Contracts Types(Y/N)
    'sSUsePSA As String * 1  'Unused-'Using PSA Contracts Types(Y/N)
    lBCxfDisclaimer As Long    'Invoice disclaimer
    'sSUsePromo As String * 1  'Unused-'Using Promo Contracts Types(Y/N)
    sBMissedDT As String * 1    'A=Use MG/Outside Date/Time for Missed Date/Time when Invoicing by Order; R=Random
    iSLen(0 To 9) As Integer 'Default spot lengths
    iPkageGenMeth As Integer       'Ordered Package Generate $ Method: 0=Week; 1=Line
    sAExcl As String * 1    'Y=Using program exclusuions; N=Not using program exclusions
    iATargets As Integer    'Number of target demos (max 4)
    sAISCI As String * 1    'Show ISCI code on invoices (A=Always, N=Never, Y=Varies with Yes default
                            'N=Varies with No as default)
    sAEDIC As String * 1    'Y=Using EDI Service for contract, N=Not using service
    sAEDII As String * 1    'Y=Using EDI Service for invoices, N=Not using service
    sAPrtStyle As String * 1    'Contract form (W=Wide, N=Narrow, A=Ask by client)
    sARepCodes As String * 1  'Y=Using Rep codes, N=Not using Rep codes
    sAStnCodes As String * 1  'Y=Using station codes, N= Not using Station codes
    sAAgyCodes As String * 1  'Y=Using Agency codes, N=Not using agency codes
    sCBumpPast As String * 1  'Y=Bump spots in past; N=Don't
    'sLkLog As String * 1    'Lock for Log field
    sSportInfo As String * 1 'Bit Map (Right to Left) : 0=Using Sports; 1=Sports Preempting Regular Programming; 2=Using Feed;
                             '3=Using Language;
    lCLowestNo As Long      'Contract lowest start number (0=not using auto assign)
    lCHighestNo As Long     'Contract highest end number (0=not using auto assign numbers)
    lCNextNo As Long        'Next number to assign to a contract
    sCEstNo As String * 1   'Y=Using Agency estimate number; N=Not using estimate number
    'sCRot As String * 1     'Unused-'Copy rotation (I=all Independent, M=all Mixed, V=Varies with instruction)
    sCBlackoutLog As String * 1 'Using Blackouts on Logs (Y/N) - this replaced sCRot
    sBLCycle As String * 1  'Default Local billing cycle (C=Calendar, S=Standard, W=Weekly)
    sBRCycle As String * 1  'Default Regional billing cycle (C=Caldendar, S=Standard, W=Weekly)
    sBNCycle As String * 1  'Default National billing cycle (C=Caldendar, S=Standard, W=Weekly)
    lBLowestNo As Long      'Invoice lowest start number
    lBHighestNo As Long     'Invoice highest end number
    lBNextNo As Long        'Next number to assign to a invoice
    sBCombine As String * 1 'Y=combine same advertiser, different Vehicles on same invoice; N= Don't combine
    'sBSepItem As String * 1 'Unused'Y= put Item bills on separate invoices; N=combine item bill with spot bill
    sCDefLogCopy As String * 1 'Y= Check Log Assign Copy On; N=leave Log Assign Copy unchecked
    'iBTax(0 To 1) As Integer    'Tax amount (xxx.xx)
    sUsingFeatures5 As String * 1    'Bit Map (Right to Left):  Bit 0=Remote Export; Bit 1=Remote Import; Bit 2=Combine Air and NTR on same invoice
                                     '3-4 Sort Order and Invoice by (3=Rate Card; 4= Line #, both zero = Daypart); 5=Intercae; 6= RADAR; 7=Suppress time on invoice form 1
    sUsingFeatures6 As String * 1    'Bit Map(Right to Left): Bit0=Installment Billing; Bit 1= Installment Contract Revenue by Earned (field value 1) or Billed(Field value 0)
                                     'Bit 2= BB on Separate line(0=Yes;1=No); Bits 3=BB Type(0=Open/Close; 1=Closest BB)
                                     'Bit 4=GetPaid Export(0=No; 1=Yes); Bit 5=Digital Content(0=No; 1=Yes); Bit 6=Guaranteed By Gross Impressions (GrImp); Bit 7=Invoice Export
    sAutoType3 As String * 1        'Bit Map (Right to Left): Bit 0= Audio Vault Aired; 1= Wide Order; 2= Jelli; 3=ENCO-ESPN; 4 = Scott V5, 5 = Zetta; 6=Station Playlist
    sInvSpotTimeZone As String * 1  'Invoice spot time conversion zone. E; C; M; P or N (or Blank). Test for E; C; M or P.
    'iRCorp(0 To 2) As Integer   'Corporate calendar(Index 0=first period, Index 1=Next period, 2=Last period of cycle)
    '                            'Values are 0 if not used, 4 or 5 otherwise
    'sRYEnd As String * 1  'L=Corp calendar using Last sunday of dec.; D=Using last Day in dec.
    sInvExportId As String * 2  'Invoice export ID
    sSchdPSA As String * 1       'Y=Schd PSA; N=Manual Schd
    iRLastPurgedDate(0 To 1) As Integer    'Last purged date
    iRLastPay(0 To 1) As Integer    'Last payment date
    sRCurrAmt As String * 1 'Y=Include current amount in A/R when computing credit limit;N=don't include current A/R
    sRUnbilled As String * 1    'Include unbilled amount when computing credit limit; N=Don't include unbilled
    'sLkCredit As String * 1     'Lock for set credit
    sMOFCopyAssign As String * 1 'Bit Map (right to Left):  Bits 0-2 MG or Outside Copy Assigment default: 0=Orig Veh only; 1=Sch Veh only; 2=Sch or Orig Veh;
                                 'Bits 3-5 Fill Copy Assigment: 3=Orig Veh only; 4=Sch Veh only; 5=Sch or Sold Veh
                                 'Bit 6=MG Rules specified in Copy; Bit 7=Single Custom Demo
    iRPctCredit As Integer     '% credit limit to get notification (dec xxx)
    sRRP As String * 1          'Reconciling Period S=Standard; C=Calendar; F=Fiscal month
    iRPRP(0 To 1) As Integer    'End of Previous Reconciling Period
    iRCRP(0 To 1) As Integer    'End of Current Reconciling Period
    iRNRP(0 To 1) As Integer    'End of Next Reconciling period
    sRB As String * 6           'Receivables Balance, End of Previous Reconciling Period (xx,xxx,xxx.xx)
    lRPenny As Long            'Penny variance (xxx.xx)
    iBLastStdMnth(0 To 1) As Integer    'End of last standard month billed
    iBLastCalMnth(0 To 1) As Integer    'End of last calendar month billed
    sBPayName As String * 25 'Payment name
    sBPayAddr(0 To 2) As String * 25   'Payment address
    sRCollectPhoneNo As String * 14  'Collection phone number
    sRCollectContact As String * 20     'Collection contact name
    iRNoWks As Integer      'Number of weeks to compute contract dollars for credit limits
    iRCreditDate(0 To 1) As Integer    'Last Date that Advt/Agy credit computed
    'iUnused2 As Integer 'Number of months after condensed to retain contracts
    sUsingFeatures3 As String * 1  'Bit map (right to left): Bit 0 Using Hub; Bit 1 Tax on NTR; Bit 2 Using Promo Copy; Bit 3 Media Code assigned to vehicle;
                                   'Bit 4 include media code with audio vault; Bit 5 Book PSA into Contract Avails; Bit 6 Book Promo into Contract Avails; Bit 7:  Tax on Air time
    sUsingFeatures4 As String * 1  'Bit Map: 0=Lock box by Vehicle; 1=Bypass Demographic Fields; 2=Bypass Date Needed in Proposals/Orders; 3=Allow Spot Moves on Todays date
                                    'bit 4:  allow cnt price to be changed if billed (12-11-06); 5-6 Tax by USA (bit 5) or Tax by Canada (bit 6); Bit 7: Invoice sort by Vehicle (1), otherwise by Payee (0)
    iRetainAffSpot As Integer  'Number of months after Last Invoice date to retain affiliate spots and agreements
    iRetainTrafSpot As Integer 'Number of months after last invoice date to retain traffic spot and copy
    iMnfClientAbbr As Integer    'Client Abbreviation (stored with mnfType = "J" and MnfUntiType = "A")  'Number of months to retain condensed spot history after full spot history
    iMnfInvTerms As Integer  'Default terms (stored with mnfType = "J" and mnfUnitType = "D") 'Number of months to retain payment history after Last invoice date
    sAllowMGs As String * 1 'Allow spot moves outside of contract parameters (Y/N)
    sVirtPkgCompute As String * 1   'For virtual packages, when package spots changed recompute (H=Hidden number of spots or P=Package Line price)
    sGUseAffFeed As String * 1  'Using Affiliate Feed (Y/N)
    sCLnStdQt As String * 1 ' NO LONGER USED todo: when code cleaned up, mark as unused (was:Show lines by Standard Quarter(Y) or from line start date(N))
    sSSRetainDate As String * 1  'Retain Spot Screen Date when vehicle switched(Y/N)
    sMktBase As String * 1
    sPLMove As String * 1   'Post Log Move (M=MG's; O=Outsides; A=Ask)
    sAutoType As String * 1 '8-10-05 chg from unused, Automation type (bits left to right): 1=Dalet; 2=Prophet NexGen; 3=Scott; 4=Drake; 5=RCS; 6=Prophet Wizard, 7=Media Star, 8 = iMedia Touch
                            'there are 2 fields used for automation type:  see also AutoType2
    sCUseSegments As String * 1  'Using Segments(Y/N)
    iPassword As Integer    'Password for Site Option, ShowFix
    iVehLen As Integer  'Vehicle length (max 40)
    sEDICallLetter As String * 4    'Call Letters to EDI record 22- if not defined values from Vehicle (not code)
    sEDIMediaType As String * 2     'Media Type (TV,R, LC...)
    sEDIBand As String * 2          'AM or FM,..
    sUseCartNo As String * 1    'Y=Using Cart #'s, N=Use ISCI Numbers in place of cart numbers (test for N not Y)
    sExport As String * 1   'Changed definition 3/20/03 # spaces on invoice from top to Logo (0= vbDefault = standard)'Y=Show export menu item; N=Don't show
    sImport As String * 1   'Changed definition 3/20/03:  # of spaces from Logo to Address on invoice form (0 = vbDefault = Standard)'Y=Show import menu item; N=Don't show
    lDiscCurrCntrNo As Long     'Spot Discrepancy report current contract number being processed
                                'only set when running all contracts
    iDiscDateRun(0 To 1) As Integer    'Date discrepance report last run
    sUseProdSptScr As String * 1    'P=Use Short Title; A=Use Advertiser/Product on spot screen
    iGAlertInterval As Integer  'Time interval between checking for alerts(0=don't)
    sBSlspComm As String * 1    'B=on Billing Amount; C=on Collection Amount
    'Replaced with Feature10
    'sXSDAddAdvtToISCI As String * 1 'Y=Add Advt/Product to ISCI with X-Digital Spot Insertion
    '                                'Field was sBSlspBack: Y=Take back commission for write-offs; N=Ignore wirte offs
    sUsingFeatures10 As String * 1   'Bit Map(Right to Left):0=X-Digital: Add Advertiser to ISCI; 1=X-Digital: Midnigh Based Hours; 2=Package Line Rate on BR; 3=Using Wegener iPump; 4=Replace Deleted Weeks with Fills; 5=vcreative export
                                     'Bit 6 = Contract verification; Bit 7 = XDS Units by astCode
    iBarterLPD(0 To 1) As Integer         ' Barter Last Paid date
    iRetainTrafProp As Integer     'Number of months after todays date to retain Dead and Expired Proposals
    iGNoDaysPass As Integer 'Number of days to retain same password (0=no limit)
    sCPkOrdered As String * 1   'Allow Virtual(Ordered/Generated) Packages (Y/N)
    sCPkAired As String * 1     'Allow Real (Aired) Packages (Y/N)
    iRptDollarMag As Integer    'Report Dollar Magnitude (1, 100, 1000)
    sDefFillInv As String * 1   'Set Show Fill Spots on Invoice as Y=Yes; N=No
    sInvAirOrder As String * 1   'Invoice All Contracts as A=Show Aired, Update Aired; O=Show Ordered, Update Aired;
                                 'S=Show Order, Update Order; 2=Show Aired Minus Missed and MG as Aired with MG info, Update Ordered
                                'Note: if as Aired, Package lines can still be defined as Aired or Odered
                                '      sCPkAired (Allow Aired Package) and sCPkOrdered (Allow Ordered Packages)
    sBOrderDPShow As String * 1 'N=Name; T=Time; B=None
    sRUseTrade As String * 1    'Using Trade Receivables (Y/N)
    sRUseMerch As String * 1    'Using Merchandising Receivables (Y/N)
    sRUsePromo As String * 1    'Using Promotion Receivables (Y/N)
    sSEnterAgeDate As String * 1 'for billed & booked - use E = Entered Date, A = AGed DATe
    iNROBlockDate(0 To 1) As Integer    'Next rollover block date
    iNROBlockTime(0 To 1) As Integer 'Next Rollover Block Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sRUseCorpCal As String * 1  'Using Corporate Calendar (Y/N)
    sCIncludeMissDB As String * 1   'Y=Include missed with demo bar; N= Don't
    sRemoteUsers As String * 1  'Y=Remote Users Allowed; N = No remote Users Allowed
    sGUsePropSys As String * 1  'Y=Using Proposal System; N=Not Using Proposal System
    sSAudData As String * 1     'Audience Data Magnitude: T=(000); H=(00); N((0); U()
                                'Audience and Population always shown in (000): therefore H shown as xx.x; N as xx.xx and U as xx.xxx
    sSGRPCPPCal As String * 1    'GRP/CPP calculations: R=Rating; A=Audience
                                'Rating then GRP = # Spots * Rating; CPP = $/GRP
                                'GRP 2 places then GRP = # Spots * Audience/Population; CPP = $ / GRP
                                'Audience then GRP = # Spots * Audience/Population; CPP = $ * Population/GrImp
    sGUseAffSys As String * 1  'Y=Using Proposal System; N=Not Using Proposal System
    sBLaserForm As String * 1  '1=Air,Order,Recon Form; 2= Aired only form
    sCWarnMsg As String * 1     'Contract Warning Messages (Y/N)
    sCPkEqual As String * 1     'Allow Equal Packages (Y/N)- Package dollars distributed to Hidden lines
    sCAudPkg As String * 1      'Allow Audience Packages (Y/N)- Package dollars
                                'Distributed to Hidden lines via Audience if selected
                                'Automatically for Dynamic Package- User select this option for Standard Packages
    iReallMnfDemo As Integer    'Audience Reallocation Mnf Demo copy number
    iReallDate(0 To 1) As Integer   'Audience Reallocation Date
    'iLkLogUrfCode As Integer    'User that Locked Log last.
    iReconcGroupNo As Integer   '1-20-04 replaced ilklogurfcode with this field: indicates default vehicle group to major sort
                                'the recnciliation reports (in addition to the current ones)
    sAllowPrelLog As String * 1 'Allow Preliminary Logs (Y/N)
    sSchdRemnant As String * 1  'Schedule Remnant Contracts (Y/N)
    sSubCompany As String * 1   'Define Sub-Company vehicle groups
    sCommByCntr As String * 1   'Allow salesperson commission to be defined by Contract(Y/N)
    sPostCalAff As String * 1   'Post calendar affidavits (for rep posting)(S=Standar;C=Calendar;W=Week;N=None). D=Daily Spots moved to Features8 6/7/10
    iRepPrintDate(0 To 1) As Integer 'Last date Rep Printed  9/10/02
    sTax1Text As String * 25    'Tax 1 Text
    sTax2Text As String * 25    'Tax 2 Text
    iNoMnthNewBus As Integer         ' # months off before considered New Business
    iNoMnthNewIsNew As Integer         ' # months New Business remains New
    sNewBusYearType As String * 1      ' Base New Business on Calendar or Rolling year:  C = calendar, R = rolling
    iLastWeekBilled(0 To 1) As Integer       ' Last Week billed if using Weekly billing
    sAvailEqualize As String * 1    'Combo Avail Report Equalize by ( ) 30 ( ) 60 ( ) None.  3=30; 6=60; N=None. Test for 3 or 6
    sDemoEstAllowed As String * 1   'Demo Estimates Allowed (Y/N)
    lCxfDemoEst As Long         'Demo Estimate Comment to be placed on contract print-outs
    sCmmlSchStatus As String * 1    'For Reprint Log- Retain Commercial Schedule R=Retain; N=Never; A=Ask
    sSystemType As String * 1       'System Type: N=Network/Syndicator; R=Radio Station
    sUsingTraffic As String * 1     'Y or Blank = Yes ; N=No (only allow rate card; budget and Proposals)
    iPriUrfCode As Integer      'Primary Web Poll User
    iSecUrfCode As Integer      'Secondary Web Poll user
    sImptCntr As String * 1     'When Importing Orders: Y=Use Separation Rules witBite Map: Show standard package hidden lines: Y=Yes; N=No
    sUsingFeatures2 As String * 1   'Bit Map: Bit 1 - Regional; 2 - Split Copy; 3 - Split Network; 4 - Barter; 5 - Strong Password; 6 - Allow Mixture of Air Time and Rep on same contract; 7-Merch/Promo by $; 8=Great Plain G/L
    sUsingBBs As String * 1     'Only allowed if not using Station Feed.  To allow it to work with Station Feed, we must determine
                                'how it is to be coded in Station Feed, CnC stuff, Export NY, Phnx and Dallas.
    sBBsToAff As String * 1     'Send BB's to Affiliate system (Y/N).  Currently Yes is not allowed.  Might want to change the Y/N to ask three questions:
                                '( ) Send as Separate Spots ( ) Send as Part of Sold Spot ( ) Don't Send
                                'If sent as separate spots, then we need to determine how the Affiliate pledge should be
                                'defined.  If Sold sent, then send BB's or require BB avails to be set up separately in the pledge.
    sTapeShowForm As String * 1 'On the Copy Inventory screen show: A=Tape Approved; C=Tape Carted
    'sOptionFields As String * 1 'Bit Map (Right to Left): 0=Projection; 1=Bus Cat; 2=Share; 3=Rev Set; 4=Guar; 5=Billing Cycle; 6=Co-op; 7=Research
    sOptionFields As String * 1 'Bit Map (Right to Left): 0=Projection; 1=Bus Cat; 2=Share; 3=Rev Set; 4=Guar; 5=Calendar Bill; 6=Co-op; 7=Research
    sOverrideOptions As String * 1 'Bit Map (Right to Left): 0=Allocation; 1=NTR Acquisition; 2=1st Position; 3=Preferred Days/Times; 4=Solo Avail; 5=Ignore Best Fit (see sFeature for other options); 6=BBs Same Line; 7=Bypass HiPri
    sAutoType2 As String * 1    '8-10-05 continuation of more automation types (see tgspf.sAutoType for 1st set of types)
                                'bits from right to left:  bit 0 = Audio Vault Sat, Bit  1 = Audio Vault Air (unused for now); 2 = Wire Ready, 3 = Enco, 4=N-R Rep, 5 = N-R Net, 6= Simian, 7=RCS 5
    sRepRptForm As String * 1   'Rep report 2 form by M (Market) or V (Vehicle)
    sAllowDailyBuys As String * 1   'Allow Daily Buys (Y/N)
    sUsingFeatures7 As String * 1   'Bit Map(Right to Left):0=Using CSI Backup; 1=Bonus Commission (0=Yes; 1=No); 2=Commssion Year(0=Std Broadcast; 1=Fiscal); 3 = Export Revenue 10-2-08; 4=Allow Mixed Region Length
                                    '5=X-Digital; 6=Wegener; 7=OLA
    sUsingFeatures8 As String * 1   'Bit Map(Righyt to Left):0=Live/Record mandatory in contract; 1=Show Contract Comments on Detail page; 2=Allow Metro Market with Split Copy; 3=Rivendell Automation Export; 4=X-Digital by Break; 5=ISCI Export; 6=Pre-Feed; 7=Rep by Date & Time
    sUsingFeatures9 As String * 1   'Bit Map(Right to Left):0=Affiliate CRM; 1=Default Pre-recorded 1st Position; 2=Proposal XML; 3=Limit ISCI to 15 characters; 4=Enforce IDC Copy Restrictions; 5=Weekly bill; 6= Print EDI; 7 = Word Wrap Vehicle
    sWegenerGroupChar As String * 1 'Wegener Extension group start character.  If blank, use W
    iXDClientID           As Integer         ' X-Digital Prefix for Transmission ID and Unit ID
    sUsingNTR As String * 1     'Using NTR (Y/N), Default N
    sUsingFeatures As String * 1 'Bit Map (Right to Left) : 0=Using Matrix Export; 1=Using Revenue Export; 2=Live Copy; 3=MultiMedia; 4=Live Log; 5=Rep; 6=Ignore Hidden overrides with Research; 7=Weight Time(0) or Days(1) with Best fit
                                 '(0=No 1=Yes)
    sInsertAddr As String * 1   'Insertion Order Address (Payee-P, Site Invoice-S, Vehicle-V)
    lCxfContrComment As Long     'Contract Comment
    lCxfInsertComment As Long   'Insertion Order comment
    sUnitOr3060 As String * 1   'U=Units; 3=3060
End Type

'Spf key record layout- use INTKEY0
'Type SPFKEY0
'    iCode As Integer
'End Type

'Buy 6-10; Mo-We; Currently 3 spots schedule
'  Hour   Count      Day   Count        Q-H   Count
'   6       2        Mo      1          1st     2
'   7       0        Tu      0          2nd     0
'   8       0        We      2          3rd     1
'   9       1                           4th     0
'
'Cycle pattern 1
'Hour 8 8 7 7 8 7 8 7 8 8 7 7 8 7 8 7 8 8 7 7 8 7 8 7 -- 6
'Day  T T T T T T T T M M M M M M M M W W W W W W W W    W
'Q-H  4 2 4 2 3 3 1 1 4 2 4 2 3 3 1 1 4 2 4 2 3 3 1 1    1
'     A A A A A A A A A A A A A A A A A A A A A A A A    A
'           B   B   B       B   B   B       B   B   B    B
'                   C               C               C    C
'                                                   D    D
'                                                        E
'
'Cycle pattern 2
'Hour 8 8 8 8 8 8 8 8 8 8 8 8 7 7 7 7 -- 6
'Day  T T T T M M M M W W W W T T T T    W
'Q-H  4 2 3 1 4 2 3 1 4 2 3 1 4 2 3 1    1
'     A A A A A A A A A A A A A A A A    A
'       B B B   B B B   B B B   B B B    B
'           C       C       C       C    C
'                           D            C
'                                        E

'10209

'******************************************************************************
' SPFX_Site_Options Record Definition
'
'******************************************************************************
Type SPFX
    iCode                 As Integer         ' Primary Key 1-to-1 match with
                                             ' SPF_Site_Options value of site
                                             ' table
    sInvExpProperty       As String * 12     ' Invoice Export Property name
    sInvExpPrefix         As String * 6      ' Invoice Export Prefix
    sInvExpBillGroup      As String * 12     ' Invoice Export Billing Group name
    iInvExpFeature        As Integer         ' Invoice Export features: 0 =
                                             ' Disabled; 1 = Audacy WO; 2 =
                                             ' Audacy Line Invoice Export; 4 =
                                             ' Selective PDF Email
    sSageTerm             As String * 40     ' Sage Intacct Export Term Name
    sSageAccount          As String * 24     ' Sage Intacct Export Acct Number
    sSageLocation         As String * 20     ' Sage Intacct Export Loccation ID
    sSageDept             As String * 20     ' Sage Intacct Export Department ID
    iLineCostType         As Integer         ' Digital flat rate line cost
                                             ' 0=monthly; 1=daily
    iIntFeature           As Integer         ' 0 = Disabled; 1=Honor DST
    iSchdFeature          As Integer         ' 0=Disabled,1=Disallow user from scheduling own contract 2 and on for future features
    sUnused               As String * 94
End Type


Type SPFXKEY0
    iCode                 As Integer
End Type

'********************************************************
'
'Spot summary File Definition
'
'********************************************************


'PROGRAM, AVAIL and SPOT must match in size
Type PROGRAMSS
    iRecType As Integer         '(Bits 15-0 Left to right): Bit 0-3 = record type (program) = 1
    iLtfCode As Integer         'Used for testing if library changed
    iStartTime(0 To 1) As Integer   'Event Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iEndTime(0 To 1) As Integer     'Event End Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iMnfExcl(0 To 1) As Integer 'Program Exclusion code
    lLvfCode As Long            'Used for testing if library changed
End Type

Type AVAILSS
    iRecType As Integer         '(Bits 15-0 Left to right)Bit 0-3 = record type (avail) = 2-9
                                'Value matches Event type #
    iNoSpotsThis As Integer     'Number of spots associated with this avail;
    iTime(0 To 1) As Integer    'Event Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iLtfCode As Integer            'Used for library buys
    iAvInfo As Integer          'Avail information (bits 15-0 Left to right):
                                '   Bits 0-4 for units;
                                '   Bit 6 for Avail Locked;
                                '   Bit 7 for Spot Locked;
                                '   Bit 8 for sustaining allowed;
                                '   Bit 9 for sponsorship allowed (Not used);
                                '   Bit 10 for Local Only flag (SSLOCALONLY)
                                '   Bit 11 for Feed spot only flag (SSFEEDONLY)
                                '   Bit 12 indicates if across midnight (0=No; 1=Yes) (SSXMID)
                                '   Bit 13 indicates if avail overbooked
    iLen As Integer             'Length in seconds
    ianfCode As Integer         'Avail name code for booking
    iOrigUnit As Integer        'Original Units, this is only set if avail is overbooked
    iOrigLen As Integer         'Original length, this is only set if avail is overbooked
End Type

Type CSPOTSS
    iRecType As Integer         '(Bits 15-0 Left to right):
                                'Bit 0-3 = record type (contract spot) = 10
                                'Bit  Info
                                '  4  Open BB Buy
                                '  5  Close BB Buy
                                '  6  Split Network Primary spot    '  6  Floater BB Buy
                                '  7  Split Network Secondary spot  '  7  Any BB Buy
                                '  8  Donut Buy
                                '  9  Bookend Buy
                                ' 10  Avail Buy
                                ' 11  Pre-emptible
                                ' 12  Exclude Avail Buy
                                ' 13  Library buy
                                ' 14  Exclusions defined
    iRank As Integer            'Bits 0-10 = Number of Quarter Hours that this spots can be booked into
                                '            (End Time - Start Time) * Days\900
                                '            (10a - 6a) * 5 \ 900
                                '            (36000 - 21600) * 5 \ 900
                                '            14400 * 5 \ 900
                                '            80
                                '            Note:  Start and Time time are the true avail range of times
                                '                   Max price value is 1000
                                '                   Special values:
                                '                   1010=Direct Response(R); 1020=Remnant(T); 1030=per Inquiry(Q); 1040=Trade;
                                '                   1050=Promo(M); 1060=PSA(S); 1070=Reservation; 1045=Extra
                                '                   2000 is temporary used in SpotMG to indicate that this is a fill spot
                                'Bits 11-14 = Price Level (0=Fill, 1= N/C; 2-15 obtained from flight price)
                                'Special quarter hours:
    lSdfCode As Long            'Sdf Auto code
    lBkInfo As Long             '(Bits 15-0 Left to right):
                                '  Bits 0-16 for Booked Date (obtained from gDateValue);
                                '  Bits 17-22 for start minutes within hour if avail is in same hour as start time (otherwise 0).  Start Time of line is 6:10am and booked avail is 6:25, therefore 10 will be stored
                                '  Bits 23-28 for end minutes within hour if avail is in the same hour as the end time (otherwise 60)
                                '  Bit  29= Solo Avail
                                '  Bit  30= 1st position
                                '  Bits 0-16 and 17-22 and 23-28 set but not used
    iMnfComp(0 To 1) As Integer 'Competitive code
    iPosLen As Integer          '(Bits 15-0 Left to right):
                                '  Bits 0-11 for Spot length
                                '  Bits 12-15 for position
    iAdfCode As Integer         'Advertiser code number
End Type

Type BBSPOTSS   'Billborad spot
    iRecType As Integer         '(Bits 15-0 Left to right):
                                'Bit 0-3 = record type (BB spot) = 11
                                'Bit  Info
                                '  4  BB Buy only
                                '5-12 Not used
    iRank As Integer            'Not used
    lSdfRecNo As Long           'Sdf record number (not code number)
    lChfCode As Long            'Contract header code
    iMnfComp(0 To 1) As Integer 'Competitive code
    iLineNo As Integer          'Line number
    iUnused1 As Integer
End Type

'Type FSPOTSS
'    iRecType As Integer         '(Bits 15-0 Left to right):
'                                'Bit 0-3 = record type (Feed spot) = 12
'                                'Bit  Info
'                                '4-12 not used
'    lFsfRecNo As Long           'Fsf code number (fsf removed)
'    lBkInfo As Long             '(Bits 15-0 Left to right):
'                                '  Bits 0-16 for Booked Date (obtained from gDateValue);
'                                '  Bits 17-22 for start hour (in minutes);
'                                '  Bits 23-28 for end hour (in minutes)
'    iMnfComp(0 To 1) As Integer 'Competitive code
'    iPosLen As Integer          '(Bits 15-0 Left to right):
'                                '  Bits 0-11 for BB length
'                                '  Bits 12-15 for position (not required)
'    iAdfCode As Integer         'Advertiser code number
'    iRank As Integer            'Not Used
'End Type

'Type PSPOTSS
'    iRecType As Integer         '(Bits 15-0 Left to right):
'                                'Bit 0-3 = record type (PSA spot) = 13
'                                'Bit  Info
'                                '4-12 not used
'    lCrfRecNo As Long           'Crf code number
'    lBkInfo As Long             '(Bits 15-0 Left to right):
'                                '  Bits 0-16 for Booked Date (obtained from gDateValue);
'                                '  Bits 17-22 for start hour (in minutes);
'                                '  Bits 23-28 for end hour (in minutes)
'    iMnfComp(0 To 1) As Integer 'Competitive code
'    iPosLen As Integer          '(Bits 15-0 Left to right):
'                                '  Bits 0-11 for BB length
'                                '  Bits 12-15 for position (not required)
'    iAdfCode As Integer         'Unused
'    iRank As Integer            'Not defined yet- Used to determine if spot can be preempted
'End Type

'Spot summary record layout
Type SSF
    iType                 As Integer         ' 0=Non-Game Image; 1-nn is Game image
    iFillRequired         As Integer         ' Avail needs to be filled (Y/N). Test for Y
    iCount                As Integer         ' Number of Programs, avails and spots
    iVefCode              As Integer         ' Vehicle code
    iDate(0 To 1)         As Integer         ' Date of summary
    iStartTime(0 To 1)    As Integer         ' Start Time of this record(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    lCode                 As Long            ' Auto Increment
    'tPas(1 To 1200)       As AVAILSS         '5-10-06 chged from 1000 to 1200 entries, Note field within DDF- this field contains the three subrecords defined above
    tPas(0 To 1199)       As AVAILSS         '5-10-06 chged from 1000 to 1200 entries, Note field within DDF- this field contains the three subrecords defined above
End Type

Type SSFKEY0
    iType As Integer             '0=On Air; 1=Altered (partial day defined)
    iVefCode As Integer             'Vehicle code
    iDate(0 To 1) As Integer        'Date of summary
    iStartTime(0 To 1) As Integer   'Start Time of this record(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

Type SSFKEY1
    iVefCode As Integer             'Vehicle code
    iType As Integer                '0=Regular Programming; 1-NN = Sports Programming (Game number)
End Type

Type SSFKEY2
    iVefCode As Integer             'Vehicle code
    iDate(0 To 1) As Integer        'Date of summary
End Type

Type SSFKEY3
    lCode                 As Long
End Type

'********************************************************
'
'Spot Tracking File Definition
'
'********************************************************

'Spot Tracking record layout
Type STF
    lCode As Long    'Internal code number (AutoInc)
    iVefCode As Integer             'Vehicle code
    lChfCode As Long                'Contract code
    iLineNo As Integer              'Line number
    iCreateDate(0 To 1) As Integer  'Creation Date (date this record created)
    iCreateTime(0 To 1) As Integer  'Creation Time (time this record created) (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sAction As String * 1           'A=Added; R=Removed
    iLogDate(0 To 1) As Integer     'Log Date
    iLogTime(0 To 1) As Integer     'Log Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sPrint As String * 1            'R=Ready to print; P=Already Printed; D=Delete
    iLen As Integer                 'Spot length
    lSdfCode As Long                'Sdf Code
    iRotNo As Integer               'Prior Rotation #
    iUrfCode As Integer             'User Code
    lFsfCode As Long
    sUnused As String * 6
End Type

'Spot tracking key record layout
Type STFKEY0
    iVefCode As Integer             'Vehicle code
    iLogDate(0 To 1) As Integer     'Log Date
    iLogTime(0 To 1) As Integer     'Log Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

'Stf key record layout- use LONGKEY0
'Type STFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type


'********************************************************
'
'Spot Week Dump file definition
'
'*********************************************************

'Swf record layout
Type SWF
    iUrfCode As Integer        'User code
    iVefCode As Integer     'Vehicle code
    iType As Integer        '0=Header- fields contain dates; 1=Events contained within fields
    lSeqNo As Long       '2-2-05 chg to long; Sequence number (0 To --)
    sDay1 As String * 550   'Date or events
    sDay2 As String * 550
    sDay3 As String * 550   'Date or events
    sDay4 As String * 550
    sDay5 As String * 550   'Date or events
    sDay6 As String * 550
    sDay7 As String * 550   'Date or events
    lCode As Long           'Internal code number (AutoInc)
End Type

'Swf key record layout
Type SWFKEY0
    iVefCode As Integer     'Vehicle code
    iType As Integer        '0=Header- fields contain dates; 1=Events contained within fields
    lSeqNo As Long       '2-9-05 chg to long; Sequence number (0 To --)
End Type

'Swf key record layout- use LONGKEY0
'Type SWFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type


'******************************************************************************
' sxf_Spot_Extension Record Definition
'
'******************************************************************************
Type SXF
    lCode                 As Long            ' Auto-increment field
    lSdfCode              As Long            ' Spot detail Internal eference
                                             ' code
    sType                 As String * 1      ' Record type: G=was MG; W=From
                                             ' Work Area
    lChfCode              As Long            ' Contract header internal
                                             ' reference code
    iLineNo               As Integer         ' Contract line number
    iAdfCode              As Integer         ' Advertiser internal reference
                                             ' code
    iMissedVefCode        As Integer         ' Vehicle internal reference code.
                                             ' For type = G: Vehicle spot was MG
                                             ' on prior to being Missed; W:
                                             ' Missed Vehicle
    iMissedDate(0 To 1)   As Integer         ' Missed date of the spot. For
                                             ' Type=G: MG date prior to being
                                             ' Missed; W: Missed date
    iMissedTime(0 To 1)   As Integer         ' Missed time. For Type G: MG Time
                                             ' prior to being Missed; W: Missed
                                             ' Time
    iMissedGameNo         As Integer         ' Missed game number. For Type =
                                             ' G:MG game number prior to being
                                             ' Missed; W: Missed game number
    iEnteredDate(0 To 1)  As Integer         ' Date entered
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    iUrfCode              As Integer         ' User internal reference code
    sUnused               As String * 10     ' Unused
End Type


'Type SXFKEY0
'    lCode                 As Long
'End Type

Type SXFKEY1
    sType                 As String * 1
    lSdfCode              As Long
End Type

Type SXFKEY2
    sType                 As String * 1
    iMissedDate(0 To 1)   As Integer
End Type
'******************************************************************************
' pcf_Pod_CPM_Cntr Record Definition
'
'******************************************************************************
Type PCF
    lCode                 As Long            ' Contract podcast Ad Server buy auto
                                             ' code reference
    lChfCode              As Long            ' Contract header reference code
    iPodCPMID             As Integer         ' Podcast CPM buy ID.  Similar to
                                             ' line number
    iCntRevNo             As Integer         ' Contract revision number
    iPropVer              As Integer         ' Contract proposal number
    iVefCode              As Integer         ' Vehicle reference code
    iRdfCode              As Integer         ' Daypart reference code(Ad Location)
    sType                 As String * 1      ' S=Standard; P=Package; H=Hidden
    iPkCPMID              As Integer         ' CPM ID reference of package line
    sPriceType            As String * 1      ' Price Type C=CPM F=Flat rate
    iPositionMnfCode      As Integer         ' Mnf Reference Code
    iPriority             As Integer         ' 0-5
    iCopyTypeMnfCode      As Integer         ' Copy type Mnf reference code
    lRafCode              As Long            ' Ad server vendor geo-target
                                             ' reference code
    iStartDate(0 To 1)    As Integer         ' Start date of buy
    iEndDate(0 To 1)      As Integer         ' End date of buy
    lImpressionGoal       As Long            ' Impressions Goal for buy
    lImpressionCap        As Long            ' Impression Cap
    lPodCPM               As Long            ' Podcast CPM value (xxxx.xx). From
                                             ' Rate Card
    lTotalCost            As Long            ' Ordered total cost
    lCxfCode              As Long            ' comment reference code
    iUrfCode              As Integer         ' Last user who modified CPM buy
    sDelete               As String * 1      ' Blank or N = No Y=Yes
    iEntryDate(0 To 1)    As Integer         ' Entry date of CPM buy
    iEntryTime(0 To 1)    As Integer         ' Entry time of CPM buy
    sHideCBS              As String * 1      ' Do not show a 'cancel before
                                             ' start' line on the BR
    iExtImpSource         As Integer         ' 1= first party 2 =third party
    lExtLineNo            As Long            ' assigned by the external system Boostr Only
    iDeliveryType         As Integer         ' Ad Delivery Type  0=DAI,
                                             ' 1=Baked-in
    iLen                  As Integer         ' Impression Length
    sExtVendorLineID      As String * 70     ' Vendor assigned LineID for this
                                             ' contract line / UUID
    iAvfCode              As Integer         ' AdServer vendor accociated with
                                             ' this digital line
    iAdServDlvyStatus     As Integer         ' Delivery Status: 0=N/A, 1=not
                                             ' pushed, 2=pushed, 3=partial,
                                             ' 4=Issue Encountered, 5=Requires
                                             ' Repush
    sUnused               As String * 100      ' Unused
End Type


Type PCFKEY0
    lCode                 As Long
End Type

Type PCFKEY1
    lChfCode              As Long
    iPodCPMID             As Integer
    iCntRevNo             As Integer
    iPropVer              As Integer
End Type

Type PCFKEY2
    lChfCode              As Long
    iVefCode              As Integer
End Type

Type PCFKEY3
    iEndDate(0 To 1)      As Integer
End Type

Type PCFKEY4
    iStartDate(0 To 1)    As Integer
End Type

Type PCFKEY5
    lChfCode              As Long
    sType                 As String * 1
End Type

