Attribute VB_Name = "RECDEFAL"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RECDEFAL.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  CDFKEY0                       CDFKEY1                       CHFKEY3                   *
'*                                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RecDefAL.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions A-L
Option Explicit


Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Type FIELDAREA
    fBoxX As Single  'Box X position
    fBoxY As Single  'Box Y position
    fBoxW As Single  'Box Width
    fBoxH As Single  'Box Height
    iReq As Integer  'Input required
    iChg As Integer  'Field changed flag (if so, show as bold)
    iAlign As Integer 'Align: 0= left (LEFTJUSTIFY), 1= Right(RIGHTJUSTIFY), 2= Center(CENTER)
    sShow As String  'String to be shown
    'iUpArrowIndex As Integer    'ID Index to go to for up arrow
    'iDnArrowIndex As Integer    'ID Index to go to for down arrow
    'sCaption As String  'Title to be shown in top left of box
End Type

'********************************************************
'
'General structures used to populate list boxes with names and codes
'
'*********************************************************

'Extract record filter layout
Type POPCHARTYPE
    sType As String * 1
End Type

'Extract record filter layout
Type POPINTEGERTYPE
    iType As Integer
End Type
Public Const POPINTEGERTYPEPK As String = "I"

'Extract record filter layout
Type POPDATETYPE
    iDate0 As Integer
    iDate1 As Integer
End Type

'Extract record filter layout
Type POPSTRINGTYPE
    sType As String
End Type

'Extract Record layout (record returned from btrieve)
Type POPICODENAME
    iCode As Integer
    sChar As String * 100   'Arbitary size: must be larger then any name
End Type
Public Const POPICODENAMEPK As String = "IB100"

'Extract Record layout (record returned from btrieve)
Type POPLCODENAME
    lCode As Long
    sChar As String * 100   'Arbitary size: must be larger then any name
End Type
Public Const POPLCODENAMEPK As String = "LB100"

'Extract Record layout (record returned from btrieve)
Type POPICODE
    iCode As Integer
End Type
Public Const POPICODEPK As String = "I"

'Extract Record layout (record returned from btrieve)
Type POPLCODE
    lCode As Long
End Type
Public Const POPLCODEPK As String = "L"

'Record structure used to obtain the first record so extract will work
Type IPOPREC
    iCode As Integer
    'sChar(1 To 1024)  As Byte   'String * 1  'Arbitary size: must be larger then any record
    sChar(0 To 1023)  As Byte   'String * 1  'Arbitary size: must be larger then any record
End Type

'Record structure used to obtain the first record so extract will work
Type LPOPREC
    lCode As Long
    'sChar(1 To 5030)  As Byte   'String * 1  'Arbitary size: must be larger then any record
    sChar(0 To 5029)  As Byte   'String * 1  'Arbitary size: must be larger then any record
End Type

'Record structure used to obtain the first record so extract will work
Type ARRAYREC
    sChar(0 To 1023)  As String * 1  'Arbitary size: must be larger then any record
End Type

'Sort\Code
Type SORTCODE
    sKey As String * 190
End Type

'********************************************************
'
'Move record definition
'
'********************************************************
Type MOVEREC
    sChar As String * 1024
End Type

'********************************************************
'
'Key record definition
'
'********************************************************
'Integer key record layout
Type INTKEY0
    iCode As Integer
End Type

Type LONGKEY0
    lCode As Long
End Type

'********************************************************
'
'Advertiser file definition
'
'*********************************************************

'Adf record layout
Type ADF
    iCode As Integer    'Internal code number for advertiser
    sName As String * 30  'Name
    sAbbr As String * 7 'Abbreviation
    sProduct As String * 35 'Product name
    iSlfCode As Integer 'Salesperson code number
    iAgfCode As Integer 'Agency code number
    sBuyer As String * 20   'Buyers name
    sCodeRep As String * 10 'Rep advertiser Code
    sCodeAgy As String * 10
    sCodeStn As String * 10 'Station advertiser Code
    iMnfComp(0 To 1) As Integer 'Competitive code
    iMnfExcl(0 To 1) As Integer 'Program Exclusions code
    sCppCpm As String * 1    'P=CPP; M=CPM; N=N/A (was sPriceType)
    sDemo(0 To 3) As String * 6    'First-four Demo target
    iMnfDemo(0 To 3) As Integer 'Mnf Demo copy number
    lTarget(0 To 3) As Long  'CPP or CPM target (xxxxx.xx)
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    sPaymRating As String * 1   '0=Quick pay; 1=Normal pay; 2= Slow pay; 3= Difficult; 4=in Collection
    sShowISCI As String * 1     'Y=Yes; N=No
    iMnfSort As Integer     'Invoice sorting code
    sBillAgyDir As String * 1   'A=Bill Agency; D=Bill Advertiser Directly
    sCntrAddr(0 To 2) As String * 40 'Contract address
    sBillAddr(0 To 2)  As String * 40   'Billing address
    iArfLkCode As Integer   'Lock box code
    sPhone As String * 14   'Phone number and extension
    sFax As String * 10     'Fax number
    iArfCntrCode As Integer   'Contract EDI service code number
    iArfInvCode As Integer    'Invoice EDI service code number
    sCntrPrtSz As String * 1 'W=Widw; N= Narrow
    'sSlsTax(0 To 1) As String * 1   'Y= Yes; N=No
    iTrfCode As Integer     'Tax rate reference
    sCrdApp As String * 1   'Direct Only: Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)  TTP 10460 - RAB export: add advertiser credit status and agency credit status
    sCrdRtg As String * 5   'Credit rating (AAA, B,..)
    iPnfBuyer As Integer    'Personnel Buyer Code Number
    iPnfPay As Integer      'Personnel Payable Clerk Code Number
    iPct90 As Integer    '% Over 90 days (xxx)
    sCurrAR As String * 6   'Current A/R (xx,xxx,xxx.xx)
    sUnbilled As String * 6 'Unbilled amount (xx,xxx,xxx.xx)
                            'Computed in SetCrdit:
                            'CreditRestr  Value
                            '  P, N, L    Unbilled spots Plus Projected spots for Spf.iRNoWks
                            '  W          Unbilled spots plus 1 week of actual spots
                            '  M          Unbilled spots plus 4 weeks of actual spots
                            '  T          Unbilled spots plus 13 weeks of actual spots
                            'Note: Todays dates is included with the Unbilled (post Log)
    sHiCredit As String * 6 'High credit amount (xx,xxx,xxx.xx)
    sTotalGross As String * 6   'Total gross (xx,xxx,xxx.xx)
    iDateEntrd(0 To 1) As Integer   'Date entered
    iNSFChks As Integer     'Number returned checks
    iDateLstInv(0 To 1) As Integer  'Date last invoiced
    iDateLstPaym(0 To 1) As Integer 'Date of last payment
    iAvgToPay As Integer    'Average number of days to pay
    iLstToPay As Integer    'Last number of days to pay
    iNoInvPd As Integer 'Number of invoices payed
    sNewBus As String * 1   'Y=Yes; N=No (when advertiser added- set to Y;
                            'When first contract reference this advertiser,
                            'this field is set to N
    iEndDate(0 To 1) As Integer   'Latest contract date referencing this advertiser
    iMerge As Integer   'Merge code number
    iUrfCode As Integer 'Last user to modify record
    sState As String * 1    'A=Active; D=Dormant
    iCrdAppDate(0 To 1) As Integer  'Date credit approval changed
    iCrdAppTime(0 To 1) As Integer  'Time credit approval changed
    sPkInvShow As String * 1    'Package Invoice air times as D=Daypart; T=Time
    lGuar As Long        '% Guarantee (xxx) or Grimp Guarantee (xxxxx) from contract
    iLastYearNew          As Integer         ' Last Year advertiser set as New
    iLastMonthNew         As Integer         ' Last Month advertiser treated as New
    sBkoutPoolStatus      As String * 1      ' Blackout Pool Status (A=Active; D=Deactivated; N or Blank = Not part of Pool.  Test for A or D.
    sUnused2 As String * 5       'Not Used
    sRateOnInv As String * 1    'Show Rate on Invoice (Y/N).  Test for N
    iMnfBus As Integer  'Business Category from contract
    iUnused1 As Integer    'Not Used.
    sAllowRepMG As String * 1   'Allow Rep Posted spots to be counted as MG (Y/N)
    sBonusOnInv As String * 1   'Show Bonus spots on invoices (Y/N)
    sRepInvGen As String * 1    'Rep Invoice Generated Internally (I) or Externally (E)
    iMnfInvTerms As Integer     'Terms to show on invoive (Stored with mnfType = "J" and mnfUnitType = "T")
    sPolitical As String * 1      ' Airing Political spots (Y/N)
    sAddrID As String * 9       'Used for sorting and type ahead when we have multi-addresses for direct advertiser
End Type

'ADf key record layout- use INTKEY0
'Type ADFKEY0
'    iCode As Integer
'End Type


'Type ADFKEY3
'    sCodeRep As String * 10 'Rep advertiser Code
'End Type
Type ADFKEY1
    sCodeRep As String * 10 'Rep advertiser Code
End Type
'10209
'******************************************************************************
' ADFX_Advertisers Record Definition
'
'******************************************************************************
Type ADFX
    iCode                 As Integer         ' Primary Key 1-to-1 match with
                                             ' adfCode value of ADF_Advertisers
                                             ' table
    sRefID                As String * 36     ' Advertiser GUID. External
                                             ' identifier (RefId) used for the
                                             ' export
    sDirectRefId          As String * 36     ' Direct Advertiser Ref id for WO
                                             ' invoice export
    lCrmId                As Long            ' CRM ID-for RAB adv ID
    iInvFeatures          As Integer         ' Bit: 0=None, 1=Suppress Net
                                             ' Amount for trade invoices
    sUnused               As String * 58
End Type
'Advertiser extract record
Type ADFEXT
    iCode As Integer
    sName As String * 30
    sAbbr As String * 7 'Abbreviation
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
    sCrdApp As String * 1      'Credit Approved...(R=Requires Credit Check;A=Approved;D=Denied)  TTP 10460 - RAB export: add advertiser credit status and agency credit status
    iMnfSort As Integer     'Invoice sorting code
    sBillAgyDir As String * 1   'A=Bill Agency; D=Bill Advertiser
    sState As String * 1    'A=Active; D=Dormant
    sAllowRepMG As String * 1   'Allow Rep Posted spots to be counted as MG (Y/N)
    sBonusOnInv As String * 1   'Show Bonus spots on invoices (Y/N)
    sRepInvGen As String * 1    'Rep Invoice Generated Internally (I) or Externally (E)
    sFirstCntrAddr As String * 40 'Contract address
    sPolitical As String * 1        'Political flag (y/n)   5-26-06
    sAddrID As String * 9
    iTrfCode As Integer
    iArfInvCode As Integer    'Invoice EDI service code number
    iPnfBuyer As Integer
    iLastYearNew As Integer  '9-6-11 New Bus last Years Adv is New
    iLastMonthNew As Integer '9-6-11 New bus last years Month Adv is new
    iMnfInvTerms As Integer     '9-6-13
    sCodeStn As String * 10     '1-9-18
    iPnfPay As Integer          '10-14-19
End Type
Public Const ADFEXTPK As String = "IB30B7BBIBBBBBB40BB9IIIIIIB10I"            '1-9-18, 10-14-10 pnfpay

'********************************************************
'
'Agency file definition
'
'*********************************************************

'Agf record layout
Type AGF
    iCode As Integer    'Internal code number for agency
    sName As String * 40  'Name
    sAbbr As String * 5 'Abbreviation
    sCityID As String * 5   'City ID
    iComm As Integer 'Commssion (xx.xx)
    iSlfCode As Integer 'Salesperson code number
    sBuyer As String * 20   'Buyers name
    sCodeRep As String * 10 'Rep Agency Code
    sCodeStn As String * 10 'Station Agency Code
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    sPaymRating As String * 1   '0=Quick pay; 1=Normal pay; 2= Slow pay; 3= Difficult; 4=in Collection
    sShowISCI As String * 1     'Y=Yes; N=No
    iMnfSort As Integer     'Invoice sorting code
    sCntrAddr(0 To 2) As String * 40 'Contract address
    sBillAddr(0 To 2)  As String * 40   'Billing address
    iArfLkCode As Integer   'Lock box code
    sPhone As String * 14   'Phone number and extension
    sFax As String * 10     'Fax number
    iArfCntrCode As Integer   'Contract EDI service code number
    iArfInvCode As Integer    'Invoice EDI service code number
    sCntrPrtSz As String * 1 'W=Widw; N= Narrow
    'sSlsTax(0 To 1) As String * 1   'Y= Yes; N=No
    iTrfCode As Integer     'Tax rate reference
    sCrdApp As String * 1   'Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)
    sCrdRtg As String * 5   'Credit rating (AAA, B,..)
    iPnfBuyer As Integer    'Personnel Buyer Code Number
    iPnfPay As Integer      'Personnel Payable Clerk Code Number
    iPct90 As Integer    '% Over 90 days (xxx)
    sCurrAR As String * 6   'Current A/R (xx,xxx,xxx.xx)
    sUnbilled As String * 6 'Unbilled amount (xx,xxx,xxx.xx)
                            'Computed in SetCrdit:
                            'CreditRestr  Value
                            '  P, N, L    Unbilled spots Plus Projected spots for Spf.iRNoWks
                            '  W          Unbilled spots plus 1 week of actual spots
                            '  M          Unbilled spots plus 4 weeks of actual spots
                            '  T          Unbilled spots plus 13 weeks of actual spots
                            'Note: Todays dates is included with the Unbilled (post Log)
    sHiCredit As String * 6 'High credit amount (xx,xxx,xxx.xx)
    sTotalGross As String * 6   'Total gross (xx,xxx,xxx.xx)
    iDateEntrd(0 To 1) As Integer   'Date entered
    iNSFChks As Integer     'Number returned checks
    iDateLstInv(0 To 1) As Integer  'Date last invoiced
    iDateLstPaym(0 To 1) As Integer 'Date of last payment
    iAvgToPay As Integer    'Average number of days to pay
    iLstToPay As Integer    'Last number of days to pay
    iNoInvPd As Integer 'Number of invoices payed
    iMerge As Integer   'Merge code number
    iUrfCode As Integer 'Last user to modify record
    sState As String * 1    'A=Active; D=Dormant
    iCrdAppDate(0 To 1) As Integer  'Date credit approval changed
    iCrdAppTime(0 To 1) As Integer  'Time credit approval changed
    sPkInvShow As String * 1    'Package Invoice air times as D=Daypart; T=Time
    sUnused2 As String * 2
    sAgfRefID As String * 12       'External station agency id
    sCntrExptForm As String * 1 'C=CSI Form; O=Agency (OMD) form.  Default is C
    iMnfInvTerms As Integer     'Terms to show on invoice  (Stored with mnfType = "J" and mnfUnitType = "T")
    sXMPProposalBand As String * 2  'Used to override Site Option Band on XML Proposals
    s1or2DigitRating      As String * 1      ' 1 or 2 Digit research rating output. 1=One digit rating; 2=Two digit rating. Test for 2.
    sXMLCallLetters As String * 4  'Used to override Site Option Call Letters on XML Proposals
    sXMLDates As String * 1     'Used which flight dates: Aired(A) or M-Su(M or blank).
                                'Aired means set start date to the first date spots airs and end date to last day spot airs as defined by the flights.
    sUnused               As String * 9    '
End Type

'Agf key record layout- use INTKEY0
'Type AGFKEY0
'    iCode As Integer
'End Type

'Type AGFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

Type AGFKEY1
    sCodeRep As String * 10 'Rep advertiser Code
End Type
'10209
'******************************************************************************
' AGFX_Agencies Record Definition
'
'******************************************************************************
Type AGFX
    iCode                 As Integer         ' Primary Key 1-to-1 match with
                                             ' agfCode value of AGF_Agencies
                                             ' table
    sRefID                As String * 36     ' Agency GUID. External identifier
                                             ' (RefId) used for the export
    lCrmId                As Long            ' CRM ID-for RAB agency ID
    iInvFeatures          As Integer         ' Bit: 0=None,1=Supress Net Amount
                                             ' for trade invoices
    sUnused               As String * 94
End Type

'Agency extract record
Type AGFEXT
    iCode As Integer
    sName As String * 40
    sCityID As String * 5   'City ID
    sCreditRestr As String * 1  'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
    sCrdApp As String * 1       '(Direct only) Credit Approved (R=Requires Credit Check; A=Approved; D=Denied)
    iMnfSort As Integer     'Invoice sorting code
    sState As String * 1    'A=Active; D=Dormant
    iTrfCode As Integer
    iArfInvCode As Integer    'Invoice EDI service code number
    iPnfBuyer As Integer
    iCommPct As Integer         '6-16-11
    iMnfInvTerms As Integer     'Terms to show on invoice  (Stored with mnfType = "J" and mnfUnitType = "T")
    s1or2DigitRating      As String * 1      ' 1 or 2 Digit research rating output. 1=One digit rating; 2=Two digit rating. Test for 2.
    sCodeStn As String * 10             '1-9-18
    iPnfPay As Integer          '10-14-19
    sRefID As String * 12 '4-20-21 use RefID for External Agency ID (CustRevExport)
End Type
Public Const AGFEXTPK As String = "IB40B5BBIBIIIIIBB10IB12"       '6-16-11 add another integer reference, 1-9-18 added codestn, 19-14-19 add pnfpay; 22-5-16 added sCreditApp

'******************************************************************************
' ALF_Avail_Lock Record Definition
'
'******************************************************************************
Type ALF
    lCode                 As Long            ' AutoIncrement field
    iVefCode              As Integer         ' Vehicle Reference code
    sLockType             As String * 1      ' Lock Type: A=Avail; S=Spots
    iDate(0 To 1)         As Integer         ' Lock Date or zero if by Game
    iGameNo               As Integer         ' Game Number or zero if by date
    iStartTime(0 To 1)    As Integer         ' Start Time of Lock
    iEndTime(0 To 1)      As Integer         ' End Time of Lock
    iUrfCode              As Integer         ' User Reference that created lock
    sUnused               As String * 20     ' Unused
End Type


'Type ALFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type ALFKEY1
    iVefCode              As Integer
    iDate(0 To 1)         As Integer
End Type

Type ALFKEY2
    iVefCode              As Integer
    iGameNo               As Integer
End Type


'********************************************************
'
'Avail Name file definition
'
'*********************************************************

'Anf record layout
Type ANF
    iCode As Integer        'Internal code number for Avail Name
    sName As String * 20    'Identification name
    sSustain As String * 1  'Sustain allowed (Y/N)
    'sUnused1 As String * 1  'Unused field
    sState As String * 1    'A=Active; D=Dormant
    sSponsorship As String * 1 'Can the avail be sponsored (Y/N)
    iMerge As Integer       'Merge code number
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode
    sBookLocalFeed As String * 1    'Book: L= Local Spots Only, F = Feed Spots Only; B = Both
    sRptDefault As String * 1   'Combo Report default (Y/N)
    iSortCode             As Integer         ' Sort Code
    sTrafToAff            As String * 1      ' Determine if spots within the
                                             ' avail will be sent from Traffic
                                             ' to the Affiliate system. Y/N.
                                             ' Test for N
    sISCIExport           As String * 1      ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the ISCI Export. Y/N. Test for N
    sAudioExport          As String * 1      ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the audio delivery system (XDS,
                                             ' Wegener, Starguide). Y/N Test for
                                             ' N
    sAutomationExport     As String * 1      ' Determine if spots within the
                                             ' avail will be included as part of
                                             ' the automation export. Y/N. Test
                                             ' for N
    sFillRequired         As String * 1      ' Does the break need to be filled
                                             ' (Y or Blank=Yes; N=No; A=Only if
                                             ' Partially booked and treat
                                             ' Adjacent breaks a one break).
                                             ' Test for N.
    iEventAvailsGroup     As Integer         ' Event Avails Group - 0=Standard,
                                             ' 1=Billboard, 2=Drop-in, 3=Extra
    sUnused               As String * 50
End Type

'Anf key record layout- use INTKEY0
'Type ANFKEY0
'    iCode As Integer
'End Type

'********************************************************
'
'Name/Address file definition
'
'*********************************************************

'Arf record layout
Type ARF
    iCode As Integer    'Internal code number for Name/Address
    sType As String * 1 'L=Lock box; A=Agency DP; S=Sales Office (not Used); C=Content Provider; P = Producer; N=Network/Rep; K=Network Producer; F=FTP (Used in Vehicle Option)
    sID As String * 10  'Identification name (Type L and A); Abbreviation (Type C and P and N and K)
    sNmAd(0 To 3)  As String * 25   'Name and address
    iMerge As Integer   'Merge code number
    sName As String * 40    'For Type C and P and N and K
    sEMail As String * 250 'E-Mail address for Type C and P and N and K
    sFTP As String * 250     'FTP address for type C and P and N and K
    sWebSite As String * 50 'was WebSite for 70 char
    sContactFax As String * 20  'Contact Fax # for type C and P and N and K
    sSendISCITo As String * 1   'Send ISCI Export to Producer or Content Provider, Type C and P
    sContactName As String * 40 'Contact Name
    sContactPhone As String * 30    'Contact Phone number
    sSepInvByVeh As String * 1     'Separate EDI Invoices by vehicle (Y/N); Test for Y.  For Type A
    sEDIMediaType         As String * 1      ' EDI override Site Option Media
                                             ' Type.
    sUnused               As String * 18
End Type

'Arf key record layout- use INTKEY0
'Type ARFKEY0
'    iCode As Integer
'End Type

'********************************************************
'
'Alert User file definition
'
'*********************************************************

'Adf record layout
Type AUF
    lCode                 As Long            ' Internal code number for alert
                                             ' user file
    iEnteredDate(0 To 1)  As Integer         ' Entered Date
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    sStatus               As String * 1      ' Alert Status (C=Cleared;
                                             ' R=Requires Alert menu to be
                                             ' shown)
    sType                 As String * 1      ' Alert Type (C=Contract; L=Log
                                             ' Reprinting required; F=Final Log
                                             ' Generated; R=Reprint Log
                                             ' Generated; B=Block Users;
                                             ' N=Notify User;G=Generate AST)
    sSubType              As String * 1      ' Alert Sub Type;
                                             ' AufType=C:C=Proposal chgd to
                                             ' Complete; AufType=L:C=Copy
                                             ' Assigned;S=Spot Scheduled;M=New Missed in Closed Week
                                             ' aurType=F:  I=ISCI Export; S=Spot
                                             ' Export; AufType = R:  I=ISCI
                                             ' Export;S=Spot
                                             ' Export;AufType=B:Block Traffic
                                             ' (T),Affiliate(A),Both(B);aufType=
                                             ' G:T=Date/Time
    lChfCode              As Long            ' Contract Code: aufType = C
    iVefCode              As Integer         ' Vehicle Code: aufType = L; When
                                             ' first created (selling vehicle
                                             ' will be used; When generating
                                             ' Logs or click on Menu Alert, the
                                             ' selling vehicle will be replaced
                                             ' with airing vehicle
    iMoWeekDate(0 To 1)   As Integer         ' Monday week date: aufType = L
    iCreateUrfCode        As Integer         ' User Code of person who created
                                             ' the Alert
    iCreateUstCode        As Integer         ' User Code of person who created
                                             ' the Alert
    iClearUrfCode         As Integer         ' User code of person who cleared
                                             ' the Alert on Traffic system
    iClearUstCode         As Integer         ' User code of person who cleared
                                             ' the Alert on Affiliate System
    sClearMethod          As String * 1      ' Clear method: M=Manually;
                                             ' A=Automatically
    iClearDate(0 To 1)    As Integer         ' Cleared Date
    iClearTime(0 To 1)    As Integer         ' Cleared Time
    lUlfCode              As Long            ' User Log reference code
    lCefCode              As Long            ' Notification Comment (aufType =
                                             ' N)
    iCountdown            As Integer         ' Initial countdown value for shut
                                             ' down
    sSpotChg(0 To 6)      As String * 1      ' Spot changed on Monday->Sunday (Y/N). Test for Y.
    sGenericCopyChg(0 To 6) As String * 1    ' Generic copy changed on Monday->Sunday
                                             ' (Y/N). Test for Y
    sRegionCopyChg(0 To 6) As String * 1     ' Region copy changed on Monday->Sunday
                                             ' (Y/N). Test for Y
    sUnused As String * 10
End Type

'Auf key record layout- use LONGKEY0
'Type AUFKEY0
'    lCode As Long
'End Type

'Auf key record layout
Type AUFKEY1
    sType As String * 1     'Alert Type (C=Contract; L=Log)
    sStatus As String * 1   'Alert Status (C=Cleared; R=Requires Alert menu to be shown)
End Type

Type AUFKEY2
    sType                 As String * 1
    sStatus               As String * 1
    sSubType              As String * 1
    iVefCode              As Integer
    iMoWeekDate(0 To 1)   As Integer
End Type

Type AUFKEY3
    sType                 As String * 1
    sStatus               As String * 1
    sSubType              As String * 1
    lChfCode              As Long
End Type

Type AUFVIEW
    sKey As String * 20
    tAuf As AUF
End Type


'********************************************************
'
'Blackout file definition
'
'*********************************************************

'Bof record layout
Type BOF
    lCode As Long   'AutoInc
    sType As String * 1 'S=Supression; R=Replacement
    iAdfCode As Integer 'Advertiser Code
    lSifCode As Long    'Short Title or Product
    iVefCode As Integer 'Vehicle Code (required for bofType = S only)
    lCifCode As Long    'Copy Inventory Code
    iMnfComp(0 To 1) As Integer 'Product Protection Codes (Required for bofType = R)
    iStartDate(0 To 1) As Integer   'Start date
    iEndDate(0 To 1) As Integer 'End Date (TFN allowed)
    sDays(0 To 6) As String * 1 'Days (Y=Yes; N=No) Index 0 = Monday; 1= Tuedays;...
    iStartTime(0 To 1) As Integer   'Start Time
    iEndTime(0 To 1) As Integer 'End Time
    iUrfCode As Integer 'User Code (Last to add or modify)
    lSChfCode As Long   'Suppress Contract Code
    iRAdfCode As Integer    'Replace Advertiser Code
    lRChfCode As Long   'Replace Contract Code
    iLen As Integer
    sSource As String * 1   'N=Export NY; L=Log (Default)
    sUnused As String * 9
End Type

'Bof key record layout- use LONGKEY0
'Type BOFKEY0
'    lCode As Long
'End Type

'Bof key record layout
Type BOFKEY1
    sType As String * 1 'S=Supression; R=Replacement
    iEndDate(0 To 1) As Integer 'End Date (TFN allowed)
End Type


'********************************************************
'
'Budget by Salesperson file definition
'
'*********************************************************

'Bsf record layout
Type BSF
    iSlfCode As Integer 'Salesperson Code
    iYear As Integer    'Year (xxxx)
    iStartDate(0 To 1) As Integer  'Start Date (currently not used)
    lGross(0 To 53) As Long    'Rate for each week (xxxxxxx.xx)
    iUrfCode As Integer 'Last user to modify record
    iSeqNo As Integer   'Sequence Number used if Year data split between years
    lCode As Long       'Internal code number (AutoInc)
End Type

'Bsf key record layout
Type BSFKEY0
    iYear As Integer    'Year (xxxx)
    iSeqNo As Integer   'Sequence Number used if Year data split between years
    iSlfCode As Integer 'Salesperson Code
End Type

'Bvf key record layout- use LONGKEY0
'Type BSFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'********************************************************
'
'Budget by Vehicle file definition
'
'*********************************************************

'Bvf record layout
Type BVF
    lCode As Long    'Internal code number (AutoInc)
    iSofCode As Integer 'Sales Office Code
    iVefCode As Integer 'Vehicle Code
    iYear As Integer    'Year (xxxx)
    iStartDate(0 To 1) As Integer  'Start Date (currently not used)
    iMnfBudget As Integer   'Budget names stored into Mnf
    sSplit As String * 1    'Split Flag
    lGross(0 To 53) As Long    'Rate for each week (xxxxxxx.xx)
    iUrfCode As Integer 'Last user to modify record
    iSeqNo As Integer   'Sequence number (used for split year if required)
End Type

'Bvf key record layout
Type BVFKEY0
    iYear As Integer    'Year (xxxx)
    iSeqNo As Integer
    iMnfBudget As Integer   'Budget names stored into Mnf
End Type

'Bvf key record layout- use LONGKEY0
'Type BVFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'Bvf key record layout- use INTKEY0
'Type BVFKEY2
'    iMnfBudget As Integer   'Budget names stored into Mnf
'End Type

Type BVFEXT
    iVefCode As Integer 'Vehicle Code
    iYear As Integer    'Year (xxxx)
    iSeqNo As Integer   'Sequence number (used for split year if required)
    iMnfBudget As Integer   'Budget names stored into Mnf
End Type
Public Const BVFEXTPK As String = "IIII"

'******************************************************************************
' CAF_Copy_Air_Game Record Definition
'
'******************************************************************************
Type CAF
    lCode                 As Long            ' Auto increment reference
    lCrfCode              As Long            ' Copy Rotation Reference code
    sType                 As String * 1      ' Type: G=by Game #; T=by Team
    iGameNo               As Integer         ' Game Number. cafType = G
    iTeamMnfCode          As Integer         ' Team Reference code.  cafType = T
    sUnused               As String * 10     ' Unused
End Type


'CAF key record layout- use LONGKEY0
'Type CAFKEY0
'    lCode                 As Long
'End Type

Type CAFKEY1
    lCrfCode              As Long
End Type

'********************************************************
'
'Copy Combo file definition
'
'*********************************************************

'Ccf record layout
Type CCF
    lCode As Long    'Internal code number for copy combo
    lCifCode(0 To 3) As Long 'Copy inventory code
    lcpfCode As Long 'Product/Agency ID code number
End Type

'Ccf key record layout- use LONGKEY0
'Type CCFKEY0
'    lCode As Long
'End Type

'********************************************************
'
'Comments-Dunning file definition
'
'*********************************************************

'Cdf record layout
Type CDF
    lCode As Long    'Internal code number (AutoInc)
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
    iActionDate(0 To 1) As Integer   'Action Date
    iDateEntrd(0 To 1) As Integer   'Date entered
    iTimeEntrd(0 To 1) As Integer   'Time entered
    iUrfCode As Integer     'User who entered comment
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 1002   'Last two bytes after the comment must be 0
    sComment As String * 1004   'Last bytes after the comment must be 0
End Type

'Cdf key record layout
Type CDFKEY0 'VBC NR
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct) 'VBC NR
End Type 'VBC NR

Type CDFKEY1 'VBC NR
    iAdfCode As Integer 'Advertiser code number 'VBC NR
End Type 'VBC NR

'Cdf key record layout- use LONGKEY0
'Type CDFKEY2
'    lCode As Long    'Internal code number (AutoInc)
'End Type

Type CDFEXT
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
    iActionDate(0 To 1) As Integer   'Action Date
    iDateEntrd(0 To 1) As Integer   'Date entered
    iTimeEntrd(0 To 1) As Integer   'Time entered
    iUrfCode As Integer     'User who entered comment
End Type
Public Const CDFEXTPK As String = "IIIIIIIII"

Type CDFLIST
    lRecPos As Long   'Record position since no code number
    tCdfExt As CDFEXT
End Type

Type CDFAGYEXT
    iAdfCode As Integer 'Advertiser code number
    iDateEntrd(0 To 1) As Integer   'Date entered
    iTimeEntrd(0 To 1) As Integer   'Time entered
End Type
Public Const CDFAGYEXTPK As String = "IIIII"

Type CDFADVTEXT
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iDateEntrd(0 To 1) As Integer   'Date entered
    iTimeEntrd(0 To 1) As Integer   'Time entered
End Type
Public Const CDFADVTEXTPK As String = "IIIII"

'********************************************************
'
'Comment-Events file definition
'
'*********************************************************

'Cef record layout
Type CEF
    lCode As Long    'Internal code number for Comment
    sUnused As String * 1   'Fix length portion of record must be 5 bytes
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 1002   'Last two bytes after the comment must be 0
    sComment As String * 1004   'Last bytes after the comment must be 0
End Type

'Cef key record layout- use LONGKEY0
'Type CEFKEY0
'    lCode As Long
'End Type

'********************************************************
'
'Contract flight file definition
'
'*********************************************************

'Cff record layout
Type CFF
    lChfCode As Long        'Contract code
    iClfLine As Integer        'Line number
    iCntRevNo As Integer     'Contract revision number (was ClfVersion)
    iStartDate(0 To 1) As Integer  'Start Date of line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iEndDate(0 To 1) As Integer  'End Date of line
    iSpotsWk As Integer 'Spots per week, if zero, then daily
    iDay(0 To 6)  As Integer    'Spot per day if daily or flag if weekly
                                'For weekly 1=Air day; 0=not air day
                                'Index 0=Mo; 1=tu,...6=Su
    iXSpotsWk As Integer 'Spots per week, if zero, then daily
    sXDay(0 To 6)  As String * 1    'Spot per day if daily or flag if weekly
                                'For weekly 1=Air day; 0 or blank=not air day
                                'Index 0=Mo; 1=tu,...6=Su
    sDelete As String * 1       'Y=Yes(flight-line deleted); N=No
    lActPrice As Long 'Actual spot price (xxxxxxx.xx)
    lPropPrice As Long    'Proposed Rate card (yield management) Price (xxxxxxx)
    lBBPrice As Long  'Billboard price
    sDyWk As String * 1     'D=Daily spots defined; W=Weekly spots defined
    sPriceType As String * 1 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    iPropVer As Integer     'Proposal version number
    lAdjPrice As Long       'Used to correct ordered price if line is E
    lCode As Long           ' Auto Increment Code
    sUnused As String * 12
End Type

Type CFFLIST
    CffRec As CFF
    lRecPos As Long
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    iNextCff As Integer
    lNextCff As Long    'Used to filter out lines/flights that have more then 32000 records
    lStartDate As Long
    lEndDate As Long
    lAvgAud As Long     'Average audience for user selected demo on demo bar
    lPriDemoAvgAud As Long  'Average audience used for Contract Totals
                            'that show at bottom Right for primary demo
    lPriDemoPop As Long
    iGameNo As Integer
End Type

Type PCFLIST
    PcfRec As PCF
    lRecPos As Long
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    iNextPcf As Integer
    lNextPcf As Long    'Used to filter out lines/flights that have more then 32000 records
    lStartDate As Long
    lEndDate As Long
    lAvgAud As Long     'Average audience for user selected demo on demo bar
    lPriDemoAvgAud As Long  'Average audience used for Contract Totals
                            'that show at bottom Right for primary demo
    lPriDemoPop As Long
    
End Type
'Cff key record layout
Type CFFKEY0
    lChfCode As Long        'Contract code
    iClfLine As Integer        'Line number
    iCntRevNo As Integer     'Contract Revision number (descending)
    iPropVer As Integer     'Proposal version number (descending)
    iStartDate(0 To 1) As Integer  'Start Date of line
End Type

'Cff key record layout- use LONGKEY0
'Type CFFKEY1
'    lCode As Long
'End Type

Type CFFEXT
    iStartDate(0 To 1) As Integer  'Start Date of line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
End Type
Public Const CFFEXTPK As String = "II"

'********************************************************
'
'Contract Game file definition
'
'*********************************************************
Type CGF
    lCode As Long            ' Auto Increment Code
    lClfCode As Long            ' Contract Line file reference
    iGameNo As Integer         ' Game Number
    iNoSpots As Integer         ' Number of spots to air
    iAirDate(0 To 1) As Integer     ' Air Date from Game_Schedule
    sUnused As String * 10     ' Unused
End Type

'Cgf key record layout- use LONGKEY0
'Type CGFKEY0
'    lCode As Long
'End Type

Type CGFKEY1
    lClfCode As Long            ' Contract Line file reference
End Type

Type CGFLIST
    CgfRec As CGF
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    iNextCgf As Integer
    'lNextCgf As Long    'Used to filter out lines/flights that have more then 32000 records
    lStartDate As Long
    lEndDate As Long
End Type

Type CGFSORT
    sKey As String * 15 'Date|Game#
    CgfRec As CGF
End Type

'********************************************************
'
'Contract Header file definition
'
'*********************************************************

'Chf record layout
Type CHF
    lCode As Long    'Internal code number for Contract
    lCntrNo As Long  'Contract number (999,999,999)
    sAgyEstNo As String * 10 'Agenct estimate number Note Total Number obtained by concatenate this field and sTitle
    lVefCode As Long  'Multi-Vehicle File plus short titles
                      'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    'lMasterNo As Long  'Master Contract number
    iExtRevNo As Integer    'External Revision number (revision number that agy/advt sees)
    'iPkageGenMeth As Integer       'Ordered Package Generate $ Method: 0=Week; 1=Line
    iXMLSentUrfCode       As Integer         ' User Reference code that sent out
                                             ' the XML. Was chfPkageGenMeth
                                             ' Ordered Package Generate $
                                             ' Method: 0=Week; 1=Line. Set to -1
                                             ' or 0
    iOHDDate(0 To 1) As Integer  'Order/Hold/Dead Date of Contract
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iOHDTime(0 To 1)  As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iCntRevNo As Integer   'Revision number (was iRevNo)
    iPropDate(0 To 1) As Integer 'Proposal Date of Contract
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iPropTime(0 To 1)  As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sType As String * 1 'Type of contract:
                        'New
                        'C=Standard; V=Reservation; T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
                        'Sch Rank (BkHQ)
                        ' Trade   1040
                        '
                        ' Direct Response  1010
                        ' Remnant          1020
                        ' Per Inquiry      1030
                        ' Promo            1050
                        ' PSA              1060
                        ' Reservation      1070
                        'Note:
                        ' Extra            1045
                        'Old
                        'P=Proposal; H=Hold; V=Reservation; J=Rejection;
                        'E=Order; C=Contract; B=Allocation contract; D=Deferred Contract;
                        'T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
    iAdfCode As Integer 'Advertiser code number
    sProduct As String * 35   'Product (default to advertiser product)
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iSlfCode(0 To 9) As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    lComm(0 To 9) As Long     'Revenue or Commission Share % (xx.xxxx)
    iMnfComp(0 To 1) As Integer 'Competitive code
    iMnfExcl(0 To 1) As Integer 'Program Exclusion code
    sBuyer As String * 20   'Buyers name (default name from advertiser)
    sPhone As String * 14   'Buyers Phone number and extension(if direct: default to advt otherwisw default to agency)
    'sPriceCntr As String * 1   'Y=Show prices on contract; N=Don't show prices
    'sPriceInv As String * 1 'Y=Show prices on invoices; N=Don't show prices
    iEDSSentUrfCode       As Integer         ' User reference code that sent
                                             ' E-Mail to EDS. Replaced
                                             ' chfPriceCntr and chfPriceInv
    iPctBudget As Integer '% share of Budget (xxx)
    iMnfRevSet(0 To 4) As Integer    'Revenue Sets (increased by 3)
    sCppCpm As String * 1    'P=CPP; M=CPM; N=N/A (default from advertiser) (was sPriceType)
    iMerchPct As Integer     'Merchandising Percent (xx.xxx)
    iPromoPct As Integer     'Promotion Percent (xx.xxx)
    iSlspCommPct(0 To 9) As Integer    'Salesperson Commission or Commission Share % (xx.xx)
    iMnfDemo(0 To 3) As Integer     'Mnf Demo Code Number
    lTarget(0 To 3) As Long  'CPP or CPM target (xxxxx.xx) (default from advertiser)
    iPctTrade As Integer 'Percent trade (xxx)
    iRcfCode As Integer 'Rate Card code
    lInputGross As Long   'Total gross dollars (xxxxxxxxx.xx)
    sBillCycle As String * 1 'S=Standard; C=Calendar; W= Week   Not used: T=Package standard; D=Package Calendar
    sInvGp As String * 1        'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    sAdServerDefined        As String * 1      ' Y=Podcast Ad Server defined; N or
                                             ' blank=Podcast Ad Server not defined
    sPodSpotDefined       As String * 1      ' Y=Podcast Spot buy defined; N or
                                             ' blank=Not defined
    lCxfCode As Long        'Comment code number
    lCxfChgR As Long        'Comment code number for change reason
    lCxfInt As Long        'Comment code number for internal
    lCxfMerch As Long        'Comment code number for merchandizising
    lCxfProm As Long        'Comment code number for promotion
    lCxfCanc As Long        'Comment code number cancellation clause
    iPropVer As Integer     'Proposal version
    sMGMiss As String * 1   'Bill MG(G) or Bill Missed(M)
    sStatus               As String * 1      ' Status of proposal and contract:
                                             ' W=Working Proposal; D=Rejected;
                                             ' C=Completed Proposal;
                                             ' I=Unapproved
                                             ' Proposal;G=Unscheduled Hold;
                                             ' H=Hold;N=Unscheduled Order;
                                             ' O=Order(Scheduled order)
    sTitle As String * 10   'Proposal Title Note used with agyestno to obtain the agyestno
    iMnfSubCmpy(0 To 9) As Integer  'Mnf Veh Group 6: Sub-Company
    iDtNeed(0 To 1) As Integer  'Date proposal needed
    iMnfBus As Integer      'Mnf Business Category
    lGuar As Long     'Guarantee Percent (xxx) or grimp guarantee (xxxxx)
    'sUnused3 As String * 2 'this is not used
    iEDSSentExtRevNo       As Integer         ' EDS Sent external version number
    iMnfCmpy(0 To 6) As Integer     'Mnf Competitor Code Number
    iCmpyPct(0 To 6) As Integer  'Company competitor % of buyer (xxx) Note: 0 used for this company
    sResvNew As String * 1  'Use reserved Spots(R) Or Book New Spots(N)
    lChfCode As Long        'Reserve contract to remove spots from
    iMnfPotnType As Integer 'Potential Type code
    sPrint As String * 1    'Contract print state: N=New and not printed; C=Changed and not printed; P=already printed (don't ask revision number until printed)
    sDiscrep As String * 1  'Price discreopancy flag: Y= Total gross and total line prices don't agree; N= prices match
    sNewBus As String * 1   'New business state: Y= advertiser new business flag is Y-> set it to N; N= not new business
    sSchStatus As String * 1   'Scheduling status: (was sStatus)
                            'F=Fully scheduled; M=Manually scheduled (sStatus: "O", "H" & sType: "T", "R", "Q", "S", "M")
                            'P=Prevent scheduling (sStatus: "W", "I", "C", "D" & sType: All); I=Interrupted scheduling
                            'N=New contract (sStatus "O", "H" & sType: "S", "V"); A=Altered Contract (sType: "S", "V"; sStatus "O", "H")
    sAgyCTrade As String * 1  'Agency Commissionable on Trades: Y= Yes; N=No (default)
    iStartDate(0 To 1) As Integer  'Earliest start Date of all lines
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iEndDate(0 To 1) As Integer  'Latest end Date of all lines
    imnfSeg As Integer 'Segment (only used if spfCUseSegment specified On)
    iUrfCode As Integer     'Last user who modified contract
    'sSellNet As String * 1  'Y= selling (replaced by sCBSOrder)
    sCBSOrder As String * 1     'C=Contracted cancelled via MAI Contract Import, or proposals/orders changes
    iHdChg As Integer   'Bit field indicating header changes
                        ' 0     Advertiser changed
                        ' 1     Competitive # 1 changed
                        ' 2     Competitive # 2 changed
                        ' 3     Exclusion # 1 changed
                        ' 4     Exclusion # 2 changed
                        ' 5     Status P (Prevent) changed
                        ' 6     Status M (Manual) changed
    'iVerifyDate(0 To 1) As Integer   'Contract verification date
    'iVerifyTime(0 To 1) As Integer   'Contract verification time
    'sVerifyFlag           As String * 1      ' Contract verfication flags:N=Not
    '                                         ' verfied; S=Sent to Agency;
    '                                         ' V=Verified.  Test for flags
    '                                         ' values. No default
    iEDSSentDate(0 To 1)  As Integer         ' EDS Sent Date, was chfVerifyDate
                                             ' Contract verification date
    iEDSSentTime(0 To 1)  As Integer         ' EDS Sent Time, was chfVerifyTime
                                             ' Contract verfication time
    'sUnused               As String * 1      ' Unused, was chfVerifyFlag
                                             ' Contract verfication flags:N=Not
                                             ' verfied; S=Sent to Agency;
                                             ' V=Verified.  Test for flags
                                             ' values. No default
    sSource               As String * 1      ' Source:Blank = Standard definition;  P=Programmatic; F=File Maker; C=COIN
    'iSlfLock As Integer 'Code of salesperson that lock changes to proposals
    'iLockDate(0 To 1) As Integer    'Date lock inserted on proposal
    iXMLSentExtRevNo      As Integer         ' External Version number of Insertion Order
                                             ' sent via XML to station Traffic
                                             ' system
    iXMLSentDate(0 To 1)  As Integer         ' Date Insertion Order sent via XML
                                             ' to station Traffic system
    sDelete As String * 1   'Y=Deleted (superseded); N=Not superseded
    lSifCode As Long        'Short Title Code
    'lAudMGAdj As Long       'Audience MG/Outside total Adjustment todate
    iXMLSentTime(0 To 1)  As Integer         ' XML Sent Time, was chfAudMGAdj
                                             ' Audience MG/Outside total
                                             ' Adjustment todate
    'iRemoteID As Integer        'Remote ID (Note: Unique Cntr# is Cntr#-RemoteID)
    'lAutoCode As Long            'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    'iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    'iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    'iSourceID As Integer    'Original Remote User ID (avoid resending to sender)
    'iOrigRemoteID As Integer        'Original Remote ID (Used to Find Original Proposal)
    'lOrigAutoCode As Long            'Auto Incr Code (Used to Find Original Proposal)
    lAirTimeGross As Long   'Computed from air time
    '12/17/06-Change to tax by agency or vehicle
    'lVarTaxAmt As Long      'Variable tax amount (xx.xxxx)
    sInstallDefined As String * 1   'Installment SBF records defined (Y/N or blank)
    sRepDBID              As String * 5      ' Rep database ID
    sNRProcessed          As String * 1      ' Net-Rep Processed (Y=Yes; N=No)
    lEffCode As Long        'Extra Field reference code.  Used to store Project Code and SMS code required for Invoice Export
    'sUnused2 As String * 1
    sHideDemo             As String * 1      ' chfHideDemo replaced Unused which
                                             ' Replaced Remote User stuff.
    iPnfBuyer As Integer    'Buyer reference code, used with the buyer name stored above
    lSpotChfCode As Long            ' Auto Increment code initailly assiged to this image of chf.  Used to
                                  ' get MSF associated with this chf.
    sNoAssigned As String * 1   'True Proposal Number Assigned (Y or N); If N then temporary Proposal Number
    lGrImp As Long          'Contract Gross Impression
    lGRP As Long            'Contract GRP
    sNTRDefined As String * 1
    lNTRGross As Long   'compute from NTR
    lExtCntrNo            As Long            ' assigned by the external system Boostr only
    iAvfCode              As Integer         ' where avfName = what external vendor sent Boostr Only
    iAdServDlvyStatus     As Integer         ' Delivery Status: 0=N/A, 1=not pushed, 2=pushed, 3=partial, 4=Issue Encountered, 5=Requires Repush
    sUnused               As String * 48
End Type

'Chf key record layout- use LONGKEY0
'Type CHFKEY0
'    lCode As Long
'End Type

'Chf key record layout
Type CHFKEY1
    lCntrNo As Long
    iCntRevNo As Integer   'Revision number (was iRevNo)
    iPropVer As Integer     'Proposal version
End Type

'Chf key record layout- use INTKEY0
'Type CHFKEY2
'    iAdfCode As Integer 'Advertiser code number
'End Type

Type CHFKEY3 'VBC NR
    sStatus As String * 1   'Status of proposal and contract 'VBC NR
    sNRProcessed As String * 1 'VBC NR
End Type 'VBC NR

Type CHFKEY4
    sAgyEstNo As String * 10 'Agenct estimate number Note Total Number obtained by concatenate this field and sTitle
    sTitle As String * 10   'Proposal Title Note used with agyestno to obtain the agyestno
    iCntRevNo As Integer   'Revision number (was iRevNo)
    iPropVer As Integer     'Proposal version
End Type

Type CHFKEY5
    iXMLSentDate(0 To 1)  As Integer
    iXMLSentTime(0 To 1)  As Integer
End Type

Type CHFKEY6
    iEDSSentDate(0 To 1)  As Integer
    iEDSSentTime(0 To 1)  As Integer
End Type

'Contract extract record (used in gPopCntrSchBox and gPopCntrBox)
Type CHFEXT
    lCode As Long                   'Internal code number for Contract
    lCntrNo As Long                 'Contract number
    iPropDate(0 To 1) As Integer    'Revision Date of Contract
    sType As String * 1             'Type of contract: P=Proposal; H=Hold; V=Reservation; J=Rejection;
                                    'E=Order; C=Contract; B=Allocation contract; D=Deferred Contract;
                                    'T=Remnant; R=Direct Response; Q=Per inQuiry;
    iAdfCode As Integer             'Advertiser code number
    sProduct As String * 35         'Product (default to advertiser product)
    iAgfCode As Integer             'Agency code number (0 if advertiser is direct)
    iSlfCode(0 To 9) As Integer     'Salesperson code number (if direct: default to advt otherwisw default to agency)
    iStartDate(0 To 1) As Integer   'Earliest start Date of all lines
    iEndDate(0 To 1) As Integer     'Latest end Date of all lines
    lVefCode As Long                'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    sSellNet As String * 1          'Y= selling
    iUrfCode As Integer             'Last user who modified contract
End Type
Public Const CHFEXTPK As String = "LLIIBIB35IIIIIIIIIIIIIIILBI"

'Contract extract record (used in gPopAdvtViaCntrBox)
Type CHFADVTEXT
    lCode As Long    'Internal code number for Contract
    lCntrNo As Long  'Contract number
    iExtRevNo As Integer    'External Revision Number
    iCntRevNo As Integer   'Revision number (was iRevNo)
    sType As String * 1 'Type of contract:
    iAdfCode As Integer 'Advertiser code number
    sProduct As String * 35   'Product (default to advertiser product)
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iSlfCode(0 To 9) As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    iMnfDemo0 As Integer     'Mnf Demo Code Number
    lCxfInt As Long        'Comment code number for internal
    iPropVer As Integer     'Proposal version
    sStatus As String * 1   'Status of proposal and contract
    iMnfPotnType As Integer 'Potential Type code
    iStartDate(0 To 1) As Integer  'Earliest start Date of all lines
    iEndDate(0 To 1) As Integer  'Latest end Date of all lines
    lVefCode As Long 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    lSifCode As Long        'Short Title Code
    iPctTrade As Integer    '8-20-05 %trade, 0 = 100% cash
    sCBSOrder As String * 1
    sBillCycle As String * 1
    sSource As String * 1
    sAdServerDefined As String * 1
End Type
Public Const CHFADVTEXTPK As String = "LLIIBIB35IIIIIIIIIIIILIBIIIIILLIBBBB"

'TTP 10937 - Contract Dashboard: new screen to show contract header information
Type CHFDETAILLIST
    lCode As Long                   'Internal code number for Contract
    lCntrNo As Long                 'Contract number
    iExtRevNo As Integer            'External Revision Number
    iCntRevNo As Integer            'Revision number (was iRevNo)
    sType As String * 1             'Type of contract:
    iAdfCode As Integer             'Advertiser code number
    sProduct As String * 35         'Product (default to advertiser product)
    iAgfCode As Integer             'Agency code number (0 if advertiser is direct)
    iSlfCode(0 To 9) As Integer     'Salesperson code number (if direct: default to advt otherwisw default to agency)
    iMnfDemo0 As Integer            'Mnf Demo Code Number
    lCxfInt As Long                 'Comment code number for internal
    iPropVer As Integer             'Proposal version
    sStatus As String * 1           'Status of proposal and contract
    iMnfPotnType As Integer         'Potential Type code
    iStartDate(0 To 1) As Integer   'Earliest start Date of all lines
    iEndDate(0 To 1) As Integer     'Latest end Date of all lines
    lVefCode As Long                'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    lSifCode As Long                'Short Title Code
    iPctTrade As Integer            '8-20-05 %trade, 0 = 100% cash
    sCBSOrder As String * 1         'Canceled Before start
    sBillCycle As String * 1        'Bill Cycle
    sSource As String * 1           'Source
    sAdServerDefined As String * 1  'Ad Server Defined Y/N
    sAirtimeDefined As String * 1   'Airtime Defined Y/N
    sNTRDefined As String * 1       'NTR Defined Y/N
    sSchStatus As String * 1        'Schedule Status
    iOHDDate(0 To 1) As Integer     'Last Changed Date
    lInputGross As Long             'Total gross dollars (xxxxxxxxx.xx)
    iDelvyStatus As Integer         'Delivery status (0=N/A, 1=not pushed, 2=pushed, 3=partial, 4=Issue Encountered, 5=Requires Repush)
End Type
Public Const CHFDETAILLISTPK As String = "LLIIBIB35IIIIIIIIIIIILIBIIIIILLIBBBBBBBIILI"

'********************************************************
'
'Copy Inventory file definition
'
'*********************************************************

'Cif record layout
Type CIF
    lCode As Long    'Internal code number for Copy Inventory
    iMcfCode As Integer 'Media code number
    sName As String * 5 'Name
    sCut As String * 1  'Cut
    sReel As String * 10    'Reel
    iLen As Integer         'Copy length
    iEtfCode As Integer 'Event Type code number
    iEnfCode As Integer 'Event Name code number
    iAdfCode As Integer 'Advertiser code number
    lcpfCode As Long 'Product/Agency ID code number
    iMnfComp(0 To 1) As Integer 'Competitive code
    iMnfAnn As Integer 'Announcer code
    sHouse As String * 1    'Tape in House Y/N
    sCleared As String * 1  'Tape cleared Y/N
    lCsfCode As Long 'Live script code
    iNoTimesAir As Integer  'Number of times aired
    sCartDisp As String * 1 'Cart Disposition (N=N/A; S=Save; P=Purge; A=Ask after Expired)
    sTapeDisp As String * 1 'Tape Disposition (N=N/A; R=Return; D=Destroy; A=Ask after Expired)
    sPurged As String * 1   'Purge flag (A=Active; P=Purged; H=History)
    iPurgeDate(0 To 1) As Integer  'Purged Date
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iDateEntrd(0 To 1) As Integer  'Date entered
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iUsedDate(0 To 1) As Integer   'Last Used Date
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iRotStartDate(0 To 1) As Integer  'Earliest Rotation date using this inventory
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iRotEndDate(0 To 1) As Integer  'Latest Rotation date using this inventory
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iUrfCode As Integer     'Last user who modified inventory
    sPrint As String * 1    'Printed: N=New or Not Printed; P=Printed
    iLangMnfCode As Integer     'Language Reference to MNF
    iInvSentDate(0 To 1)  As Integer         ' Inventory sent to vCreative date. cifCleard used do determine vCreative status
    sUnused               As String * 4
End Type

'Cif key record layout- use LONGKEY0
'Type CIFKEY0
'    lCode As Long    'Internal code number for Copy Inventory
'End Type

Type CIFKEY1
    iMcfCode As Integer 'Media code number
    sName As String * 5 'Name
    sCut As String * 1  'Cut
End Type

'Type CIFKEY2
'    lcpfCode              As Long
'End Type

'Type CIFKEY3
'    iadfCode              As Integer
'End Type

Type CIFKEY4
    sCleared              As String * 1
End Type

'Type CIFKEY5- use POPDATETYPE
'    iDateEntrd(0 To 1) As Integer  'Date entered
'End Type

'Type CIFKEY6- use POPDATETYPE
'    iPurgeDate(0 To 1) As Integer  'Purged Date
'End Type

Type CIFEXT
    lCode As Long    'Internal code number for Copy Inventory
    iMcfCode As Integer 'Media code number
    sName As String * 5 'Name
    sCut As String * 1  'Cut
    iLen As Integer         'Copy length
    lcpfCode As Long 'Product/Agency ID code number
    sPurged As String * 1   'Purge flag (A=Active; P=Purged; H=History)
    iRotStartDate(0 To 1) As Integer  'Latest Rotation date using this inventory
    iRotEndDate(0 To 1) As Integer  'Latest Rotation date using this inventory
    lCsfCode As Long 'Live script code
    sReel As String * 10
End Type
Public Const CIFEXTPK As String = "LIB5BILBIIIILB10"

Type CIFINVEXT
    lCode As Long    'Internal code number for Copy Inventory
    sName As String * 5 'Name
    sCut As String * 1  'Cut
    sPurged As String * 1   'Purge flag (A=Active; P=Purged; H=History)
    iRotEndDate(0 To 1) As Integer  'Latest Rotation date using this inventory
    iPurgeDate(0 To 1) As Integer  'Purged Date
End Type
Public Const CIFINVEXTPK As String = "LB5BBIIII"

'********************************************************
'
'Contract Line file definition
'
'*********************************************************

'Clf record layout
Type CLF
    lChfCode As Long        'Contract code
    iLine As Integer        'Line number
    iCntRevNo As Integer     'Contract Revision number (was iVersion)
    iVefCode As Integer     'Vehicle Code (- => combo. )
    iRpfCode As Integer     'Rate Card Program/Time Code
    sBB As String * 1       'NO Used- instead iBBOpenLen and iBBCloseLen used
                            'BB: N=N/A; O=Open BB; C=Close BB; B= Open/Close BB;
                            'F=Floating; 1=1 Open/week; 2=1 Close/week; 4=1 Open & Close/Week;
                            '3=1 Floating/Week; A=Any
    sExtra As String * 1    'Extra Spot Types: N=N/a; C=Commercial Promo (create breaks if not found);
                            'D=Donuts(two spots within break- any position but one spot between);
                            'B=Bookends (two spots within break-first and last)
    sPgmTime As String * 1  'If library buy retain spots with P=Program buy; T=Time buy when library moved
    iBreak As Integer       'Break number if buying a specific break
    iPosition As Integer    'Position number if buying a specific break and position
    iStartTime(0 To 1) As Integer 'Override Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                            'If not defined-> use rate card times (not defined is hund sec = 1 all other times = 0)
    iEndTime(0 To 1) As Integer 'Override End Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iNoGames As Integer     'Number of games to distribute spot across
    iSpotsOrdered As Integer    'Number of spots to be disttibuted across the number of games
    iSpotsBooked As Integer 'Total spots booked for a game buy
    iSpotsWrite As Integer  'Number of spots written onto the log (for late weekly buys only)
    sCntPct As String * 1   'Entered # of spots/week and spots/day are by: 0= Daily by count; 1=weekly by count;2=Daily by %; 3=weekly by %
                            'The scheduler will convert 2 and 3 to 0 & 1. 2 and 3 are for
                            'reservation contract only
    iLen As Integer         'Spot length
    sPreempt As String * 1  'N=Non-preemptible (spot can't be moved when scheduling other spots);
                            'P=Preemptible
    'sActPrice As String * 5 'Actual spot price (xxxxxxx.xx)
    'sPropPrice As String * 5    'Proposed Rate card (yield management) Price (xxxxxxx.xx)
    'sHighPrice As String * 5    'Highest previous sold spot
    'sAvgPrice As String * 5 'Average previous sold price
    'sLowPrice As String * 5 'Lowest previous sold price
    'sBBPrice As String * 5  'Billboard price
    iPrefStartTime(0 To 1) As Integer   'Preferred start time
    iPrefEndTime(0 To 1) As Integer     'Preferred end time
    sPrefDays(0 To 6) As String * 1     'Preferred Days: Y=Yes, N=No; 0=Mo, 1=Tu,...6=Su
    iPctAllocation As Integer           'Percent allocation (Participant Revenue %, xx.xx)
    lAcquisitionCost As Long            'Cost to purchase time at a station
    sSoloAvail As String * 1            'Only spot allowed in avail.  The avail length must match the spot length to be booked into
    sLiveCopy As String * 1                       '11-7-05 L = Live, R = Recorded (default)
    sSportsByWeek         As String * 1      ' Flag indicating if Event selection was by Week or by
                                             ' Numbers (W=Week; N=Number).  Test for W.
    iEDSSentExtRevNo      As Integer         ' EDS Sent External Rev Number
    iXMLSentExtRevNo      As Integer         ' XML Sent External Rev Number
    sLineChgd             As String * 1      ' Line Changed (Y/N). Test for Y.
    sHideCBS As String * 1  'Do not show a Cancel Before Start line on BR
    sPriceType As String * 1 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    iPriority As Integer    'Priority flag (calculated)
    lCxfCode As Long        'Comment code
    sSchStatus As String * 1  'Schedule status:
                            'F=Fully scheduled; M=Manually scheduled (sType: D; T; R; Q; S or M)
                            'P=Prevent scheduling (sType: P; H; J; E); I=Interrupted scheduling
                            'N=New contract (sType: C; B; or A); A=Altered Contract
    iUrfCode As Integer     'Last user who modified line
    sDelete As String * 1   'Y=Yes(line deleted); N=No
    iEntryDate(0 To 1) As Integer  'Entry Date of Line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iEntryTime(0 To 1)  As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iMnfDemo As Integer     'Mnf Demo Code # (obtained from the first one from header)
    lCPM As Long 'CPP price (xxxxx.xx)
    lCPP As Long 'CPP price (xxxxx.xx)
    lGrImp As Long     'Gross Impressions
    iDnfCode As Integer 'Rating Book that was used in computing above values
    iRdfCode As Integer 'Daypart Code
    iPropVer As Integer 'Proposal version number
    iStartDate(0 To 1) As Integer  'Start Date of Line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iEndDate(0 To 1) As Integer  'End Date of Line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iMnfSocEco As Integer   'Soc/Eco Group Last Referenced if specified at update time
    sType As String * 1 'Line Type: S=Standard; O=Ordered Package(Virtual); A=Aired Package(Real); H=Hidden; E=Equal Price Package
    iPkLineNo As Integer    'Package line number (type = O or A), only for Type = H
    iAdvtSepFlag As Integer 'Advertiser separation flag: 0=Usedline; -1=Determine from multi-lines; nnn=Use nnn as number of spots (this # is set if two or more lines match Vehicle, daypart and # spots don't very)
    iBBOpenLen As Integer   'Billboard spot open length
    iBBCloseLen As Integer  'Billboard spot close length
    sOV2DefinedBits As String * 1   'Field defined within line Bit Map (Right to Left): 0=Allocation; 1=Acquisition; 2=1st Position
    lCode As Long           'AutoInc
    lghfcode As Long        'Game Header Reference
    sGameLayout As String * 1   'Y=Include pattern on BR (Game #- Visiting @ Home); N=Don't include game information
    lRafCode As Long        'Split Network Region definition
    sACT1LineupCode       As String * 11     ' ACT1 Lineup Code
    sACT1StoredTime       As String * 1      ' use ACT1 Stored Time (T or
                                             ' Blank). Test for T
    sACT1StoredSpots      As String * 1      ' use ACT1 Stored Spots (S or
                                             ' Blank). Test for S
    sACT1StoreClearPct    As String * 1      ' Use ACT1 Clearance Percent (C or
                                             ' Blank). Test for C
    sACT1DaypartFilter    As String * 1      ' use ACT1 Daypart Filter (F or
                                             ' Blank). Test for F
    sUnused               As String * 10     ' Unused
End Type

'Clf key record layout
Type CLFKEY0
    lChfCode As Long        'Contract code
    iLine As Integer        'Line number
    iCntRevNo As Integer    'Contract Revision number number
    iPropVer As Integer     'Proposal version number
End Type

Type CLFKEY1
    lChfCode As Long        'Contract code
    iVefCode As Integer     'Vehicle Code (- => combo. )
End Type

'Cif key record layout- use LONGKEY0
'Type CLFKEY2
'    lCode As Long    'Internal code number for Copy Inventory
'End Type

Type CLFKEY3
    lghfcode As Long        'Game Header Reference
    iEndDate(0 To 1) As Integer  'End Date of Line
End Type

Type CLFKEY4
    lRafCode As Long        'Split Network Region definition
    iEndDate(0 To 1) As Integer  'End Date of Line
End Type

'Clf key record layout
Type CLFEXT
    lChfCode As Long        'Contract code
    iLine As Integer        'Line number
    iCntRevNo As Integer    'Contract Revision number
    iVefCode As Integer     'Vehicle Code (- => combo. )
    sSchStatus As String * 1  'Schedule status:
    iPropVer As Integer     'Proposal Version Number
End Type
Public Const CLFEXTPK As String = "LIIIBI"

Type CLFLIST
    sKey As String * 67 'Sort Line "|" Vehicle Group #  Vehicle Sort #  Vehicle Name  Daypart Sort #  Line #  Revision #  Proposal Version #
                        '    5      1      3                  3             40              3            4        3             3
                        'Sort Line  Conventional: 9999~;  Package: Line#A;  Hidden:  Package Line #C
                        'Vehicle Group #: First participant group number (mnfGroupNo)
                        'Vehicle Sort:  sort number from vehicle
                        'Vehicle name:  Vehicle name
                        'Daypart Sort #: 999 or Item Sort #(rifSort) or daypart sort (rdfSort if rifSort not found)
    ClfRec As CLF
    lRecPos As Long
    lCxfRecPos As Long
    iStatus As Integer  '0=New; 1=old and retain; 2=old and delete; -1=Not used
    iFirstCff As Integer
    iFirstCgf As Integer
    lFirstCff As Long       'Used to filter out lines/flights that have more then 32000 records
    lFirstCgf As Long       'Used to filter out lines/flights that have more then 32000 records
    sGameLayout As String * 1
    iCancel As Integer  'True=Canceled before start; False = not canceled
    iOverride As Integer    'True=Allow subset of times to be defined
    iGame As Integer    'True = Game buy; False = Non-Game buy
    iLibBuy As Integer  'True= Library buy; False = Time buy
    iInPast As Integer  'True=Line has flights in the past which are scheduled
    lUnbilledDate As Long    'date of first unbilled spot in past
    lEndBilledDate As Long  'Billing End Date
    iResearch As Integer    'True=Research data is valid; False=Research data must be computed
    iSPChg As Integer     'Line Spot or Price changed(Y= Don't show totals if using Compute Button)
    lCost As Long       'Total Cost (used when computing Research)
    lAvgAud As Long     'Avg Aud
    iAvgRating As Integer   'Average rating
    lGrImp As Long      'Gross Impressions
    lGRP As Long        'GRP
    lCPM As Long
    lCPP As Long
    lPop As Long    'Population
    'lRCCost As Long       'Rate Card Total Cost (used when computing Research)
    'iRCAvgRating As Integer   'Rate Card Average rating
    'lRCGrImp As Long      'Rate Card Gross Impressions
    'lRCGRP As Long        'Rate CardGRP
    'lRTCost As Long      'Request Total Cost
    lTSpots As Long  'Total number of spots
    iSaveIndex As Integer   'Save and Comment Index
    iOrigAnySpots As Integer    'Line have any spots originally(-1 = No Spots & no Price; 0=No Spots & Price defined, 1=Spots defined)
    iLineSchd As Integer 'True= Line scheduled; False = New Line and not scheduled
    sCurrentPrice As String * 12 'Current way price is to be set for a week (from Line price field)
    iPriceSet As Integer    '0=This flag not set; 1=Price not set; 2=Price set
    lPriDemoPop As Long     'Population used for Contract Totals
                            'that show at bottom Right for primary demo
    iPriDemoBook As Integer 'dnf reference used for Contract Totals
                            'that show at bottom Right for primary demo
    iPriChgd As Integer      'If Vehicle; Daypart or Overrides changed
    lPriDemoAvgAud As Long  'Primary Demo Average audience
    iPriDemoAvgRating As Integer  'Primary Demo Average Rating
    lPriDemoGRP As Long
    lPriDemoCPP As Long
    lPriDemoCPM As Long
    lPriDemoGrImp As Long
    lLLD As Long
    iMonThruToDaysSpots As Integer  'Number of spots in past within current week (Monday thru Today)
    iChgBilledPrice As Integer  'Allow billed price to be changed
    '12/16/14: Add Hide/Show Hidden Lines
    bShowLine As Boolean
    bPkExpanded As Boolean
    iAcqCostSpotRate As Integer '0=Indepentent; 1=Acquisition always zero; 2=Acquisition Cost must match the Spot Rate
    sOwnership As String * 1    'A=Owned-Network; B=Owned-Station: C=Unowned-Network; D=Unowned-Station.  Defaul A or Blank. Test for B, C and D
                                'Proposal/order acquisition rules
                                'If Network(Owned or unowned): Use current acquisition rules
                                'If Station-Owned: Bypass Acquisition cost and set to zero
                                'If Station-Unowned: Acquisition Cost and Spot rate will match
                                'Proposal/Order package overrides copy to hidden line rules
                                'If Network(Owned or Unowned): Retain current rules (don't violated the hidden times/days)
                                'If Station (Owned or Unowned): Copy package overrides to hidden lines (allow hidden line times/days to be violated)
    sLineChgd As String * 1   'Line changed (Vehicle, Daypart, Overrides, Length, Spots, Acquisition, Spot rate)
    lFirstAvail As Long
    bLnPriceSet As Boolean  'Used to know that resaerch numbers can be computed as tab thru line. New Lines price change is false, this is used to treat new lines as price changed
    sSourceFilter As String * 1 'A=Show on Air Time tab; S=Show on Podcast Spot tab
    lCxfCode As Long 'TTP 10193 - ReRate Comment
End Type

Type LNAVAILINFO
    lNextIndex As Long
    iInv(0 To 51) As Integer
    lSpotCount(0 To 51) As Long  'Spot count by week
    iSAvail(0 To 51) As Integer     'Spot avail
    iAPAvail(0 To 51) As Integer    'Advertiser or Protection avail
    lStartDate As Long
    sLineType As String * 1
End Type



'********************************************************
'
'Comment- Boilerplate file definition
'
'*********************************************************

'Cmf record layout
Type CMF
    lCode As Long    'Internal code number for Comment-Contract
    sTitle As String * 20   'Boilerplate title
    sComType As String * 1  'O=Other; L=Line; C=Cancellation; R=Change Reason;
                            'M=Merchandising; P=Promotion; I=Internal
    sShProp As String * 1   'Show on Propsals (Y/N)
    sShOrder As String * 1  'Show on Orders (Y/N)
    sShSpot As String * 1   'Show on Spots/Logs (Y/N)
    sShInv As String * 1    'Show on Invoices (Y/N)
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 5002   'Last two bytes after the comment must be 0
    sComment As String * 5004   'Last bytes after the comment must be 0
End Type

'Cmf key record layout- use LONGKEY0
'Type CMFKEY0
'    lCode As Long
'End Type

'********************************************************
'
'Copy Instruction file definition
'
'*********************************************************

'Cnf record layout
Type CNF
    lCrfCode As Long    'Rotation code number
    iInstrNo As Integer 'Instruction number
    lCifCode As Long    'Copy inventory or Copy combo code (minus if combo)
    lCode As Long       'Internal code number
End Type

'Cnf key record layout
Type CNFKEY0
    lCrfCode As Long    'Rotation code number
    iInstrNo As Integer 'Instruction number
End Type

'Cnf key record layout- use LONGKEY0
'Type CNFKEY1
'    lCode As Long
'End Type

'********************************************************
'
'Corporate Calendar file definition
'
'*********************************************************

'Cof record layout
Type COF
    iYear As Integer    'Corporate Calendar Year
    iStartMnthNo As Integer 'Start Month Number
    'iNoWks(1 To 12) As Integer  'Number of weeks for each months (index 1 for start month)
    iNoWks(0 To 11) As Integer  'Number of weeks for each months (index 1 for start month)
    'iStartDate(0 To 1, 1 To 12) As Integer
    'iEndDate(0 To 1, 1 To 12) As Integer
    iStartDate(0 To 1, 0 To 11) As Integer
    iEndDate(0 To 1, 0 To 11) As Integer
End Type

'Cof key record layout- use INTKEY0
'Type COFKEY0
'    iYear As Integer    'Corporate Calendar Year
'End Type


'********************************************************
'
'Copy Product/Agency (ISCI) file definition
'
'*********************************************************

'Cpf record layout
Type CPF
    lCode As Long    'Internal code number for copy Product/Agency
    sName As String * 35    'Name
    sISCI As String * 20    'Agency ISCI code
    sCreative As String * 30    'Creative title
    iRotEndDate(0 To 1) As Integer  'Latest Rotation date using this inventory
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    lSifCode As Long    'Short Title
End Type

'Cpf key record layout- use LONGKEY0
'Type CPFKEY0
'    lCode As Long
'End Type

Type CPFKEY1
    sISCI As String * 20    'Agency ISCI code
End Type

'********************************************************
'
'Copy Rotation Header file definition
'
'*********************************************************

'Crf record layout
Type CRF
    lCode As Long       'Internal code number for copy rotation
    iEtfCode As Integer 'Event Type code number
    iEnfCode As Integer 'Event Name code number
    iAdfCode As Integer 'Advertiser code number
    lChfCode As Long    'Contract code number
    iVefCode As Integer 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    sRotType As String * 1  'A=Contract spot; O=Open BB; C=Close BB; E=Any BB; D=Donut; B=Bookend; M=Cmml Promo
                            'S=PSA into contract avails; R=Promo into contract avail; X=Event type
    iRotNo As Integer   'Rotation number
    iStartDate(0 To 1) As Integer  'Earliest start Date of all lines
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iEndDate(0 To 1) As Integer  'Latest end Date of all lines
    iStartTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                            'If not defined-> use rate card times (not defined is hund sec = 1 all other times = 0)
    iEndTime(0 To 1) As Integer 'End Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sDay(0 To 6)  As String * 1 'Index 0 = Mon, 1= Tue,... Y=Allowed day; N=Not allowed day
    iLen As Integer     'Spot length
    sInOut As String * 1    'N=N/A; I=Book into avail specified; O=Exclude avail specified
    ianfCode As Integer     'Avail name code not to book into
    iNextFinal As Integer   'Next copy instruction to assign within final log
    iNextPrelim As Integer  'Next copy instruction to assign within the prelimary log
    lCsfCode As Long        'Comment
    iEntryDate(0 To 1) As Integer  'Entry Date of the rotation instructions
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    iFeedDate(0 To 1) As Integer   'Feed date of transmission via Bulk Feed
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
    sFeedStatus As String * 1   'S=Sent; R=Ready; P=Suppress
    'lMixRotCode As Long     'Rotation header code of mixed rotation header
    lCvfCode As Long        'Copy Vehicle reference code. Used when more than one vehicle defined with a rotation. Replaced crfMixRotCode
    sZone As String * 3     'Zone name (Note: R is used to indicate Regional Rotation)
    iUrfCode As Integer     'Last user who modified line
    iVersionDate(0 To 1) As Integer 'Date version created or modified
    iNoTimesMod As Integer  'Number of times only the dates changed
    iEarliestDateAssg(0 To 1) As Integer    'Earliest date spot assigned (spot date)
    iLatestDateAssg(0 To 1) As Integer  'Latest date spot assigned (spot Date)
    iLastDateAssg(0 To 1) As Integer    'Last date spot assigned (spot date)
    iDateAssgDone(0 To 1) As Integer    'Date that last assignment was performed
    iTimeAssgDone(0 To 1) As Integer    'Time that last assignment was performed
    lSifCode As Long    'Short Title (place here or in cpf)
    lRafCode As Long    'Region Area Code (note: sZone will contain a R)
    iAffFdDate(0 To 1) As Integer  'Affiliate Feed Date
    sAffFdStatus As String * 1   'S=Sent; R=Ready; P=Suppress
    iAffFdWk As Integer     '7 week bit map (using right 7 bits with most left bit for week of date in AffFdDate)
    iAffTranDate(0 To 1) As Integer  'Affiliate Transmission Date
    sAffXMitChar As String * 1      'Last Letter speicified when transmitted
    lFsfCode As Long
    iKCFdDate(0 To 1) As Integer  'KenCast: Feed Date
    sKCFdStatus As String * 1   'KenCast: S=Sent; R=Ready; P=Suppress
    iKCFdWk As Integer     'KenCast: 7 week bit map (using right 7 bits with most left bit for week of date in AffFdDate)
    iKCTranDate(0 To 1) As Integer  'KenCast: Transmission Date
    sKCXMitChar As String * 1      'KenCast: Last Letter speicified when transmitted
    iKCNoTimesMod As Integer  'Number of times only the dates changed
    sLiveCopy As String * 1                       '11-7-05 L = Live, R = Recorded (default)
    sMGCopyAssign As String * 1 'MG, Outside and Fill Copy Assigment: O=from Original Vehicle only; S=from Scheduled Vehicle only; B=Try Schedule first, then Original Vehicle
    sState As String * 1        'A=Active; D=Dormant (test for D only)
    iLangMnfCode As Integer     'Language Reference to MNF
    iBkoutInstAdfCode As Integer    'Blackout Instruction advertiser reference code
    sAirGameType As String * 1      ' Air Game Type: D or
                                    ' Blank=Day/Time; G=By Game #; T =
                                    ' By Teams.  Test for G or T.
    sUnused As String * 2       'unused, 11-7-05 chged from 8 to 7 for live copy flag
End Type

'Crf key record layout- use LONGKEY0
'Type CRFKEY0
'    lCode As Long       'Internal code number for copy rotation
'End Type

Type CRFKEY1
    sRotType As String * 1
    iEtfCode As Integer 'Event Type code number
    iEnfCode As Integer 'Event Name code number
    iAdfCode As Integer 'Advertiser code number
    lChfCode As Long    'Contract code number
    lFsfCode As Long
    iVefCode As Integer 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    iRotNo As Integer   'Rotation number
End Type

Type CRFKEY2
    sRotType As String * 1
    iAdfCode As Integer 'Advertiser code number
    iEndDate(0 To 1) As Integer  'Latest end Date of all lines
End Type

Type CRFKEY3
    sRotType As String * 1
    iBkoutInstAdfCode As Integer    'Blackout Instruction advertiser reference code
    iEndDate(0 To 1) As Integer  'Latest end Date of all lines
End Type

Type CRFKEY4
    sRotType As String * 1
    iEtfCode As Integer 'Event Type code number
    iEnfCode As Integer 'Event Name code number
    iAdfCode As Integer 'Advertiser code number
    lChfCode As Long    'Contract code number
    lFsfCode As Long
    iRotNo As Integer   'Rotation number
End Type

'********************************************************
'
'Copy script file definition
'
'*********************************************************

'Csf record layout
Type CSF
    lCode As Long    'Internal code number for Comment-Contract
    iAdfCode As Integer 'Advertiser code number
    sType As String * 1 'S=Copy inventory scripts; R=Rotation comment
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 5002   'Last two bytes after the comment must be 0
    sComment As String * 5004   'Last bytes after the comment must be 0
End Type

'Csf key record layout- use LONGKEY0
'Type CSFKEY0
'    lCode As Long
'End Type


'********************************************************
'
'Comment- Contract and line file definition
'
'*********************************************************

'Cxf record layout
Type CXF
    lCode As Long    'Internal code number for Comment-Contract
    sComType As String * 1  'Comment Type(O=Other; L=Contract Line; C=Cancellation;
                            'R=Change Reason; M=Merchandising; P=Promotion; I=Internal; N=Personnel;
                            'D=Invoice Disclaimer; G=Game Multi-Media)
    sShProp As String * 1     'Show on Proposal Y/N
    sShOrder As String * 1 'Show on Order Y/N
    sShSpot As String * 1     'Show on Log Y/N
    sShInv As String * 1  'Show on invoices
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    lAutoCode As Long    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iSourceID As Integer    'Remote User ID (avoid resending to sender)
    sShInsertion As String * 1  'Show on Insertion order
    sUnused As String * 19
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 5002   'Last two bytes after the comment must be 0
    sComment As String * 5004   'Last bytes after the comment must be 0
End Type

'Cxf key record layout- use LONGKEY0
'Type CXFKEY0
'    lCode As Long
'End Type

'Type CXFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

'********************************************************
'
'Copy Feed Dates file definition
'
'*********************************************************

'Cyf record layout
Type CYF
    lCifCode As Long    'Copy inventory
    iVefCode As Integer 'Vehicle code number
    sTimeZone As String * 3     'Time zone
    iFeedDate(0 To 1) As Integer 'Feed Date
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
    lRafCode As Long    'Region Code(zero unless TimeZone = "R")
    sSource As String * 1   'Blank=From Bulk Feed; S=From Station Feed: StartGuide; K=From Station Feed KenCast
    sAffOrigXMitChar As String * 1  'Original Letter specified when transmitted
    iAffOrigXMitDate(0 To 1) As Integer 'Original Start date taht card was sent as New
    sKCOrigXMitChar As String * 1  'KenCast: Original Letter specified when transmitted
    iKCOrigXMitDate(0 To 1) As Integer 'KenCast: Original Start date taht card was sent as New
    lCode As Long       'Internal code number
    sUnused As String * 2
End Type

'Cyf key record layout
Type CYFKEY0
    lCifCode As Long    'Copy inventory
    iVefCode As Integer 'Vehicle code number
    sSource As String * 1   'Blank=From Bulk Feed; S=From Station Feed
    sTimeZone As String * 3     'Time zone
    lRafCode As Long    'Region Code(zero unless TimeZone = "R")
End Type

'Cyf key record layout- use LONGKEY0
'Type CYFKEY0
'    lCode As Long
'End Type

'********************************************************
'
'Demo Estimate file definition
'
'*********************************************************

Type DEF
    lCode As Long                   'Autoincrement field
    iDnfCode As Integer             'DNF code value
    lPopulation As Long             'Population
    iStartDate(0 To 1) As Integer   'Start Date of new population
    lEstimatePct As Long            'xxxx.xxxx (4 place number- computed from Population/original Population with book)
    sUnused As String * 10
End Type

'Def key record layout- use LONGKEY0
'Type DEFKEY0
'    LCode As Integer
'End Type

Type DEFKEY1
    iDnfCode As Integer             'DNF code value
    iStartDate(0 To 1) As Integer   'Start Date of new population
End Type


'********************************************************
'
'Delivery link file definition
'
'*********************************************************

'Dlf record layout
Type DLF
    lCode As Long    'Internal code number (AutoInc)
    iVefCode As Integer 'Vehicle code number
    sAirDay As String * 1     '0=M-F, 6= Sa, 7=Su
    iAirTime(0 To 1) As Integer 'Air Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iFeedTime(0 To 1) As Integer 'Feed Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sZone As String * 3     'Time zone
    iEtfCode As Integer         'Event type code
    iEnfCode As Integer         'Event name code
    sProgCode As String * 5 'Program code #
    sCmmlSched As String * 1    'Y=Yes; N=No
    iMnfFeed As Integer     'Multi-name: Feed
    sBus As String * 5      'Bus
    sSchedule As String * 2 'Schedule
    iStartDate(0 To 1) As Integer    'Start (Effective) date
    iTermDate(0 To 1) As Integer   'Termination date
    iMnfSubFeed As Integer  'Feed Code
    sFed As String * 1  'Transmit event: Y= Yes; N=No
                        'Program events always set to N
End Type

'Dlf key record layout
Type DLFKEY0
    iVefCode As Integer 'Vehicle code number
    sAirDay As String * 1     '0=M-F, 6= Sa, 7=Su
    iStartDate(0 To 1) As Integer    'Start (Effective) date
    iAirTime(0 To 1) As Integer 'Air Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

'Dlf key record layout- use LONGKEY0
'Type DLFKEY1
'    lCode As Long    'Internal code number (AutoInc)
'End Type

'********************************************************
'
'Demo Research Name file definition
'
'*********************************************************

Type DNF
    iCode As Integer    'Internal code number for Demo Book name
    sBookName As String * 30    'Book Name
    iBookDate(0 To 1) As Integer    'Book Date
    iEnteredDate(0 To 1) As Integer 'Entered date
    iUrfCode As Integer         'Last user who imported book
    sType As String * 1 'I=Imported; C=Custom; M=Modelled from Import
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    iAutoCode As Integer    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    sForm As String * 1     'blank or 6=Sixteen buckets; 8= eighteen buckets (test for 8)
    iPopDnfCode As Integer  'Get population from this DnfCode instead of this book if defined
    iQualPopDnfCode As Integer  'Get population from this DnfCode instead of this book if defined
    sEstListenerOrUSA As String * 1 'L=Lister; U=USA (default is L)
    sExactTime As String * 1      ' Exact Time: Y or N.  If blank
                                  ' set. Rule if any drfInfoType is
                                  ' defined as V(Vehicle), then set
                                  ' to Y.  Used only in Act 1 import
    sSource As String * 1      ' Source: A=Act 1; R=RADAR; C=CSV
                               ' (was IMS) and M=Manual
    sUnused As String * 12
End Type

'Dnf key record layout- use INTKEY0
'Type DNFKEY0
'    iCode As Integer
'End Type

'Type DNFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

Type POPDNFINFO
    iDnfCode As Integer
    iPopDnfCode As Integer
    iQualPopDnfCode As Integer
End Type

'********************************************************
'
'Demo Plus Data file definition
'
'*********************************************************

Type DPF
    lCode As Long                   'Autoincrement field
    lDrfCode As Long                'Auto Code from DRF
    iMnfDemo As Integer             'Mnf Code for Demo (M18-34)
    iDnfCode As Integer             'DNF code value
    lDemo As Long                   '# of people in thousands
    lPop As Long                    'Population for plus demos, used when getting population for non-standard demos (base 16)
    sUnused As String * 5
End Type

'Dpf key record layout- use LONGKEY0
'Type DPFKEY0
'    LCode As Integer
'End Type

Type DPFKEY1
    lDrfCode As Long             'Auto Code from DNF
    iMnfDemo As Integer             'Mnf Code for Demo (M18-34)
End Type

Type DPFKEY2
    iDnfCode As Integer             'DNF code value
    iMnfDemo As Integer             'Mnf Code for Demo (M18-34)
End Type



'********************************************************
'
'Demo Research Data file definition
'
'*********************************************************

'Drf record layout
Type DRF
    lCode As Long                   'Autoincrement field
    iDnfCode As Integer             'Auto Code from DNF
    sDemoDataType As String * 1     'P=Population Demos; D=Imported Demos; M=Manually Defined Demos
    iMnfSocEco As Integer           'Mnf Code for Social/Ecomonical Group # (A1; A2,...L3)
    iVefCode As Integer             'Vehicle Code
    sInfoType As String * 1         'D=Daypart data; T=Time Data;V=Vehicle Data
    iRdfCode As Integer             'Daypart code if sInfoType = D and matching one found when importing
    sProgCode As String * 5         'Prog Codes for sInfoType = T (importing)
    iStartTime(0 To 1) As Integer   'Start Time of avail if sInfoType = T or Start Time of Daypart if sInfoType = D and match not found (RdfCode = 0)
    iEndTime(0 To 1) As Integer     'End Time of avail if sInfoType = T or End Time of Daypart if sInfoType = D and match not found
    sDay(0 To 6) As String * 1      'Monday thru Sunday (Y or N) if sInfoType = T or sInfoType = D and match not found (RdfCode = 0)
    iStartTime2(0 To 1) As Integer  'Start Time of Second Daypart time if sInfoType = D and match not found (RdfCode = 0)
    iEndTime2(0 To 1) As Integer    'End Time of Second Daypart time if sInfoType = D and match not found
    iQHIndex As Integer             'sInfoType = T; then 1=12m; 2=12:15am, 2=12:30am;....
    iCount As Integer               'sInfoType = D, then # of Broadcast; sInfoType = T, then # of stations
    sExStdDP As String * 1          'sInfoType = T, then Excluded from standard daypart (Y or N)
    sExRpt As String * 1            'sInfoType = T, then Excluded from report (Y or N)
    sDataType As String * 1         'A= per Broadcast; C=Cume
    'lDemo(1 To 18) As Long          'drfForm = blank or 6 then:
    lDemo(0 To 17) As Long          'drfForm = blank or 6 then:
                                    '# of people in thousands (1=M12-17; 2=M18-24; 3=M25-34; 4=M35-44; 5=M45-49
                                    '6=M50-54; 7=M55-64; 8=M65+; 9=W12-17;...;16=W65+)
                                    'drfForm = 8 then:
                                    '# of people in thousands (1=M12-17; 2=M18-20; 3=M21-24; 4=M25-34; 5=M35-44; 6=M45-49
                                    '7=M50-54; 8=M55-64; 9=M65+; 10=W12-17;...;18=W65+)
    iDemoChgDate(0 To 1) As Integer 'Date demo value changed (used for alerts)
    iDemoChgTime(0 To 1) As Integer 'Time demo value changed (used for alerts)
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    lAutoCode As Long    'Unique ID = Remote ID + AutoCode
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    lPopDrfCode As Long     'Population DrfCode values.  Set only for Act1 import 'Station with Population by Market'
                            'This will allow one book name for many import instead of one book name per station
    sForm As String * 1     'blank or 6=Sixteen buckets; 8= eighteen buckets (test for 8)
    sACTLineupCode        As String * 11     ' ACT 1 Lineup Code
    sACT1StoredTime       As String * 1      ' ACT 1 stored time (T or blank)
                                             ' Test for T
    sACT1StoredSpots      As String * 1      ' ACT 1 stored spots (S or blank)
                                             ' Test for S
    sACT1StoreClearPct    As String * 1      ' ACT 1 stored Clear Percentt(C or
                                             ' blank) Test for C
    sACT1DaypartFilter    As String * 1      ' ACT 1 stored daypart Filter (F or
                                             ' blank) Test for F
    sUnused               As String * 50
End Type

'Drf key record layout
Type DRFKEY0
    iDnfCode As Integer             'Auto Code from DNF
    sDemoDataType As String * 1     'P=Population Demos; D=Imported Demos; M=Manually Defined Demos
    iMnfSocEco As Integer           'Mnf Code for Social/Ecomonical Group # (A1; A2,...L3)
    iVefCode As Integer             'Vehicle Code
    sInfoType As String * 1         'D=Daypart data; T=Time Data
    iRdfCode As Integer             'Daypart code if sInfoType = D and matching one found when importing
End Type

'Drf key record layout
Type DRFKEY1
    iDnfCode As Integer             'Auto Code from DNF
    sDemoDataType As String * 1     'P=Population Demos; D=Imported Demos; M=Manually Defined Demos
    iMnfSocEco As Integer           'Mnf Code for Social/Ecomonical Group # (A1; A2,...L3)
    iVefCode As Integer             'Vehicle Code
    sInfoType As String * 1         'D=Daypart data; T=Time Data
    iStartTime(0 To 1) As Integer   'Start Time of avail if sInfoType = T or Start Time of Daypart if sInfoType = D and match not found (RdfCode = 0)
End Type

'Drf key record layout- use LONGKEY0
'Type DRFKEY2
'    lCode As Long
'End Type

'Type DRFKEY3- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type


'********************************************************
'
'Delete Stamp file definition
'
'*********************************************************

Type DSF
    lCode As Long    'Internal code number for Demo Book name
    sFileName As String * 3    'File Name
    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    iRemoteID As Integer    'Unique ID = Remote ID + AutoCode
    lAutoCode As Long    'Unique ID = Remote ID + AutoCode
    iSourceID As Integer    'Remote User ID (avoid resending to sender)
    lCntrNo As Long     'Contract Number
    sUnused As String * 20
End Type

'Dsf key record layout- use LONGKEY0
'Type DSFKEY0
'    lCode As Long
'End Type

'******************************************************************************
' ECF_Event_Chg_Date Record Definition
'
'******************************************************************************
Type ECF
    lCode                 As Long            ' Internal Reference to Event
                                             ' Change Date
    lGsfCode              As Long            ' Game Schedule reference
    iFromDate(0 To 1)     As Integer         ' From Date
    iToDate(0 To 1)       As Integer         ' Changed "To Date"
    sWebCleared           As String * 1      ' Web Spot removed (Y/N). Test for
                                             ' Y
    sMarketronCleared     As String * 1      ' Marketron Spots removed (Y/N).
                                             ' Test for Y
    sUnivisionCleared     As String * 1      ' Univision spots removed (Y/N).
                                             ' Test for Y
    iEnteredDate(0 To 1)  As Integer         ' Entered date
    iEnteredTime(0 To 1)  As Integer         ' Entered time
    iUrfCode              As Integer         ' User reference
    sUnused               As String * 20     ' Unused
End Type


'Ecf key record layout- use LONGKEY0
'Type ECFKEY0
'    lCode                 As Long
'End Type

Type ECFKEY1
    lGsfCode              As Long
End Type

Type ECFKEY2
    sWebCleared           As String * 1
End Type

Type ECFKEY3
    sMarketronCleared     As String * 1
End Type

Type ECFKEY4
    sUnivisionCleared     As String * 1
End Type



'******************************************************************************
' EFF_Extra_Fields Record Definition
'
'******************************************************************************
Type EFF
    lCode                 As Long            ' Auto Increment reference code
    sType                 As String * 1      ' Record Type (I=Invoice Export)
    sString(0 To 2)       As String * 20     ' String 1 (Type = I: Program Code); String 2 (Type = I: SMS Code)
    lLong(0 To 2)         As Long            ' Long 2
    sUnused               As String * 20
End Type


'Eff key record layout- use LONGKEY0
'Type EFFKEY0
'    lCode                 As Long
'End Type



'********************************************************
'
'Error log file definition
'
'*********************************************************

'Elf record layout
'Type ELF
'    iDate(0 To 1) As Integer  'Date of error
'    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    iUrfCode As Integer     'Last user who modified line
'    sMessage As String * 100'Error message
'End Type

'Elf key record layout
'Type ELFKEY0
'    iDate(0 To 1) As Integer  'Date of error
'    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'End Type

'********************************************************
'
'Event Name file definition
'
'*********************************************************

'Enf record layout
Type ENF
    iCode As Integer    'Internal code number for Event name
    iVefCode As Integer 'Vehicle code number
    iEtfCode As Integer 'Event code number
    sName As String * 30  'Identification name
    iMnfGenre As Integer    'Multi-name reference to Genre (0 except for program event types)
    sPgmSource As String * 3    'Program source
    sType(0 To 1)  As String * 3    'Primary/Secondary program type
    lCefCode As Long    'Comment code number
    sTimeForm As String * 1 'Log time format
    sLenForm As String * 1  'Log length format
    sPgmForm As String * 1  'Log program name format
    iMerge As Integer   'Merge code number
End Type

'Enf key record layout- use INTKEY0
'Type ENFKEY0
'    iCode As Integer
'End Type

'********************************************************
'
'Event Types file definition
'
'*********************************************************

'Etf record layout
Type ETF
    iCode As Integer    'Internal code number for Vehicle
    sType As String * 1 'Event type (1=program; 2=Contract Avail; 3=Open BB; 4=floating avail;
                        '5=Close BB; 6=Cmml Promo; 7=Feed Avail; 8=PSA; 9=Promo;
                        'a=Page skip; b=line space; c=line space 2; d=line space 3
    sName As String * 20  'Identification name
    sTimeForm As String * 1   'Log time format
    sLenForm As String * 1  'Log length format
    sPgmForm As String * 1  'Log program name format
    sInUse As String * 1    'Y=Type in use: show with event names and program events as allowed type
    iMerge As Integer   'Merge code number
End Type

'Etf key record layout- use INTKEY0
'Type ETFKEY0
'    iCode As Integer
'End Type

'********************************************************
'
'Game Header file definition
'
'*********************************************************
Type GHF
    lCode As Long            ' Auto Increment Code
    iVefCode As Integer         ' Vehicle Reference Code
    sSeasonName           As String * 20     ' Season Name
    iSeasonStartDate(0 To 1) As Integer      ' Season Start Date
    iSeasonEndDate(0 To 1) As Integer        ' Season End Date
    iNoGames As Integer         ' Number of Games
    sUnused As String * 10     ' Unused
End Type

'Ghf key record layout- use LONGKEY0
'Type GHFKEY0
'    lCode As Long
'End Type

Type GHFKEY1
    iVefCode As Integer         ' Vehicle Reference Code
End Type

Type SEASONINFO
    sKey As String * 6     'Start Date (6)
    sSeasonName As String * 20 ' Season Name
    lCode As Long            ' Auto Increment Code
End Type

'********************************************************
'
'Game Schedule file definition
'
'*********************************************************
Type GSF
    lCode As Long            ' Auto Reference Code
    lghfcode As Long            ' Game Header Reference Code
    iVefCode As Integer
    iGameNo As Integer         ' Game Number
    sFeedSource As String * 1      ' Feed Source: Home (H) or Visiting (V) or National (N)
    iLangMnfCode As Integer         ' Language MultiName Reference Code
    iVisitMnfCode As Integer         ' Visiting Team Name MultiName Reference Code
    iHomeMnfCode As Integer         ' Home team Name MultiName Reference Code
    lLvfCode As Long            ' Library Reference Code
    iAirDate(0 To 1) As Integer     ' Air Date of Game
    iAirTime(0 To 1) As Integer     ' Air Time of Game
    iAirVefCode As Integer         ' Air Vehicle Reference Code (used to
                                          ' preempt spots from airing vehicle).
    sGameStatus As String * 1      ' Game Date and Time status (T=Tentative;
                                          ' F=Firm; C=Cancelled)
    sLiveLogMerge As String * 1    'If Vehicle Option (vpfGenLog) set to A (Ask): L=Live Log+Preempt; M=Merge+Preempt
    sXDSProgCodeID As String * 8      ' XDS Program Code ID to be used if vffXDProgCodeID is defined as Event.
    sBus As String * 20      'Engineering Export Bus value for Sports so that Engineering Links are not required.  Allow multi-bus to be separated with a ; as a separator
    iSubtotal1MnfCode     As Integer         ' Event Subtotal 1 MnfCode
                                             ' reference
    iSubtotal2MnfCode     As Integer         ' Event Subtotal 2 MnfCode
                                             ' reference
    sUnused               As String * 6      ' Unused
End Type

'Gsf key record layout- use LONGKEY0
'Type GSFKEY0
'    lCode As Long
'End Type

Type GSFKEY1
    lghfcode As Long            ' Game Header Reference Code
    iGameNo As Integer         ' Game Number
End Type

Type GSFKEY2
    lghfcode As Long            ' Game Header Reference Code
    iAirDate(0 To 1) As Integer     ' Air Date of Game
    iAirTime(0 To 1) As Integer     ' Air Time of Game
End Type

Type GSFKEY3
    iVefCode As Integer
    iGameNo As Integer         ' Game Number
End Type

Type GSFKEY4
    iAirVefCode As Integer         ' Air Vehicle Reference Code (used to
    iAirDate(0 To 1) As Integer     ' Air Date of Game
End Type

Type GSFINFO
    sKey As String * 24     'Date (6) + Time (6) + Game No
    iGameNo As Integer
    lGameDate As Long
    sVisitName As String * 4
    sHomeName As String * 4
    sAffPost As String * 1
    sGameStatus As String * 1
    sLiveLogMerge As String * 1
    lGsfCode As Long
End Type


'******************************************************************************
' ief_IDC_Enfored Record Definition
'
'******************************************************************************
Type IEF
    lCode                 As Long            ' Auto Increment reference code
    lScheduleDate         As Long            ' IDC Schedule Date
    lTriggerID            As Long            ' Trigger ID
    lScheduleID           As Long            ' IDC Schedule ID
    sUnused               As String * 15     ' Unused
End Type


'Type IEFKEY0
'    lCode                 As Long
'End Type

'Type IEFKEY1
'    lScheduleDate         As Long
'End Type

'Type IEFKEY2
'    lScheduleID           As Long
'End Type



'********************************************************
'
'Inventory Header file definition
'
'*********************************************************

Type IHF
    iCode                 As Integer         ' Auto Increment Code
    lghfcode              As Long            ' Game Header Reference
    iVefCode              As Integer         ' Vehicle Reference
    iItfCode              As Integer         ' Inventory Type reference
    iIifCode              As Integer         ' Inventory Item reference
    sOversell             As String * 1      ' Allow items to be Over sold (Y or
                                             ' N)
    sGameIndependent      As String * 1      ' Game-Independent (Y/N).  If Y
                                             ' then only ask info without asking
                                             ' game number.  If No, then ask
                                             ' info by game
    sUnused               As String * 10     ' Unused
End Type

'Ihf key record layout- use INTKEY0
'Type IHFKEY0
'    iCode As Integer
'End Type


Type IHFKEY1
    lghfcode As Long            ' Game Header Reference
End Type


Type IHFKEY2
    iVefCode As Integer         ' Vehicle Reference Code
End Type


'********************************************************
'
'Inventory Item file definition
'
'*********************************************************

Type IIF
    iCode              As Integer         ' Auto Increment Code
    iItfCode           As Integer         ' Inventory Type reference
    sName              As String * 60     ' Inventory Item Name (i.e. Full Page)
                                          '
    sUnused            As String * 10     ' Unused
End Type

'Iif key record layout- use INTKEY0
'Type IIFKEY0
'    iCode As Integer
'End Type

Type IIFKEY1
    iItfCode           As Integer         ' Inventory Type reference
End Type


'********************************************************
'
'Inventory Specification file definition
'
'*********************************************************

Type ISF
    lCode              As Long            ' Auto Increment Code
    iIhfCode           As Integer         ' Inventory Header Reference
    lghfcode           As Long            ' Game Header reference
    iVefCode           As Integer         ' Vehicle Reference
    iGameNo            As Integer         ' Game Number
    iNoUnits           As Integer         ' Number of units (whole number)
    lCost              As Long            ' Cost of Item to the Network (xxx.xx)
                                          '
    lRate              As Long            ' Rate to charge client (advertiser) f
                                          ' or Item (xxxx.xx)
    iInvGameNo         As Integer         ' Inventory Game number that match Teams,
                                          ' Dates andTimes with this image
    sUnused            As String * 10     ' Unused
End Type

'Isf key record layout- use LONGKEY0
'Type ISFKEY0
'    lCode As Long
'End Type

Type ISFKEY1
    lghfcode           As Long            ' Game Header reference
    iGameNo            As Integer         ' Game Number
End Type

Type ISFKEY2
    iVefCode           As Integer         ' Vehicle Reference
    iGameNo            As Integer         ' Game Number
End Type

Type ISFKEY3
    iIhfCode           As Integer         ' Inventory Header Reference
    iGameNo            As Integer         ' Game Number
End Type

'********************************************************
'
'Inventory Type file definition
'
'*********************************************************

Type ITF
    iCode              As Integer         ' Auto Increment code
    sName              As String * 50     ' Inventory Type Name (i.e. Yearbook)
    sMultiFeed         As String * 1      ' Y=Yes; N=No.  Yes=Combine same games into one event
    sGLCash            As String * 20     ' G/L Cash: Replaces
                                             ' vpfBilledRevenue or Great Plains
                                             ' Primary Code Gross Sales
                                             ' (vafPCGrossSalesCash)
    sGLTrade           As String * 20     ' G/L Trade: Replaces
                                             ' vpfBilledRevenue or Great Plains
                                             ' Primary Code Gross Sales
                                             ' (vafPCGrossSalesTrade)
    sUnused            As String * 10     ' Unused
End Type

'Itf key record layout- use INTKEY0
'Type ITFKEY0
'    iCode As Integer
'End Type


'******************************************************************************
' AXF_Audio_XDS_Map Record Definition
'
'******************************************************************************
Type AXF
    lCode                 As Long            ' Audio Vault X-Digital Map Auto-Invrement internal code
    iAdfCode              As Integer         ' Advertiser internal reference code
    sAudioVaultID         As String * 20     ' Audio Vault Indicator ID
    sXDSCue               As String * 30     ' X-Digital Cue replacement if
                                             ' split copy
    sUnused               As String * 10     ' Unused
End Type


'Type AXFKEY0 key record layout- use LONGKEY0
'    lCode                 As Long
'End Type

'Type AXFKEY1 key record layout- use INTKEY0
'    iAdfCode              As Integer
'End Type

'********************************************************
'
'Log Calendar file definition
'
'*********************************************************

'Lcf record layout
Type LCF
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    iLogDate(0 To 1) As Integer 'Log Date
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
                                'TFN: iLogDate(0) = 1 (mon), 2 (tue),... 7 (sun) and iLogDate(1) = 0
    iSeqNo As Integer           'Sequency number (used to order log dates)
    iType As Integer            '0=Regulare Programming; 1->NN = Sports programming (Game number)
    sStatus As String * 1       'C=Current; P=Pending
    sTiming As String * 1       'Log timing indicator (N=Not started; I=Timing Incomplete; C=Timing Completed)
    sAffPost As String * 1      'Affidavit posting (N=No posting done yet; I=Posting Incomplete; C=Posting completed)
    iLastTime(0 To 1) As Integer 'Last Time Timed or posted (signoff time if completed)
                                '(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    'lLvfCode(1 To 50) As Long  'Log version library code
    lLvfCode(0 To 49) As Long  'Log version library code
    'iTime(0 To 1, 1 To 50) As Integer 'Log library start time
    iTime(0 To 1, 0 To 49) As Integer 'Log library start time
                                '(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iUrfCode As Integer         'Last user who modified Log library
    lCode As Long               ' Autoincrement field
End Type

Type LCFKEY0
    iType As Integer            '0=Regulare Programming; 1->NN = Sports programming (Game number)
    sStatus As String * 1       'C=Current; P=Pending; D=Deleted
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    iLogDate(0 To 1) As Integer 'Log Date
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
                                'TFN: iLogDate(0) = 1 (mon), 2 (tue),... 7 (sun) and iLogDate(1) = 0
    iSeqNo As Integer           'Sequency number (used to order log dates)
End Type

Type LCFKEY1
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    iType As Integer            '0=Regulare Programming; 1->NN = Game programming
End Type

Type LCFKEY2
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    iLogDate(0 To 1) As Integer 'Log Date
End Type

'Lcf key record layout- use LONGKEY0
'Type LCFKEY3
'    lCode As Long
'End Type

'******************************************************************************
' LDF_Live_Log_Detail Record Definition
'
'******************************************************************************
Type LDF
    lCode                 As Long            ' Autoincrement field
    lLlfCode              As Long            ' LLF Reference code
    sType                 As String * 1      ' Type: S=User Sign On/Off; P=Prog
                                             ' Start/End; G=Game Start/End;
                                             ' F=Feed Change
    sSubType              As String * 1      ' S=Start Time, E=End Time
    iTime(0 To 1)         As Integer         ' Time of type.  See subtype to
                                             ' determine which time
    iUrfCode              As Integer         ' User Reference code
    sCurrentFeed          As String * 1      ' Current Feed Source: Home(H) or
                                             ' Visiting(V) or National (N)
    lEnteredDate          As Long            ' Entered date
    lEnteredTime          As Long            ' Entered time
    sUnused               As String * 12     ' Unused
End Type


'Type LDFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type LDFKEY1
    lLlfCode              As Long
End Type

'********************************************************
'
'Library events file definition
'
'*********************************************************

'Lef record layout
Type LEF
    lLvfCode As Long            'Log library that event beyond withincode
    iStartTime(0 To 1) As Integer 'Start Time of event(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSeqNo As Integer           'Sequency number (used to order events at same time)
    iEtfCode As Integer         'Event type code
    iEnfCode As Integer         'Event name code
    iMnfExcl(0 To 1) As Integer 'Program Exclusion code
    ianfCode As Integer         'Avail name code to book into (Event type = 2,3,4,5,6,7,8)
    iMaxUnits As Integer        'Max units allowed to be book into avail
    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sTrue As String * 1         'True start time (Y=Yes-can't alter time in log; N=No)
    'lCifCode As Long            'Copy Inventory code
    lEvtIDCefCode As Long       'Event ID comment code number
    sPreFinal As String * 1     'P=Preliminary copy; F=Final copy
    lCefCode As Long            'Comment code number (initially from event names)
    iUrfCode As Integer         'Last user who modified Log library
    lCode As Long               ' Auto Increment
End Type

Type LEFKEY0
    lLvfCode As Long            'Log library that event beyond withincode
    iStartTime(0 To 1) As Integer 'Start Time of event(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSeqNo As Integer           'Sequency number (used to order events at same time)
End Type

'Type LEFKEY1
'    lLvfCode As Long            'Log library that event beyond withincode
'End Type

'Lef key record layout- use LONGKEY0
'Type LEFKEY2
'    lCode As Long
'End Type

'********************************************************
'
'Live Log file definition
'
'*********************************************************
Type LLF
    lCode              As Long            ' Auto Increment
    sName              As String * 40     ' Name of person that Sign on
    sStudioID          As String * 10     ' Studio ID
    iVefCode           As Integer         ' Vehicle Reference Code
    iAirDate(0 To 1)   As Integer         ' Air Date
    iGameNo            As Integer         ' Game Number or Zero if Program
    sSignOffStatus     As String * 1      ' Sign Off Status (Y=Sign off; N=Did n
                                          ' ot Sign Off)
    sReplaceStatus     As String * 1      ' Program or Game Replace Status (Y=Ye
                                          ' s; N=Not Replaced).
    lReplaceLlfCode    As Long            'Reference LLF that replaced this Sign-On
    sLogCompleted      As String * 1      'All Events Logged (Y/N)
    iLcfStartTime(0 To 1)   As Integer         'Library start time
    lEnteredDate       As Long            ' Entered date
    lEnteredTime       As Long            ' Entered time used to sort
                                          ' replacement records in order
    sUnused            As String * 16     ' Unused
End Type

'Llf key record layout- use LONGKEY0
'Type LLFKEY0
'    lCode As Long    'Internal code number for Copy Inventory
'End Type

Type LLFKEY1
    iAirDate(0 To 1)   As Integer         ' Air Date
    iVefCode           As Integer         ' Vehicle Reference Code
End Type

Type LLFKEY2
    lReplaceLlfCode    As Long            'Reference LLF that replaced this Sign-On
End Type


'********************************************************
'
'Log Title Library file definition
'
'*********************************************************

Type LTF
    iCode As Integer            'Log Title library Code
    sType As String * 1         'Type: R=Regular; S=Special; P=Sports; F=Std Format
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    sName As String * 20        'Log Library Name
    iVar As Integer             'Variation number
End Type

'Ltf key record layout- use INTKEY0
'Type LTFKEY0
'    iCode As Integer               'Log library Code
'End Type

'Type LTFKEY1
'    iVefCode           As Integer         ' Vehicle Reference Code
'End Type

Type LTFEXT
    iCode As Integer            'Log Title library Code
    sName As String * 20        'Log Library Name
    iVar As Integer             'Variation number
End Type
Public Const LTFEXTPK As String = "IB20I"

'********************************************************
'
'Log version Library file definition
'
'*********************************************************

'Lvf record layout
Type LVF
    lCode As Long               'Log version library Code
    iLtfCode As Integer         'Log title code
    iVersion As Integer         'Version number (0=Not used-> not in a hold day or air day)
    iBaseTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                                'If Hund sec = 1 and all other times equal 0, then relative time
    iLen(0 To 1) As Integer     'Length of library
    iUrfCode As Integer         'Last user who modified Log library
End Type

'Lvf key record layout- use LONGKEY0
'Type LVFKEY0
'    lCode As Long               'Log version library Code
'End Type

'Lvf key record layout
Type LVFKEY1
    iLtfCode As Integer         'Log title code
    iVersion As Integer         'Version number (0=Not used-> not in a hold day or air day)
End Type

'************************************************************
'
'Rvf sort by key
'
'*************************************************************

Type GP
    sKey As String * 40 'Division(10) Branch(10) Invoice # (10) Cash/Trade (1)
    bDelete As Boolean
    'iRvfPointer As Integer     'Dan M 8-05-08 changed from integer to long
    lRvfPointer As Long
End Type

'GetPaid Common Definitions
'8-4-10 moved from Collectionsubs.bas for a version of ZEroPurge in GetPaid export
Type ZEROPURGE
    lInvNo As Long
    iAgfCode As Integer
    iAdfCode As Integer
    sAmount As String * 20
    sType As String * 1
    sCheckNo As String * 10
    lFirstLk As Long
End Type

Type ZPLINK
    lNextLk As Long
    lRvfCode As Long
End Type

' FTP Operations
Declare Function csiFTPInit Lib "csi_utils.dll" (ByRef FTPInfo As CSIFTPINFO) As Integer
Declare Function csiFTPFileToServer Lib "csi_utils.dll" (ByVal slFileName$) As Integer
Declare Function csiFTPFileFromServer Lib "csi_utils.dll" (ByVal slFileName$) As Integer
Declare Function csiFTPGetStatus Lib "csi_utils.dll" (ByRef FTPStatus As CSIFTPSTATUS) As Integer
Declare Function csiFTPGetError Lib "csi_utils.dll" (ByRef FTPErrorInfo As CSIFTPERRORINFO) As Integer
Declare Function csiFTPGetFileListing Lib "csi_utils.dll" (ByRef FTPFileListing As CSIFTPFILELISTING) As Integer
' End FTP Operation

' FTP Operations
Type CSIFTPINFO
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sSendFolder As String * 128
   sRecvFolder As String * 128
   sServerDstFolder As String * 128
   sServerSrcFolder As String * 128
   sLogPathName As String * 128
End Type
 
Type CSIFTPSTATUS
   iState As Integer    ' 0=Complete, 1=Busy.
   iStatus As Integer   ' 0=Success, 1=Errors occured
   iJobCount As Integer ' The # of files yet to process.
   lLastError As Long   ' Contains the results of GetLastError if an error occurs.
End Type

Type CSIFTPERRORINFO
    sInfo As String * 1024
    sFileThatFailed As String * 128
End Type
 
' End FTP Operation - See delcares above
' FTP Operations
Type CSIFTPFILELISTING
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sPathFileMask As String * 128
   sSavePathFileName As String * 128
   sLogPathName As String * 128
   nTotalFiles As Integer
End Type

Public tgCsiFtpFileListing As CSIFTPFILELISTING
'******************************************************************************
' ibf_Impression_Bill Record Definition
'
'******************************************************************************
Type IBF
    lCode                 As Long            ' Impression bill auto code
                                             ' reference
    lCntrNo               As Long            ' Contract number
    iPodCPMID             As Integer         ' Podcast CPM ID
    iVefCode              As Integer         ' Vehicle reference code
    sBillCycle            As String * 1      ' S=Standard month bill; C=Calendar
                                             ' month bill
    iBillYear             As Integer         ' Billing year
    iBillMonth            As Integer         ' Billing month
    lImpressions          As Long            ' Impressions for the billing
                                             ' period
    sBilled               As String * 1      ' Billed flag: Y=Yes; N or Blank =
                                             ' No. Test for Y
    iUrfCode              As Integer         ' User reference
    sSource               As String * 1      ' M=Manual posting; A=Auto posting; I = Invoice created (Flat rate only and IBF does not exist)
    lBilledImpression     As Long            ' Retain number of impressions
                                             ' billed; Set as Part of Final
                                             ' Invoice and used during Reprint
    sUnused               As String * 16     ' Unused
End Type

'------IBF--------
Type IBFKEY0
    lCode                 As Long
End Type
Type IBFKEY1
    lCntrNo               As Long
    iPodCPMID             As Integer
End Type
Type IBFKEY2
    lCntrNo               As Long
    iVefCode              As Integer
End Type
Type IBFKEY3
    iBillYear             As Integer
    iBillMonth            As Integer
End Type
Type IBFKEY4
    lCntrNo               As Long
    iBillYear             As Integer
    iBillMonth            As Integer
End Type


'******************************************************************************
' avf_AdVendor Record Definition
'
'******************************************************************************
Type AVF
    iCode                 As Integer         ' Ad Server vendor auto code
                                             ' reference
    sName                 As String * 255    ' Vendor name
    sVendorAbbr           As String * 7      ' no longer used
    sVendorUserName       As String * 255    ' User name to log into API (ours or vendors).  Must be unique if not blank.
    sVendorPassword       As String * 255    ' Password to log into API (ours or vendors)
    sStatus               As String * 1      ' D=Dormant.  Test for D
    sUnused1              As String * 1
    sVendorURL            As String * 255    ' URL of vendor API
    sOrganizationID       As String * 70     ' Organization ID. UUID format
                                             ' Megaphone only
    sBucketName           As String * 63     ' Amazon Bucket Name - between 3
                                             ' (min) and 63 (max) characters
    sBucketFolder         As String * 80     ' Amazon Bucket Folder (optional,
                                             ' 63 (max) characters)
    sBucketRegion         As String * 16     ' Amazon Bucket Region. Example:
                                             ' USEast1, USEast2, USWest1 or
                                             ' USWest2
    sBucketAccessKey      As String * 20     ' Amazon Bucket Access Key
    sBucketPrivateKey     As String * 40     ' Amazon Bucket Private Key
    iLastImportDate(0 To 1) As Integer       ' Date of the last imported data
    sTimeZone             As String * 1      ' E=Eastern, C=Central, M=Mountain,
                                             ' P=Pacific, A=Alaska, H=Hawaiian A
                                             ' and H not currently allowed
    sUnused2              As String * 255
    sUnused3              As String * 255
    sUnused4              As String * 216
    iPriorityStart        As Integer         ' Priority Start number
    iPriorityEnd          As Integer         ' Priority End number
    sUnused               As String * 16     ' Unused
End Type


Type AVFKEY0
    iCode                 As Integer
End Type



'******************************************************************************
' Category Items Definition
' Created: 01/05/21       By: L.Bianchi
' Modified: 02/04/21      By: Luigi
'
'******************************************************************************
Type PodCategoryTHFHeaderInfo
    lCode                 As Long 'THF header code
    sName                 As String * 40     ' item name
    iCategoryCode      As Integer         ' Pod target category reference
    sCategoryName As String * 20  'Name of Category from MNF
End Type

Type RAFPodTarget
    lCode As Long 'raf region code
    iAdfCode As Integer
    sName As String * 80 'region name
    rType As String * 1   'raf type
    rState As String * 1 'raf state
    iUrfCode As Integer         ' User reference
    sChangeDate As String * 10 'date of raf change
    sDormantDate As String * 10 'Dormant Date
End Type

Type RAFPodTargetKEY0
    lCode                 As Integer
End Type

Type SEFPodTarget
    lCode As Long 'SEF code
    lRafCode As Long 'reference to Raf region code
    lThfCode As Long 'reference to thf header - Sef Long COde
    lMnfCode As Long ' category id
    iSequence As Integer 'sequence Number
    sCategoryName As String ' category name
    sPodItemName As String ' pod item name
    iPodItemIndex As Integer ' pod item index
End Type

Type PostCpmManual
    sKey As String * 80
    lContractNumber As Long
    iID As Integer 'pcfPodCPMID pcf_pod_cpm_Cntr
    iVehCode As Integer ' Vehicle Code
    sVehName As String 'Vehicle Name
    iAdfCode As Integer 'Advertiser Code
    sAdfName As String * 30 'Advertiser Name
    sChfBillCycle As String 'contract bill cycle
    iRdfCode As Integer 'rdfCode
    sAdLocation As String 'rdf name
    StartDate As Date 'pcfStartDate
    EndDate As Date 'pcfEndDate
    sPriceType As String * 1 'pcfPriceType
    lCPM As Long 'pcfPodCPM xxxx.xx
    lOrdered As Long 'pcfImpresssionGroals
    ilbfUnbilledCode As Long 'ibfCode when ibfBilled = N
    lBilled As Long 'Monthly ibfImpressions when ibfBilled = y
    lUnBilled As Long 'Monthly ibfImpressions when ibfBilled = N
    lTotalBilled As Long 'Sum of each month ibfImpression When ibfBilled = y
    lTotalInvoiced As Long 'Sum of each month ibfBilledImpression When ibfBilled = y
    lPcfTotalCost As Long ' pcfTotalCost
    lRvfTotalGross As Long
    lPhfTotalGross As Long
    phfTranDate As Date
    rvfTranDate As Date
    Length As Integer
End Type
'******************************************************************************
Type ACF
    lCode                 As Long
    lCntrNo               As Long
    sFileName             As String * 255
    sDescription          As String * 100
    iUrfCode              As Integer
    iEnteredDate(0 To 1)  As Integer
    iEnteredTime(0 To 1)  As Integer
    sDeleted              As String * 1      ' Y/N blank is NO
    iDeletedDate(0 To 1)  As Integer
    iDeletedTime(0 To 1)  As Integer
    sUnused               As String * 25
End Type


Type ACFKEY0
    lCntrNo               As Long
End Type
Type ACFList
    lCode                 As Long
    lCntrNo               As Long
    sFileName             As String
    sDescription          As String
    iUrfCode              As Integer
    sEnteredDate          As String
    sEnteredTime          As String
    bDeleted              As Boolean
    sDeletedDate          As String
    sDeletedTime          As String
    sTempPath             As String
    bChanged              As Boolean
End Type
