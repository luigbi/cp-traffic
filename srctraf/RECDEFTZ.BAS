Attribute VB_Name = "RECDEFTZ"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RECDEFTZ.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  VAFKEY1                                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RecDefTZ.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions M- Z
Option Explicit

'******************************************************************************
' TMF_Task_Monitor Record Definition
'
'******************************************************************************
Type TMF
    lCode                 As Long            ' Auto-increment
    sTaskCode             As String * 3      ' CSC=ContractSpotCreation;ASC=Affi
                                             ' liateSpotCreation;AEQ=AffiliateEx
                                             ' portQueue;ASI=AffidavitSpotImport
                                             ' ;ARQ=AffiliateReportQueue;ASQ=Ava
                                             ' ilSummaryGen;SC=SetCredit;PE=Proj
                                             ' ection;SFE=SalesForce;ME=Matrix;E
                                             ' PE=EfficioProj;ERE=EfficioRev;GPE
                                             ' =GetPaid;BD=BackupData
                                             ' TE = Tableau export
    sTaskName             As String * 30     ' Task Name
    sService              As String * 1      ' C=CSI_Service; T=Task
                                             ' Schedule;N=None
    sRunMode              As String * 1      ' P=Periodic; C=Continuous
    iRunningDate(0 To 1)  As Integer         ' Cintinuously Updated every X
                                             ' seconds. Used to determine if
                                             ' program is running
    iRunningTime(0 To 1)  As Integer         ' Cintinuously Updated every X
                                             ' seconds. Used to determine if
                                             ' program is running
    i1stStartRunDate(0 To 1) As Integer      ' 1st run start date for a day
    i1stStartRunTime(0 To 1) As Integer      ' 1st run time for a day
    i1stEndRunDate(0 To 1) As Integer        ' 1st run end date for a day
    i1stEndRunTime(0 To 1) As Integer        ' 1st run end time for a day
    iStartRunDate(0 To 1) As Integer         ' Last Start Date for a day task
                                             ' started to run. Periodic run
                                             ' mode. Used to determine if
                                             ' program is running
    iStartRunTime(0 To 1) As Integer         ' Last Start Time for a day task
                                             ' started to run. Periodic run
                                             ' mode. Used to determine if
                                             ' program is running
    iEndRunDate(0 To 1)   As Integer         ' Last End Date for a day task
                                             ' completed. Periodic run mode.
                                             ' Used to determine if program is
                                             ' running
    iEndRunTime(0 To 1)   As Integer         ' Last End Time for a day task
                                             ' completed. Periodic run mode.
                                             ' Used to determine if program is
                                             ' running
    sMo                   As String * 1      ' Run on Monday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sTu                   As String * 1      ' Run on Tuesday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sWe                   As String * 1      ' Run on Wednesday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sTh                   As String * 1      ' Run on Thursday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sFr                   As String * 1      ' Run on Friday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sSa                   As String * 1      ' Run on Saturday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sSu                   As String * 1      ' Run on Sunday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sMonthPeriod          As String * 2      ' SE=End Standard; EC=End Calendar;
                                             ' I=Invoice. Periodic Run Mode
    iDaysAfter            As Integer         ' Number of days after period that
                                             ' task is run. Periodic Run Mode
    iEMailReqDate(0 To 1) As Integer         ' Date detected that E-Mail needs
                                             ' to be sent. Wait W minutes before
                                             ' sending
    iEMailReqTime(0 To 1) As Integer         ' Time detected that E-Mail needs
                                             ' to be sent. Wait W minutes before
                                             ' sending
    iEMailSentDate(0 To 1) As Integer        ' E-Mail sent date
    sStatus               As String * 1      ' Status: S=Started; C=Completed; E=Error; A=Aborted
    sUnused               As String * 10     ' Unused
End Type

'Type TMFKEY0
'    lCode                 As Long
'End Type

Type TMFKEY1
    sTaskCode             As String * 3
End Type

'******************************************************************************
' TRF_Tax_Rate Record Definition
'
'******************************************************************************
Type TRF
    iCode                 As Integer         ' Auto Increment
    sTax1Name             As String * 30     ' Tax 1 rate name
    lTax1Rate             As Long            ' Tax 1 Rate (xx.xxxx)
    sTax2Name             As String * 30     ' Tax 2 rate name
    lTax2Rate             As Long            ' tax 2 rate (xx.xxxx)
    iUrfCode              As Integer         ' User reference
    sGrossNet             As String * 1      'Tax on Gross (G) or Net(N).  Defaukt USA to G and Canada to N.
    sUnused               As String * 9      ' Unused
End Type


'Trf key record layout- use INTKEY0
'Type TRFKEY0
'    iCode                 As Integer
'End Type

'********************************************************
'
'Time Zone Copy File Definition
'
'********************************************************

'Time zone copy record layout
Type TZF
    lCode As Long           'Internal code number for Time zone copy
    'lCifZone(1 To 6) As Long       'Code number for cif or ccf (if negative)
    'iRotNo(1 To 6) As Integer      'Rotation number
    'sZone(1 To 6) As String * 3    'Zone (ALL for all others) (if assign for EST only then EST for index 1 and ALL for PST, CST, MST as Index 2)
    lCifZone(0 To 5) As Long       'Code number for cif or ccf (if negative)
    iRotNo(0 To 5) As Integer      'Rotation number
    sZone(0 To 5) As String * 3    'Zone (ALL for all others) (if assign for EST only then EST for index 1 and ALL for PST, CST, MST as Index 2)
End Type

'Time zone copy key record layout- use LONGKEY0
'Type TZFKEY0
'    lCode As Long
'End Type

'******************************************************************************
' UAF_User_Activity Record Definition
'
'******************************************************************************
Type UAF
    lCode                 As Long            ' Auto Increment code
    sSystemType           As String * 1      ' System Type (T=Traffic;
                                             ' A=Affiliate)
    sSubType              As String * 1      ' Sub type (T=Task; R=Report)
    lUlfCode              As Long            ' User Log reference code
    iUserCode             As Integer         ' User Refernce code(urf if
                                             ' uafSystemType = T; or ust if
                                             ' uafSystemType = A)
    sName                 As String * 50     ' Task Name or Report name
    sStatus               As String * 1      ' Status (C=Complete; I=In
                                             ' Progress; A=Aborted)
    iStartDate(0 To 1)    As Integer         ' Start Date of Task or Report
    iStartTime(0 To 1)    As Integer         ' Start Time of Task or Report
    iEndDate(0 To 1)      As Integer         ' End date of Task or Report
                                             ' (12/31/2069 if not set)
    iEndTime(0 To 1)      As Integer         ' End Time of Task or Report
                                             ' (00:00:00 if not set)
    iCSIDate(0 To 1)      As Integer         ' CSI date
    sUnused               As String * 20     ' Unused
End Type


'User Activity key record layout- use LONGKEY0
'Type UAFKEY0
'    lCode                 As Long
'End Type

Type UAFKEY1
    sSystemType           As String * 1
    iUserCode             As Integer
End Type

Type UAFKEY2
    lUlfCode              As Long
    sStatus               As String * 1
    iStartDate(0 To 1)    As Integer
    iStartTime(0 To 1)    As Integer
End Type

Type UAFKEY3
    iStartDate(0 To 1)    As Integer
End Type

'******************************************************************************
' ULF_User_Log Record Definition
'
'******************************************************************************
Type ULF
    lCode                 As Long            ' Auto Increment code
    sDBType               As String * 1      ' Database Type: Production (P) or
                                             ' Test (T)
    sSystemType           As String * 1      ' System Type: Traffic(T) or
                                             ' Affiliate(A) or Engineering(E)
    iUrfCode              As Integer         ' Traffic User Reference code
    iUstCode              As Integer         ' Affiliate User reference code
    iUieCode              As Integer         ' Engineering User reference code
    iSignOnDate(0 To 1)   As Integer         ' Sign On Date
    iSignOnTime(0 To 1)   As Integer         ' Sign On Time
    iSignOffDate(0 To 1)  As Integer         ' Sign Off Date
    iSignOffTime(0 To 1)  As Integer         ' Sign off time
    sPCName               As String * 255    ' Computer name
    sPCMACAddr            As String * 20     ' Network card address (Unique by
                                             ' PC)
    sTimeZone             As String * 1      ' Time zone: Eastern(E);
                                             ' Central(C); Mountain(M);
                                             ' Pacific(P)
    iActiveLogDate(0 To 1)   As Integer         ' Updated hourly to be used to determine if user logged off incorrectly
    iActiveLogTime(0 To 1)   As Integer         ' Updated hourly to be used to determine if user logged off incorrectly
    iTrafJobNo            As Integer         'Traffic Job number
    iTrafListNo           As Integer         'Traffic List number
    iTrafRnfCode          As Integer         'Traffic report reference code
    iAffTaskNo            As Integer         'Affiliate Task number
    iAffSubtaskNo         As Integer         'Affiliate subtask number
    iAffRptNo             As Integer         'Affiliate report number
    'sUnused               As String * 12     ' Unused
End Type

'Use LONGKEY0
'Type ULFKEY0
'    lCode                 As Long
'End Type

Type ULFKEY1
    sDBType               As String * 1
    sSystemType           As String * 1
    iUrfCode              As Integer
    iUstCode              As Integer
    iUieCode              As Integer
End Type


'********************************************************
'
'User Option file definition
'
'*********************************************************

'Urf record layout
Type URF
    iCode                 As Integer         ' Internal code number for user
                                             ' Option
    sName                 As String * 20     ' User name
    iVefCode              As Integer         ' Vehicle code (0=All Vehicles)
    sRept                 As String * 20     ' User name to show on reports
    sPassword             As String * 20     ' Password
    iSlfCode              As Integer         ' Salesperson code
    sClnMoYr              As String * 1      ' M=Calendar by month, Y=Calendar
                                             ' by year
    sClnType              As String * 1      ' Calendar: S=Standard, C=Cal,
                                             ' U=Julian+, D=Julian -, C=Corp
    sClnLayout            As String * 1      ' Layout of calendar: A=Across,
                                             ' D=Down
    iClnLeft              As Integer         ' Left coordinate of calendar
    iClnTop               As Integer         ' Top coordinate of calendar
    iClcLeft              As Integer         ' Left coordinate of calculator
    iClcTop               As Integer         ' Top coordinate of calculator
    sWin(0 To 69)         As String * 1      ' Form (window) authorization
                                             ' (H=Hide; V=View only; I=input)
    iDefVeh               As Integer         'Default Vehicle (Vehicle code number)
    sGrid                 As String * 1      ' Todays grid level (H=Hide;
                                             ' V=View;  I=input)
    sPrice                As String * 1      ' Affidavit- spot price (H=Hide;
                                             ' V=View;  I=input)
    sCredit               As String * 1      ' Credit restriction in advt;
                                             ' agency (H=Hide; V=View;  I=input)
    sPayRate              As String * 1      ' Payment rating in Advt; Agency
                                             ' (H=hide; V=View;  I=input)
    sMerge                As String * 1      ' Merge button on screens (H=hide;
                                             ' V=View;  I=input)
    sDelete               As String * 1      ' Y=Yes; N=No
    sHideSpots            As String * 1      ' Allow spots to be hidden (H=hide;
                                             ' I=Input)
    sChgBilled            As String * 1      ' Allow billed spots to be
                                             ' changed(H=hide; I=input)
    sChgCntr              As String * 1      ' Allow contract weeks in pass with
                                             ' unbilled spots to be
                                             ' altered(H=hide;I=Input)
    sChgCrRt              As String * 1      ' Allow to change credit rating of
                                             ' advertisers and
                                             ' agencies(H=hide;I=Input)
    sBouChk               As String * 1      ' Get notice of bounced
                                             ' checks(H=hide;I=Input)
    sReprintLogAlert      As String * 1      ' Y=Show reprint log alerts;
                                             ' N=Don't
    sIncompAlert          As String * 1      ' Y=User gets alert when proposal
                                             ' changed to incomplete
    sCompAlert            As String * 1      ' Y=User gets alert when proposal
                                             ' changed to complete
    sSchAlert             As String * 1      ' Y=User gets alert when a order or
                                             ' hold required scheduling
    sHoldAlert            As String * 1      ' Y=Show "holds" on alert list box
    sRateCardAlert        As String * 1      ' Y=User gets alert when rate card
                                             ' changed that affects proposals
    sResearchAlert        As String * 1      ' Y=User gets an alert when new
                                             ' research data affects proposals
    sAvailAlert           As String * 1      ' Y=User gets an alert when
                                             ' insufficient avails exist for
                                             ' proposal
    sCrdChkAlert          As String * 1      ' Y=User gets an alert when credit
                                             ' approved for proposal
    sDeniedAlert          As String * 1      ' Y=User gets an alert when credit
                                             ' denied for proposal
    sCrdLimitAlert        As String * 1      ' Y=User gets an alert when credit
                                             ' limit exceeded for order/hold
    sMoveAlert            As String * 1      ' Y=User gets an alert when unschd
                                             ' order/hold has spots that will be
                                             ' altered in week prior to Last Log
                                             ' date
    iMnfHubCode           As Integer         ' Multi-Name Hub Code (mnfType = U)
    sChgLnBillPrice       As String * 1      '12-11-06 Allow cnt priced to be changed if billed
    sShowNRMsg            As String * 1      ' Show N-R messages (Y or N)
    sPrgmmaticAlert       As String * 1      ' Programmatic Buys (I/V/H); Test for I or V
    sAdvanceAvails        As String * 1      ' Advance Avail allowed
                                             ' checks(H=hide;I=Input). Test for I
    iDashOnStart       As Integer            ' 0=Don't show; 1=show Dash on
                                             ' startup
    sContractCreation     As String * 1      ' Y or Blank= user can create new
                                             ' contracts. N = user cannot
    sUnused2           As String * 7
    sWorkToDead           As String * 1      'Y=User Allowed to Change Working Proposal to Dead Proposal
    sWorkToComp           As String * 1      ' Y=User Allowed to Change Working
                                             ' Proposal to Complete Proposal
    sWorkToHold           As String * 1      ' Y=User Allowed to Change Working
                                             ' Proposal to Hold
    sWorkToOrder          As String * 1      ' Y=User Allowed to Change Working
                                             ' Proposal to Order
    sCompToIncomp         As String * 1      ' Y=User Allowed to Change Complete
                                             ' Proposal to Incomplete Proposal
    sCompToDead           As String * 1      ' Y=User Allowed to Change Complete
                                             ' Proposal to Dead Proposal
    sCompToHold           As String * 1      ' Y=User Allowed to Change Complete
                                             ' Proposal to Hold
    sCompToOrder          As String * 1      ' Y=User Allowed to Change Complete
                                             ' Proposal to Order
    sIncompToDead         As String * 1      ' Y=User Allowed to Change
                                             ' Incomplete Proposal to Dead
                                             ' Proposal
    sIncompToComp         As String * 1      ' Y=User Allowed to Change
                                             ' Incomplete Proposal to Complete
                                             ' Proposal
    sIncompToHold         As String * 1      ' Y=User Allowed to Change
                                             ' Incomplete Proposal to Hold
    sIncompToOrder        As String * 1      ' Y=User Allowed to Change
                                             ' Incomplete Proposal to Order
    sDeadToWork           As String * 1      ' Y=User Allowed to Change Dead
                                             ' Proposal to Working Proposal
    sHoldToOrder          As String * 1      ' Y=User Allowed to Change Hold to
                                             ' Order
    iPasswordDate(0 To 1) As Integer         ' Date that password was defined
    sChangeCSIDate        As String * 1
    sAllowInvDisplay      As String * 1
    sActivityLog          As String * 1      ' Activity Log: V=View; H=Hide
    sCntrVerify           As String * 1      ' Allowed to set the contract verification (I or H).  Test for I
    sChgAcq               As String * 1      ' Allowed to Change Acquisition Cost. I=Input; V=View. Test for I
    sAddAttach            As String * 1      ' Add attachments Y/N blank is no
    sRemoveAttach         As String * 1      ' Remove attachments Y/N blank is
                                             ' no
    sAdHocType            As String * 1      ' Ad Hoc Type: C=Create and Run;
                                             ' R=Run Only; N or Blank=Not
    sResvType             As String * 1      ' Reservation: V=View; H=Hide; I=Input
    sRemType              As String * 1      ' Remnant: V=View; H=Hide; I=Input
    sDRType               As String * 1      ' Direct Response: V=View; H=Hide;
                                             ' I=Input
    sPIType               As String * 1      ' Per Inquiry: V=View; H=Hide;
                                             ' I=Input
    sPSAType              As String * 1      ' PSA: V=View; H=Hide; I=Input
    sPromoType            As String * 1      ' Promo: V=View; H=Hide; I=Input
    sRefResvType          As String * 1      ' I=Allowed to reference
                                             ' reservation order
    iSnfCode              As Integer         ' Set Name Code
    sUseComputeCMC        As String * 1      ' I=Using Compute Button on the
                                             ' Contract Screen
    iGroupNo              As Integer         ' Remote Group No
    sReviseCntr           As String * 1      ' Allowed to revise Existing Holds
                                             ' or Orders (Y/N)
    sBlockRU              As String * 1      ' Block Remote User from
                                             ' exporting/importing(Y/N)
    sRCView               As String * 1      ' Rate Card view: D=Daypart; R =
                                             ' Rate; all other also=Rate
    sRegionCopy           As String * 1      ' Region Copy: V=View; H=Hide;
                                             ' I=Input
    sPDFDrvChar           As String * 1      ' 1st character of PDF Print Driver
                                             ' Name
    iPDFDnArrowCnt        As Integer         ' Number of down arrows to select
                                             ' PDF Driver from printer setup box
    sPrtDrvChar           As String * 1      ' 1st character of Default Print
                                             ' Driver Name
    iPrtDnArrowCnt        As Integer         ' Number of down arrows to select
                                             ' Default Driver from printer setup
                                             ' box
    sPrtNameAltKey        As String * 1      ' Alt Key for defining Printer Name
    iPrtNoEnterKeys       As Integer         ' Number of times the enter key
                                             ' must be entered to leave the
                                             ' Printer Setup screen
    sChgPrices            As String * 1      ' Allow contract prices to be
                                             ' changed (Y/N)
    sSpotFont             As String * 1      ' Spot Screen Font:S=Small; A=Arial
    sActFlightButton      As String * 1      ' Active Flight button
    sAvailSettings        As String * 1      ' Bit 0 (0=Wkly, 1=Daily); bit 1
                                             ' (0=Hourly; 1=Daypart); Bit 2-3
                                             ' (0=Sold; 1=Avails; 2=Inventory);
                                             ' Bit 4 (0=#; 1=%); Bit 5 (0=Unit;
                                             ' 1=Seconds)
    iRemoteUserID         As Integer         ' Remote user ID
    iRemoteID             As Integer         ' Remote ID (Note: Unique ID=Remote
                                             ' ID + AutoCode)
    iAutoCode             As Integer         ' Auto Incr Code (Note: Unique
                                             ' ID=Remote ID + AutoCode)
    iSyncDate(0 To 1)     As Integer         ' Sync Date (from Master Server or
                                             ' Remote System)
    iSyncTime(0 To 1)     As Integer         ' Sync Time (from Master Server or
                                             ' Remote System)
    sOldPassword(0 To 2)  As String * 20     ' Password used prior to current
                                             ' password (urfPassword)
    sLiveLogPostOnly As String * 1  'Live Log Posting Only (Y/N).  If yes then user can only see the live log posting screen.
    sSportPropOnly        As String * 1      ' Sport Proposal Enter only (Y/N).
                                             ' If yes then user can only enter
                                             ' proposals, enter inventory and
                                             ' run reports
    lEMailCefCode         As Long            ' E-Mail address
    sPhoneNo              As String * 25     ' Phone number
    sCity                 As String * 50     ' City
    sAllowedToBlock       As String * 1      ' Allowed to set Block on
                                             ' system(Y/N)
    sDigitalImpAlert      As String * 1      ' Alert setting Digital impression
                                             ' import notifications. Y=Enabled,
                                             ' N or Null = Disabled
    sDigitalCntrAlert     As String * 1      ' Alert setting Digital contract
                                             ' email notifications. Y=Enabled, N
                                             ' or Null = Disabled
    sUnused               As String * 100
End Type

'Urf key record layout- use INTKEY0
'Type URFKEY0
'    iCode As Integer
'End Type


Type URFEXT
    iCode As Integer        'Internal code number for user Option
    sName As String * 20    'User name
End Type
Public Const URFEXTPK As String = "IB20"


'******************************************************************************
' VAF_Vehicle_Account Record Definition
'
'******************************************************************************
Type VAF
    lCode                 As Long            ' Autoincrement
    iVefCode              As Integer         ' Vehicle Reference code
    sDivisionCode         As String * 10     ' Prefix to the Invoice # stored in
                                             ' G/L
    sBranchCodeCash       As String * 10     ' Branch code appended to Primary
                                             ' code
    sBranchCodeTrade      As String * 10     ' Branch code appended to Primary
                                             ' code
    sPCGrossSalesCash     As String * 10     ' Primary Code Gross Sales-Cash
                                             ' prefix to branch code
    sPCAgyCommCash        As String * 10     ' Primary Code Agency commision
                                             ' Cash  prefix to branch code
    sPCRecvCash           As String * 10     ' Primary code receivable cash
                                             ' prefix to branch code
    sPCGrossSalesTrade    As String * 10     ' Primary code gross sales trade
                                             ' prefix to branch code
    sPCRecvTrade          As String * 10     ' Primarycode receivable trade
                                             ' prefix to branch code
    sVendorID             As String * 15     ' Vendor reference ID for the vehicle
    sUnused               As String * 5      ' Unused
End Type


'Type VAFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type VAFKEY1 'VBC NR
    iVefCode              As Integer 'VBC NR
End Type 'VBC NR


'********************************************************
'
'Vehicle Conflict File Definition
'
'********************************************************

'Vehicle conflict record layout
Type VCF
    iSellCode As Integer    'selling vehicle code number
    iSellDay As Integer     '0=M-F, 5= Sa, 6=Su
    iSellTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSellPosNo As Integer  'Position (spot) no- currently zero
    'iCSV(1 To 5) As Integer 'Conflict selling vehicle code
    iCSV(0 To 4) As Integer 'Conflict selling vehicle code
    'sCSD(1 To 5) As String * 1  'Conflict selling day ("0"=M-F, "5" = Sa, "6" = Su)
    sCSD(0 To 4) As String * 1  'Conflict selling day ("0"=M-F, "5" = Sa, "6" = Su)
    'iCST(0 To 1, 1 To 5) As Integer    'Conflict selling avail time
    iCST(0 To 1, 0 To 4) As Integer    'Conflict selling avail time
    'iCSP(1 To 5) As Integer 'Position (spot) no
    iCSP(0 To 4) As Integer 'Position (spot) no
    iEffDate(0 To 1) As Integer    'Effective date
    iTermDate(0 To 1) As Integer   'Termination date
    sDelete As String * 1       'Delete flag (Y/N)
    lCode As Long           ' Autoincrement
End Type

'Vehicle conflict key record layout
Type VCFKEY0
    iSellCode As Integer    'selling vehicle code number
    iSellDay As Integer     '0=M-F, 5= Sa, 6=Su
    iEffDate(0 To 1) As Integer    'Effective date
    iSellTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSellPosNo As Integer   'Position (spot) no- currently zero
End Type

'Type VCFKEY1- use LONGKEY0
'    lCode                 As Long
'End Type

'********************************************************
'
'Vehicle File Definition
'
'********************************************************

'Vehicle record layout
Type VEF
    iCode As Integer        'Internal code number for Vehicle
    sName As String * 40    'Name
    sAddr(0 To 2)  As String * 25   'Address
    sPhone As String * 14   'Phone plus extension
    sFax As String * 10     'Fax number
    sUnused1 As String * 1   'Unused was OwnRep: O=Owned, R=Repped
    sDialPos As String * 5  'Dial Position
    lPvfCode As Long
    iReallDnfCode As Integer    'Reallocation Book Name Code
    'sUpdateRVF(1 To 8) As String * 1    'Participant-Update Rvf(Y) or Phf(N) flag
    sUpdateRVF(0 To 7) As String * 1    'Participant-Update Rvf(Y) or Phf(N) flag
    iCombineVefCode As Integer  'combine vefCode with vefCombineVefCode to generate log
    iMnfHubCode As Integer      'Multi-Name Hub Code (mnfType = U)
    'sUnused3 As String * 2  'Was: Format of Vehicle
    iTrfCode As Integer     'Tax rate reference
    sType As String * 1     'Type: C=Conventional; S=Selling; A=Airing; L=Log;V=Virtual; T=Simulcast;
                            '      R=Rep; N = NTR; P=Package; G=Sports (Games); I=Import
                            'Note: Conventional with Vpf.iGMnfNCode(1)>0 is
                            '      included with airing for creating delivery
                            ' The Log (L) vehicle is used to combine conventional vehicles
                            ' that reference the Log vehicles into one LOG.
    sCodeStn As String * 5  'Station vehicle code
    iVefCode As Integer     'Combination vehicle code (this code number is stored into
                            'ODF as the vehicle code if not zero)
    iOwnerMnfCode As Integer     'Owner set from first participant in latest Particpant table
    'iProdPct(1 To 8) As Integer  'If sOwnRep = R, then producer's %  (xx.xx)
    iProdPct(0 To 7) As Integer  'If sOwnRep = R, then producer's %  (xx.xx)
    sState As String * 1    'A=Active; D=Dormant
    'iMnfGroup(1 To 8) As Integer    'Participant (Vehicle Group) code number
    iMnfGroup(0 To 7) As Integer    'Participant (Vehicle Group) code number
    iSort As Integer        'Sort code
    iDnfCode As Integer     'Latest rating book code or default book
    iMnfDemo As Integer     'Prime Demo Code
    'iMnfSSCode(1 To 8) As Integer    'Sales Source code number
    iMnfSSCode(0 To 7) As Integer    'Sales Source code number
    sExportRAB As String * 1   'Export RAB: Y=Yes; N=No. Test for Y
    lVsfCode As Long        'Virtual vehicle or Standard Package Vehicle
    lRateAud As Long        'Rating(xx.xx) or Audience(xxxxxx). Stored as xxxxxx.xx
    lCPPCPM As Long         'CPP(xxxxx) or (CPM(xx.xx). Stored as xxxxx.xx)
    lYearAvails As Long     'Avails for Year (xxxxx)
    iPctSellout As Integer  '% sellout (xxx)
    iMnfVehGp2 As Integer     'Vehicle Group Set # 2
    iMnfVehGp3Mkt As Integer    'Market- Vehicle Group Set # 3
    iMnfVehGp4Fmt As Integer    'Format- Vehicle Group Set # 4
    iMnfVehGp5Rsch As Integer   'Research- Vehicle Group Set # 5
    iMnfVehGp6Sub As Integer       'SubCompany- Vehicle Group Set # 6
    iNrfCode              As Integer         ' Nrf reference.  if [X] N-R Rep,
                                             ' then Network Information; If [X]
                                             ' N-R Net, Then Rep Information
    iSSMnfCode As Integer       'Sales Source from Latest participant table owner row (1st row)
    sStdPrice As String * 1     'Standard Package Price (Distribute price to hidden lines: R= Rate; A= Audience; P= Percent; S= Spot Count)
    sStdInvTime As String * 1   'Standard Package Invoice Generate Time (A=Real-use hidden line times; O=Virtual- use Package line times; E= Price for package line)
    sStdAlter As String * 1     'Standard Package Alter Hidder Flag (Y=Allow Hidden Lines to be Altered; N= Only Package Line can be Altered; C=N except Comments)
    iStdIndex As Integer        'Standard Package Dollar index
    sStdAlterName As String * 1  'Standard Package Alter Hidder Flag (Y=Allow Hidden Lines to be Altered; N= Only Package Line can be Altered; C=N except Comments; R=N except Rate and Comments)
    iRemoteID As Integer        'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    iAutoCode As Integer        'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    'iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    'iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    'sExtUpdateRvf(1 To 8) As String * 1    'Participant-External Update Rvf(Y) or Phf(N) flag
    sExtUpdateRvf(0 To 7) As String * 1    'Participant-External Update Rvf(Y) or Phf(N) flag
    'iSourceID As Integer    'Remote User ID (avoid resending to sender)
    sStdSelCriteria As String * 1   'Standard package vehicle selection: A= All; M=Matching dayparts
    sStdOverrideFlag As String * 1  'Standard package override flag: S=Hidden line as subset; A=Set All overrides even if not subset
    sContact As String * 40 'Contact name  9/10/02
    lExtId                As Long         ' vendor's vehicle id Boostr Only
    sUnused               As String * 50
End Type

'Vehicle key record layout- use INTKEY0
'Type VEFKEY0
'    iCode As Integer
'End Type


'Vef key record layout
Type VEFKEY3
    sName As String * 40    'Name
End Type

'******************************************************************************
' VFF_Vehicle_Features Record Definition
'
'******************************************************************************
Type VFF
    iCode                 As Integer         ' Auto Increment.  Vehicle features
    iVefCode              As Integer         ' Vehicle reference
    sGroupName            As String * 10     ' Group Name
    sWegenerExportID      As String * 10     ' Wegener export vehicle ID
    sOLAExportID          As String * 7      ' OLA export vehicle ID
    sExportCustom         As String * 1      ' include vehicle in custom revenue
                                             ' export (Y/N)
    iLiveCompliantAdj     As Integer         ' Live pledge Plus/Minus window
                                             ' adjustment in minutes.
    iUstCode              As Integer         ' Affiliate User reference code
    iUrfCode              As Integer         ' Traffic user reference code
    sXDXMLForm            As String * 1      ' X-Digital XML Avail Cue Form: S= Avail ID by Hour and Break #, A=Avail ID by Hour #, Break # and Position; R=Replace Cue tag with Generic ISCI
    sXDISCIPrefix         As String * 6      ' X-Digital ISCI Prefix.  Required if two clients share the same X-Digital Head End
    sXDSaveCF             As String * 1      ' X-Digital File Delivery Save on
                                             ' Compact Flash Drive (Y/N). Test
                                             ' for Y
    sXDSaveHDD            As String * 1      ' X-Digital File Delivery Save on
                                             ' Hard Drive (Y/N). Test for Y
    sXDSaveNAS            As String * 1      ' X-Digital File Delivery Save on
                                             ' Netwoork Attached Storage device
                                             ' (External) (Y/N). Test for Y
    'sUnused               As String * 10     ' Unused
    sXDProgCodeID         As String * 8      'Program Code ID
    'sUnused               As String * 2      ' Unused
    iCwfCode              As Integer         'Combine web vehicle reference code
    sAirWavePrgID         As String * 10     ' AirWave Program ID
    sExportAirWave        As String * 1      ' Export via AirWave (Y/N). Default
                                             ' is N
    sExportNYESPN         As String * 1      ' Export NY ESPN (Y/N). Test for Y
    sPledgeVsAir As String * 1  'Include vehicle in the Pledge vs Air CSV Affiliate export (Y/N). Test for Y
    'sFedDelivery(1 To 5) As String * 1      ' Fed setting for zone 1 used when
    sFedDelivery(0 To 4) As String * 1      ' Fed setting for zone 1 used when
                                             ' generating Delivery Links
    iLastAffExptDate(0 To 1) As Integer     'Last Affiliate export date (independent of which export)
    sMoveSportToNon       As String * 1      ' Allow spots to be moved from
                                             ' Sport vehicle to non-sports
                                             ' vehicles (Y/N); Test for Y.
    sMoveNonToSport       As String * 1      ' Allow spots to be moved from
                                             ' non-sports vehicle to sports
                                             ' vehicles (Y/N); Test for Y.
    sMoveSportToSport     As String * 1      ' Allow spots to be moved from one
                                             ' Sport vehicle to another sports
                                             ' vehicles (Y/N); Test for Y.
    sMergeTraffic         As String * 1      ' In Traffic:  M=Merge into Log
                                             ' Vehicle; S=Separate from Log
                                             ' vehicle.  Test for S.
    sMergeAffiliate       As String * 1      ' In Affiliate: M=Merge into Log
                                             ' vehicle; S=Separate from Log
                                             ' vehicles. Test for S.
    sMergeWeb             As String * 1      ' In Affiliate Affidavit: M=Merge
                                             ' into Log vehicle; S=Separate from
                                             ' Log vehicles; C=Choose (Only one
                                             ' airs).  Test for S or C.
    sPledgeByEvent        As String * 1      ' Pledge defined by Event.  Y=Yes,
                                             ' N=No (or blank).  Test for Y.
    lPledgeHdVtfCode      As Long            ' Pledge Header RTF reference for
                                             ' the Sport Declaration report
    lPledgeFtVtfCode      As Long            ' Pledge Foot note RTF reference
                                             ' for the Sports Declaration
                                             ' report.
    iPledgeClearance      As Integer         ' Pledge clearance event count.
                                             ' This is the number of events that
                                             ' the Station must run
    sExportEncoESPN       As String * 1      ' Select vehicle when ENCO-ESPN
                                             ' automation export selected (Y/N).
                                             ' Test for Y.
    sWebName              As String * 40     ' Replacement name to appear on Web
                                             ' if defined for all vehicles. For
                                             ' Merged vehicles, show name within
                                             ' Log if defined.
    lSeasonGhfCode        As Long            ' Season Default selection Ghf
                                             ' reference.
    iMcfCode              As Integer         ' Mcf Reference Code.  Only show
                                             ' copy inventory for specified
                                             ' Media Definition.
    sExportAudio          As String * 1      ' Export Audio ISCI Titles (Y/N).
                                             ' Test for Y
    sExportMP2            As String * 1      ' Export MP2 (Y/N). Test for Y
    sExportCnCSpot        As String * 1      ' Export CnC Spots
                                             ' (Y/N). Test for Y
    sExportEnco           As String * 1      ' Export Enco (Y/N). Test for Y
    sExportCnCNetInv      As String * 1      ' Export CnC Network Inventory
                                             ' (Y/N). Test for Y
    sIPumpEventTypeOV     As String * 2      ' iPump Event Type ID override of
                                             ' the Media Code Event Type. If
                                             ' blank, then use Media Code Event
                                             ' Type
    sExportIPump          As String * 1      ' Export Wegener-iPump(Y/N). Test
                                             ' for Y
    sAddr4                As String * 25     ' Additional allowed address for
                                             ' vehicle (Line 4 of the address).
                                             ' Added here instead of in vef so
                                             ' that vef size would not be
                                             ' changed.
    lBBOpenCefCode        As Long            ' Open Billboard Comment displayed
                                             ' on BR if defined
    lBBCloseCefCode       As Long            ' Close Billboard Comment displayed
                                             ' on BR if defined
    lBBBothCefCode        As Long            ' Open and Close Billboard Comment.
                                             ' Not used at this time
    sXDSISCIPrefix        As String * 6      ' ISCI Prefix for the XDS Form by
                                             ' ISCI
    sXDSSaveCF            As String * 1      ' X-Digital File Delivery Save on
                                             ' Compact Flash Drive (Y/N) for the
                                             ' XDS Form by ISCI
    sXDSSaveHDD           As String * 1      ' X-Digital File Delivery Save on
                                             ' Hard Drive (Y/N) for the XDS Form
                                             ' by ISCI
    sXDSSaveNAS           As String * 1      ' X-Digital File Delivery Save on
                                             ' Netwoork Attached Storage device
                                             ' (External) (Y/N) for the XDS Form
                                             ' by ISCI
    sMGsOnWeb             As String * 1      ' Allow MGs on Web (Y/N). Test for Y
    sReplacementOnWeb     As String * 1      ' Allow Replacement on Web (Y/N). Test for Y
    sExportMatrix         As String * 1      ' Export Matrix (Y/N). Test for Y
    sSentToXDSStatus      As String * 1      ' Vehicle information sent to
                                             ' X-Digital (Y=Yes, N= No;
                                             ' M=Modiifed)
    sStationComp          As String * 1      ' Station Compensation (Y/N). Test for Y
    sExportSalesForce     As String * 1      ' Export Sales Force (Y/N). Test for Y. No Vehicle selectivity in Sales force, therefore not defined in Vehicle Option Export
    sExportEfficio        As String * 1      ' Export Efficio (Y/N). Test for Y. Efficio not created yet and don't know if any vehicle selectivity will be required.
    sExportJelli          As String * 1      ' Export Jelli (Y/N). Test for Y
    sOnXMLInsertion       As String * 1      ' Include on Insertion order for
                                             ' vendor (W=Wide Orbit;
                                             ' M=Marketron,...,N or Blank = No).
                                             ' Test for vendor letter
    sOnInsertions         As String * 1      ' On Insertion Orders. Y=Yes; N=No. Test for Y
    sPostLogSource        As String * 1      ' Post Log Station Source Times. N=None; S=Station Invoice
    sExportTableau        As String * 1      ' Export Tableau (Y/N). Test for Y
    sStationPassword      As String * 10     ' Station Password only asked when Station can post their invoiced spots. Used to allow stations to sign into Traffic and post spot times only
    sHonorZeroUnits       As String * 1      ' For airing vehicle honor zero
                                             ' unit when generating the Log and
                                             ' creating Pledge avails (Y/N).
                                             ' Test for Y
    sHideCommOnLog        As String * 1      ' Hide comments on Logs (Y/N). Test
                                             ' for Y
    sHideCommOnWeb        As String * 1      ' Hide comments on Electron
                                             ' Affidavit System (Web) (Y/N).
                                             ' Test for Y
    iConflictWinLen       As Integer         ' Conflick Window Length (0=60 seconds). Used when creating the vcf records
    sACT1LineupCode       As String * 11     ' Default vehicle ACT1 Lineup--Historic 7 16 21
                                             ' Code--Deprecated 7 16 21
    sPrgmmaticAllow       As String * 1      ' Programmatic buy Allowed (Y/N). Test for Y
    sSalesBrochure        As String * 20     ' Sales Brochure.  Path defined in Traffic.ini (SalesBrochure = ----)
    sCartOnWeb            As String * 1      ' Show Crat # on Wedb Log (Y/N, Default N and Blank = No. Test for Y)
    sDefaultAudioType     As String * 1      ' Default Aidio Type (R=Rec Cmml;
                                             ' L=Live Cmml; M=Live Promo; S=Rec
                                             ' Promo; P=Pre-Rec Live Cmml;
                                             ' Q=Pre-Rec Live Promo; Treact
                                             ' Blank same as R)
    iLogExptArfCode       As Integer         ' Reference to arf_Addresses file.  In the Arf file the Log Export Drive/Path will be stored
    'sUnused               As String * 6      ' Unused
    sASICallLetters       As String * 4      ' Audio Service Interface call
                                             ' letters (Used for vCreative)
    sASIBand              As String * 2      ' Audio Service Interface band
                                             ' (used for vCreative)
    iAvfCode              As Integer         ' Ad Server vendor reference code
    lAdVehNameCefCode     As Long            ' No Longer Used was Ad Server vendor vehicle name
                                             ' reference code
    sXDEventZone       As String * 1         ' cue code by zone Y or N(Blank)
    sUnused            As String * 19        ' Unused
End Type


'Type VFFKEY0
'    iCode                 As Integer
'End Type

'Type VFFKEY1
'    iVefCode              As Integer
'End Type
Type VFFKEY2
    lAdVehNameCefCode     As Long
End Type
'********************************************************
'
'Vehicle Link File Definition
'
'********************************************************

'Vehicle link record layout
Type VLF
    lCode As Long    'Internal code number (AutoInc)
    iSellCode As Integer    'selling vehicle code number
    iSellDay As Integer     '0=M-F, 6= Sa, 7=Su
    iSellTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSellPosNo As Integer   'Unit (spot) no- currently zero
    iSellSeq As Integer     'Sequence number start at 1
    sStatus As String * 1   'C=Current; P=Pending
    iAirCode As Integer    'airing vehicle code number
    iAirDay As Integer     '0=M-F, 5= Sa, 6=Su
    iAirTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iAirPosNo As Integer   'Unit (spot) no- currently zero
    iAirSeq As Integer     'Sequence number start at 1
    iEffDate(0 To 1) As Integer    'Effective date
    iTermDate(0 To 1) As Integer   'Termination date
    sDelete As String * 1       'Delete flag (Y/N)
End Type

'Vehicle link key record layout
Type VLFKEY0
    iSellCode As Integer    'selling vehicle code number
    iSellDay As Integer     '0=M-F, 6= Sa, 7=Su
    iEffDate(0 To 1) As Integer    'Effective date
    iSellTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSellPosNo As Integer   'Position (spot) no- currently zero
    iSellSeq As Integer     'Sequence number start at 1
End Type

Type VLFKEY1
    iAirCode As Integer    'airing vehicle code number
    iAirDay As Integer     '0=M-F, 6= Sa, 7=Su
    iEffDate(0 To 1) As Integer    'Effective date
    iAirTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iAirPosNo As Integer   'Position (spot) no- currently zero
    iAirSeq As Integer     'Sequence number start at 1
End Type

'Vlf key record layout- use LONGKEY0
'Type VLFKEY2
'    lCode As Long    'Internal code number (AutoInc)
'End Type

Type VLFKEY3
    iSellCode As Integer    'selling vehicle code number
    iSellDay As Integer     '0=M-F, 6= Sa, 7=Su
    sStatus As String * 1   'C=Current; P=Pending
End Type

Type VLFKEY4
    iAirCode As Integer    'airing vehicle code number
    iAirDay As Integer     '0=M-F, 6= Sa, 7=Su
    sStatus As String * 1   'C=Current; P=Pending
End Type

'********************************************************
'
'Vehicle Option file definition
'
'*********************************************************

'Vpf record layout
Type VPF
    iVefKCode As Integer    'Internal code number for Vehicle Option
    iGTime(0 To 1)  As Integer  'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sGMedium As String * 1  'Medium (T=TV; R=Radio; C=Cable; N=Radio Net; V=TV Network; S=Radio ROS Net, P = Podcast)
    iUrfGCode As Integer    'User code
    sAdvtSep As String * 1  'Advertiser separation by B=Break; T=Time
    sCPLogo As String * 3      'Certificate of Perforance Logo File Name (Gxx.Bmp)
    'iGLink(1 To 5) As Integer   'Unused: Other vehicles that this vehicle is linked with (selling to airing;
    '                            ' or Airing to selling)
    lLgHd1CefCode As Long       'Header comment 1 from vehicle options
    lLgNmCefCode As Long     'Vehicle Log Name from vehicle options
    sUnsoldBlank As String * 1  'If blackout replace is missing: Y=Show as Unsold Blank line; N=Remove Event (spot).  Test for N
    'sFeedLogOrder As String * 1    'Feed: A= Log (it was L but changed to A so that the user would not have to set all vehicles); I=Insertion Order; N=Neither
    sUsingFeatures1 As String * 1   'Bit Map (Right to Left): Bit 0=Export Insertion; Bit 1=Export Log; Bit 2=Import Insertion; Bit 3=Import Affiliate Spot;
                                    'Bit 4=Export ISCI by Pledge; 5=Unused as Old db might have field set on(option never set); 6=Suppress Web Log; 7=Export Posted Tiem to Web
    iSAGroupNo As Integer
    'sGPriceStat As String * 1   'Calculate spot price statistics on schedule lines (Y=Yes, N=No)
    sOwnership As String * 1   'A=Owned-Network; B=Owned-Station; C=Unowned-Network; D=Unowned-Station.  Defaul A or Blank. Test for B, C and D
    sGGridRes As String * 1     'Default grid resolution (F=Full hour, H=Half hour, Q=Quarter hour)
    'sGScript As String * 1   'Using scripts (Y=Yes, N=No)
    sExpHiCorp As String * 1     'Export Corporate ( or N; default is Y)
    'iGLocalAdj(1 To 5) As Integer   'Local time adjustment
    iGLocalAdj(0 To 4) As Integer   'Local time adjustment
    'iGFeedAdj(1 To 5) As Integer   'Feed time adjustment
    iGFeedAdj(0 To 4) As Integer   'Feed time adjustment
    'sGZone(1 To 5) As String * 3 'Time zone
    sGZone(0 To 4) As String * 3 'Time zone
    'iGV2Z(1 To 5) As Integer   'Versions displacement in minutes
    iGV2Z(0 To 4) As Integer   'Versions displacement in minutes
    'iGV3Z(1 To 5) As Integer   'Versions displacement in minutes
    iGV3Z(0 To 4) As Integer   'Versions displacement in minutes
    'iGV4Z(1 To 5) As Integer   'Versions displacement in minutes
    iGV4Z(0 To 4) As Integer   'Versions displacement in minutes
    'sGCSVer(1 To 5) As String * 1   'Show on cmml schd report:O=Original only; A=All version
    sGCSVer(0 To 4) As String * 1   'Show on cmml schd report:O=Original only; A=All version
    'iGMnfNCode(1 To 5) As Integer   'Feed codes (test this field if conventional to determine
    iGMnfNCode(0 To 4) As Integer   'Feed codes (test this field if conventional to determine
                                    'type of conventional (Index 1 only: with:field>0 or without delivery:field is zero)
    'sGBus(1 To 5) As String * 2  'Bus code
    sGBus(0 To 4) As String * 2  'Bus code
    'sGSked(1 To 5) As String * 2   'Schedule
    sGSked(0 To 4) As String * 2   'Schedule
    sSVarComm As String * 1  'Agency commission variable (Y= Yes; N=No)
    sSCompType As String * 1    'Separate competitives (T=Time; B=Break; N=not back to back)
    iSCompLen(0 To 1) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSBBLen As Integer          'BB spot length
    sSSellOut As String * 1     'Sellout defined (U=Units; B=Both; T= 30 second units; M=Matching Units)
    sSOverBook As String * 1    'Allow overbooking of avails (Y=Yes; N=No)
    sSForceMG As String * 1     'Set spots moved outside contract limits as MG's (W=Always; A=Ask)
    '7/15/14
    sEmbeddedOrROS As String * 1    'Default Delivery for agreements:E= Embedded or R = ROS.  Test for E. Blank should be treated same as R (replaced vpfSPlaceNet)
    sWegenerExport As String * 1     'Wegener Export(Y=Yes, N=No). Test for Y
    sOLAExport As String * 1     'OLA (OnLine Affidavit) Export (Y=Yes, N=No). Test for Y
    'sSIntoRC As String * 1      'Unused-'Insert all programs into Rate cards automatically (Y=Yes, N=No)
    sAvailNameOnWeb As String * 1      'Show Avail Names on Web (Y = yes; N=No)
    sUsingFeatures2 As String * 1         'Bit Mat (Right to Left): Bit 0=XDS Apply "Merge" ProgID; Bit 1=Invoice Show Air Date as W/O;Bit 2=Duplicate Rate Card By Base
    sSAvailOrder As String * 1  'Order spots within break (1=Media codes with clustered groups; 2= Media code with separated groups;
                                '3=Lengths with clustered groups; 4=Lengths with separated groups;
                                '5=Announcer codes with clustered groups; 6= Announcer codes with separated groups; 7= N/A)
                                'With clustered group=>place all same group #'s together and in # order (all 1 first, then 2's, 3's,..)
                                'With separated goups=>place all groups in number order (1,2,3,..1,2,3,.. 1,2,3...)
    iSLen(0 To 9)  As Integer   'Valid spot lengths
    iSLenGroup(0 To 9)  As Integer  'Group # for lengths
    sSCommCalc As String * 1    'Method of calculating salesperson commission (B=on billing; C=on collections)
    lMPSA60 As Long          'Index to last PSA 60 filler schedule line scheduled
    lMPSA30 As Long          'Index to last PSA 30 filler schedule line scheduled
    lMPSA10 As Long          'Index to last PSA 10 filler schedule line scheduled
    lMPromo60 As Long        'Index to last Promo 60 filler schedule line scheduled
    lMPromo30 As Long        'Index to last Promo 30 filler schedule line scheduled
    lMPromo10 As Long        'Index to last Promo 10 filler schedule line scheduled
    'iMMFPSA(1 To 24) As Integer 'M-F: Max PSAs allowed per hour
    iMMFPSA(0 To 23) As Integer 'M-F: Max PSAs allowed per hour
    'iMSAPSA(1 To 24) As Integer 'Sa: Max PSAs allowed per hour
    iMSAPSA(0 To 23) As Integer 'Sa: Max PSAs allowed per hour
    'iMSUPSA(1 To 24) As Integer 'Su: Max PSAs allowed per hour
    iMSUPSA(0 To 23) As Integer 'Su: Max PSAs allowed per hour
    'iMMFPromo(1 To 24) As Integer 'M-F: Max Promos allowed per hour
    iMMFPromo(0 To 23) As Integer 'M-F: Max Promos allowed per hour
    'iMSAPromo(1 To 24) As Integer 'Sa: Max Promos allowed per hour
    iMSAPromo(0 To 23) As Integer 'Sa: Max Promos allowed per hour
    'iMSUPromo(1 To 24) As Integer 'Su: Max Promos allowed per hour
    iMSUPromo(0 To 23) As Integer 'Su: Max Promos allowed per hour
    iLLD(0 To 1) As Integer     'Last Log Date Byte 0:Day, 1:Month, followed by 2 byte year
    iLPD(0 To 1) As Integer     'Last Preliminary Log Date Byte 0:Day, 1:Month, followed by 2 byte year
    slTimeZone As String * 1    'Time zone (E=Eastern; C=Central; M=Mountain; P=Pacific)
    sLDayLight As String * 1    'Daylight savings (Y=Yes; N=No)
    sLTiming As String * 1      'Using Log Timing (Y=Yes; N=No)
    sLAvailLen As String * 1    'Show length on unsold avails (Y=Yes; N=No)
    iSDLen As Integer           'Default spot length
    'sLCP As String * 1          'Unused-'Create cmml promo avails (Y/N)
    'sLBBs As String * 1         'Unused-'Create bb's (Y/N)
    iFTPArfCode As Integer      'FTP address stored into ARF (Audio Stored address so that affiliate can get the Audio)
    sLShowCut As String * 1      'Cut/instruction #'s (C=show cut only; I=show instruction only; B=show both; N=show neither)
    sLTimeFormat As String * 1  'Show time (A=AM/PM; M=Military)
    slZone As String * 1        'Log/C of P/Play List Zone (A=All; E=EST; C=CST; M=MST; P=PST)
    sCPTitle As String * 1      'Certificate of Performance Title required(Y/N)
    sPrtCPStation As String * 1 'Print Certificate of Performance by Station (Y/N)
    iRnfPlayCode As Integer  'Default Play List (rnf code)
    iLastCP(0 To 1) As Integer  'Last CP date (Affiliate)
    sStnFdCart As String * 1    'Show Carts on Station Feed (Y/N)
    sStnFdXRef As String * 1    'Include in Vehicle cross reference on Station Feed (Y/N)
    sGenLog As String * 1         'Generate Logs (Y=Print Log on separate pages; N=None; L=Live Log and Print; M=Merge Sports Log into Preempt vehicle log; A=Ask if Live+Preempt or Merge+Preempt)
    sCopyOnAir As String * 1    'Allow copy definition for Airing Vehicles (Y/N)
    sBillSA As String * 1   'Bill Airing Spots (Y/N)
                            'This field is used for partial simulcast vehicles
    sExpVehNo As String * 2     'Export Vehicle Number
    sExpBkCpyCart As String * 1      'Y=Show Cart # on Bulk Feed Export
    sExpHiCmmlChg As String * 1      'Export Commercial Change: Y=Highlight vehicle; N=Don't Highlight but show
    lAPenny As Long       'Highest penny variance amount (xxx.xx)
    'iGV1Z(1 To 5) As Integer   'Versions displacement in minutes
    iGV1Z(0 To 4) As Integer   'Versions displacement in minutes
    'sGFed(1 To 5) As String * 1   'Transmit: N=No; Y=Yes
    sGFed(0 To 4) As String * 1   'Transmit: N=No; Y=Yes
    sGGroupNo As String * 1     'Group number- Used to group Vehicles together in bulk feed
    iLLastDateCpyAsgn(0 To 1)  As Integer   'Last Date Copy Assigned Byte 0:Day, 1:Month, followed by 2 byte year
    'iESTEndTime(1 To 5) As Integer  'Daypart End Time of EST- in minutes
    iESTEndTime(0 To 4) As Integer  'Daypart End Time of EST- in minutes
    'iCSTEndTime(1 To 5) As Integer  'Daypart End Time of CST- in minutes
    iCSTEndTime(0 To 4) As Integer  'Daypart End Time of CST- in minutes
    'iMSTEndTime(1 To 5) As Integer  'Daypart End Time of MST- in minutes
    iMSTEndTime(0 To 4) As Integer  'Daypart End Time of MST- in minutes
    'iPSTEndTime(1 To 5) As Integer  'Daypart End Time of PST- in minutes
    iPSTEndTime(0 To 4) As Integer  'Daypart End Time of PST- in minutes
    'sMapZone(1 To 4) As String * 3  'Zone which program code is to be remapped into Daypart
    sMapZone(0 To 3) As String * 3  'Zone which program code is to be remapped into Daypart
    'sMapProgCode(1 To 4) As String * 5  'Program code to be mapped into different daypart
    sMapProgCode(0 To 3) As String * 5  'Program code to be mapped into different daypart
    'iMapDPNo(1 To 4) As Integer     'Daypart number that program code is to be mapped into
    iMapDPNo(0 To 3) As Integer     'Daypart number that program code is to be mapped into
    sExpHiClear As String * 1   'Export Clearance Spot: Y=Highlight vehicle; N=Don't Highlight
    sExpHiDallas As String * 1  'Export Dallas Feed: Y=Highlight vehicle; N=Don't Highlight but show
    sExpHiPhoenix As String * 1 'Export Phoenix Feed: Y=Highlight vehicle; N=Don't Highlight but show
    sExpHiNY As String * 1      'Export New York Feed: Y=Highlight vehicle; N=Don't Highlight but show
    sBulkXFer As String * 1     'Export Bulk Feed Cross Reference: Y=Include vehicle; N=bypass vehicle
    sClearAsSell As String * 1  'In Export Clearance Spots treat vehicle as Selling (Y or N)
    sClearChgTime As String * 1 'In Export Clearance Spots change 11:59:55pm to 12:00:01AM (Y or N)
    sMoveLLD As String * 1  'Y=Allowed to move spots between todays date and Last Log date
    iRnfLogCode As Integer  'Default Log Name (rnf code)
    iRnfCertCode As Integer 'Default Log Certification Name (rnf code)
    iLNoDaysCycle As Integer    'Number of days in closing cycle (i.e. 1, 7 ,14)
    iLLeadTime As Integer   'Closing lead time (number of days required before closing that log should be generated, excluses sa, su)
    sShowTime As String * 1 'On invoice show air time as D=Daypart; S=Spot Time; A=Avail Time
    sEDICallLetters As String * 4   'EDI Call Letters
    sAccruedRevenue As String * 20  'Accrued Revenue G/L #
    sAccruedTrade As String * 20  'Accrued Trade G/L #
    sBilledRevenue As String * 20  'Billed Revenue G/L #
    sBilledTrade As String * 20  'Billed Trade G/L #
    sLCmmlSmmyAvNm As String * 1 'Commercial Summary Aval Filter (first letter of avail name or blank for all)
    'iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
    'iSyncTime(0 To 1) As Integer  'Sync Time (from Master Server or Remote System)
    lEDASWindow As Long         'EDAS Time window in seconds
    sKCGenRot As String * 1     'KenCast: Generate Rotation (Y/N).  If no, then create the rotatopn but don't include it in the Envelope export
    sExportSQL As String * 1      ' Allow export to SQL server (Y/N)
    sAllowSplitCopy       As String * 1      ' Allow split copy (Y/N).  Only if
                                             ' Site Set Using Split Copy
    sExpHiNYISCI As String * 1    'Initially Highlight Vehicle with NY Export with ISCI (Y/N)
    iLastLog(0 To 1) As Integer   'Last CP date (Affiliate)
    iRnfSvLogCode As Integer    'Save to File Log Report Code
    iRnfSvCertCode As Integer    'Save to File C.P. Report Code
    iRnfSvPlayCode As Integer    'Save to File Other Report Code
    lLgFt1CefCode As Long       'Footer comment 1 from vehicle options
    lLgFt2CefCode As Long       'Footer comment 2 from vehicle options
    sStnFdCode As String * 2    'Station feed code- this must be unique for each station
    iProducerArfCode As Integer     'Produce ArfCode reference
    iProgProvArfCode As Integer   'Program Content Provider ArfCode Reference
    iCommProvArfCode As Integer   'Commercial Content Provider ArfCode reference
    sEmbeddedComm As String * 1   'Export Commercial Audio (Yes/No); Y = Export Comm Aud for Selected reference (not to stations); N=Export Comm Aud for Stations
    sARBCode As String * 6          'Arbitron code for vehicle
    lEMailCefCode         As Long            ' Station E-Mail Address
                                             ' Orders (Y/N)
    sShowRateOnInsert As String * 1 'Show Spot Rate on Insertion Orders (Y/N)
    iAutoExptArfCode As Integer   'Automation Export Path
    iAutoImptArfCode As Integer   'Automation Import Path
    sWebLogSummary As String * 1    'Show web log spot summary (Y/N)
    sWebLogFeedTime As String * 1   'Show on web log spot feed time (Y/N)
    sRadarCode As String * 5        'Radar Import codes
    sEDIBand As String * 1          'Frequency Band (A or F)
    iInterfaceID As Integer         'Interface ID (Used as ID for vehicle)
End Type

'Vpf key record layout
Type VPFKEY0
    iVefKCode As Integer
End Type

'Type VPFKEY1- use POPDATETYPE
'    iSyncDate(0 To 1) As Integer  'Sync Date (from Master Server or Remote System)
'End Type

'********************************************************
'
'Vehicle/Salesperson Combo file definition
'
'*********************************************************

'Vsf record layout
Type VSF
    lCode As Long        'Internal code number for user Option
    sType As String * 1     'Record Type: F=Vehicle combo (created for contract header); S=Salesperson; V=Virtual Vehicle
    sName As String * 20    'Combo name
    lLkVsfCode As Long      'Link to Next in Chain of VSF records (0=No More)
    sUnused1 As String * 16 'Market name
    iFSCode(0 To 49) As Integer 'Vehicle or Salesperson code 49
    lFSComm(0 To 49) As Long   'Revenue Allocation xxx Or Commission ( xx.xxxx)
    iNoSpots(0 To 49) As Integer   'Number of spots
    sSource As String * 1   'U=User defined; S=System defined
    iMerge As Integer       'Merge code number
End Type

'Vsf key record layout- use LONGKEY0
'Type VSFKEY0
'    lCode As Long
'End Type

'******************************************************************************
' VTF_Vehicle_Text Record Definition
'
'******************************************************************************
Type VTF
    lCode                 As Long            ' vehicle text reference auto
                                             ' increment
    iVefCode              As Integer         ' Vehicle reference code
    sType                 As String * 1      ' Type. H=Sports Declaration
                                             ' Header; F=Sports Declaration Foot note
                                             ' 1=Event Title 1; 2=Event Title 2
    sText                 As String * 5004   ' Text.  vtfType = H and F will be
                                             ' in RTF format.
End Type

'Vtf key record layout- use LONGKEY0
'Type VTFKEY0
'    lCode                 As Long
'End Type

Type VTFKEY1
    iVefCode              As Integer
    sType                 As String * 1
End Type


'********************************************************
'
'One Line Help file definition
'
'*********************************************************

'Hlf record layout
Type HLF
    sIDName As String * 3   'Help ID name
    iMsgNo As Integer       'Message number
    sFieldName As String * 10   'Field name
    lContextID As Long      'Context help ID number
    sMessage As String * 100    'One line help message
End Type

'Hlf key record layout
Type HLFKEY0
    sIDName As String * 3   'Help ID name
    iMsgNo As Integer       'Message number
End Type

'Agency and advertiser collection extract record
Type AGFADFCOLL
    iCode As Integer
    sName As String * 30
    sBillAgyDir As String * 1
    sCityID As String * 9   'City ID if agency(5 bytes), AddrID if Advertiser (9 bytes)
    iSlfCode As Integer 'Salesperson code number
    sCreditRestr As String * 1  'N=No Restricions; L=Credit Limit; W=Cash in Advance each Week;
                                'M=Cash in Advance each Month; T=Cash in Advance Quarterly;
                                'P=Prohibit New Orders
    'sCreditLimit As String * 5  'Credit limit amount (xxxxxxx.xx)
    lCreditLimit As Long  'Credit limit amount (xxxxxxx.xx)
    sPaymRating As String * 1   'O=Ok; S=Slow pay; K=Keep anEye On; C=in Collection
    'sPct90 As String * 2    '% Over 90 days (xxx)
    iPct90 As Integer    '% Over 90 days (xxx)
    sCurrAR As String * 6   'Current A/R (xx,xxx,xxx.xx)
    iDateLstPaym(0 To 1) As Integer 'Date of last payment
    iAvgToPay As Integer    'Average number of days to pay
End Type
Public Const AGFADFCOLLPK As String = "IB30BB9IBLBIB6III"
'********************************************************
'
'Program Dates/times definition
'
'*********************************************************

'Program dates/time record layout
'This record is used to communicate between program/events and program dates
Type PRGDATE
    sStartTime As String    'Start time
    iDay(0 To 6) As Integer 'Days 0=No; 1=Yes
    sStartDate As String    'Start Date
    sEndDate As String      'End date
End Type

'Library and events for library/Layout/Count display
Type LLC
    iDay As Integer         'Day index 1-7
    sType As String * 1     'Library (R,P, or S) or Event type (1, 2,..D or Y)
    sStartTime As String    'Start time of library or event
    sLength As String       'Length of library or event
    iUnits As Integer       'Units if avail; MnfExcl # 1 if program
    sName As String         'Library name (version#/Name-Variation)
    lLvfCode As Long        'Lvf Code (used by gBuildEventDay only)
    iLtfCode As Integer     'Ltf Code if avail ; MnfExcl # 1 if program (used by gBuildEventDay only)
    iAvailInfo As Integer   'Avail flags (used by gBuildEventDay only); iAnfCode (used by LinksDef->mReadLcfLefLnf & mMoveEvents)
    iEtfCode As Integer
    iEnfCode As Integer
    lCefCode As Long
    lEvtIDCefCode As Long
    sVisitName As String * 4
    sHomeName As String * 4
    sXMid As String * 1     'Used by prog and others as availinfo is used with avails
End Type

'Library corrdinates used for dragging libraries
Type LCD
    sName As String * 20    'Library name
    iDay As Integer         'Library day (1= Mo; 2=Tu;...7=Su)
    iCurOrPend As Integer   '0=Current; 1=Pending
    iLLCIndex As Integer    'Index for LLC
    fX1 As Single           'Left X coordinate within picture area
    fX2 As Single           'Right X coordinate within picture area
    fY1 As Single           'Top Y coordinate within picture area
    fY2 As Single           'Bottom Y coordinate within picture area
End Type

'********************************************************
'
'Schedule information definition
'
'*********************************************************

'Vehicle schedule status and association
Type VEFSCH
    sOnAirSchStatus As String * 1 'P=Pending; C=Current; I=Pending but links are incomplete
    sAltSchStatus As String * 1 'P=Pending; C=Current; I=Pending but links are incomplete
    iDay As Integer         '0=M-F; 6=Sa; 7=Su (set to -1 in gBuildVefInfo, set to value in gBuildVcf)
    iVefCode As Integer     'Vehicle Code
    sType As String * 1     'Type: C=Conventional; S=Selling; A=Airing; D=Delivery
    iGroup As Integer       'Group number- all vehicles with same group number belong together
    iTermDate0(0 To 1) As Integer 'Terminate of vehicle for Vlf and Vcf: Mo-Fr
    iTermDate6(0 To 1) As Integer 'Terminate of vehicle for Vlf and Vcf: Sa
    iTermDate7(0 To 1) As Integer 'Terminate of vehicle for Vlf and Vcf: Sunday
End Type

'********************************************************
'
'Spot function information
'
'*********************************************************

'Event table
Type EVTINFO
    iType As Integer    'Event type (1=Program; 2-9=Avail; 100=Spot, 101=Conflict spot, 99=Remaining units/sec for avail)
                        'Note: Spots adds a iType of -1 as a place holder to synch times across days
    lTime As Long       'Start time(Type = 1 - 9); Sdf Rec # (Type = 100 and 101)
    lLen As Long        'Length of event(Type = 1-9); lChfCode (Type = 100 and 101)
    iUnits As Integer   'LtfCode (type=1) or number of units defined (type = 2-9) or Length (type = 100 and 101)
    lInfo As Long       'LvfCode (type = 1) or AnfCode (type = 2-9) Or AdfCode (type = 100 and 101)
    iLineInfo As Integer    'Bit 0  1=Pending state (type = 100 and 101) (H1)
                            'Bit 1  1=MG Spot (H2)
                            'Bit 2  1=Outside spot (H4)
                            'Bit 3  1=Virtual Vehicle (H8)
                            'Bit 4  1=Extra spot (billable or non-billable) (H10)
                            'Bit 5  1=PSA/Promo (H20)
                            'Bit 6  1=Not in original vehicle (H40)
                            'Bit 7  1=Reservation contract (H80)
                            'Bit 8  1=Comment (H100)
                            'Bit 9  1=M for N spot (H200)
                            'Bit 10 1=Reservation Contract (H400)
                            'Bit 11 1=Primary Split Network spot (H800)
                            'Bit 12 1=Secondary Split Network spot (H1000)
                            'Bit 13 1=OPen/Close BB (H2000)
                            'Bit 14 1-Impression contract (H4000)
    lChfCode As Long    'ChfCode
    lFsfCode As Long    'FSFCode
    iLineNo As Integer  'Type 100 and 101
    iMnfComp1 As Integer    'MnfExcl0 (Type = 1) or MnfComp1 (Type 100 and 101)
    iMnfComp2 As Integer    'MnfExcl1 (Type = 1) or MnfComp2 (Type 100 and 101)
    lSsfRecPos As Long  'Ssf record position
    iSsfIndex As Integer 'Ssf index
    lSdfCode As Long    'Auto code from sdf
    sCntrType As String * 1 'Contract type (chfType)
    sSpot As String * 20 'Spot time\line number (type = 100 and 101)
    sShow As String * 80   'Display string
    sPrice As String * 12   'Price without pennies
    sAirDate As String * 10 'Air Datee if different from scheduled date
    sAirTime As String * 10 'Air time if different from schedule time
    lchfcxfCode As Long 'Header comment is show on spots
    lClfCxfCode As Long 'Line comment is show on spots
    lRafCode As Long    'Region
    lchfcxfInt As Long  'Internal comment
    sLiveCopy As String * 1
    sNetRegionAbbr As String * 5
    iRotNo As Integer
    sPtType As String * 1
    lCopyCode As Long
    sCITFlag As String * 1  'C=Cart; I=ISCI; T=creative Title
    sCopyCIT As String * 30    'Cart # or ISCI or Creative title.  sCITFlag indicates which value
    sPrgName As String * 20
End Type
'Table of values:
' Item      Program     Avail       Spot        Unsold  Conflict
'iType      1           2->9        100         99      101
'lTime      Start time  Start time  Sdf Rec#    -       Sdf Rec#
'                                   Not Code#           Not Code#
'lLen       Length      Length      lChfCode    Length  lChfCode
'iUnits     iLtfCode    Units       Spot Len    Units   Spot Len
'lInfo      lLvtCode    iAnfCode    iAdfCode    -       iAdfCode
'iLineInfo  -           iAvInfo     Bit 0=Pend  -       Bit 0=Pend
'                       (w/o units) Bit 1=MG            Bit 1=MG
'                                   Bit 2=Outside       Bit 2=Outside
'iLineNo    -           -           Line #      -       Line #
'iMnfComp1  Prog Excl   -           Competitive -       Competitive
'iMnfComp2  Prog Excl   -           Competitive -       Competitive
'lSsfRecPos lSsfRecPos  lSsfRecPos  lSsfRecPos  -       lSsfRecPos
'lSsfIndex  lSsfIndex   lSsfIndex   lSsfIndex   -       lSsfIndex
'lSdfCode   -           -           Sdf Code    -       Sdf Code
'sCntrType  -           -           Cntr Type   -       Cntr Type
'sSpot      -           -           Time/Line#  -       -
'sShow      Start time  Time        Length      -       Length
'           End Time    Units       Advt                Advt
'           Prog        Seconds     Product             Product
'                       Avail type  Comp1               Comp1
'                                   Comp2               Comp2
'                                   Rpf Name            Rpf Name

'********************************************************
'
'Delivery link function information
'
'*********************************************************

'Event table
Type DELEVT
    lTime As Long       'Feed time
    iEtfCode As Integer 'Event type
    iEnfCode As Integer 'Event name
    sProgCode As String * 5
    iStatus As Integer  '0=unused; 1=used
End Type
'******************************************************************************
' thf_Target_Header Record Definition
'
'******************************************************************************
Type THF
    lCode                 As Long            ' Geo-Target item definitions
                                             ' reference code
    sName                 As String * 40     ' item name
    iCategoryMnfCode      As Integer         ' Pod target category reference
    iUrfCode              As Integer         ' User reference code
    sUnused               As String * 20     ' Unused
End Type


Type THFKEY0
    lCode                 As Long
End Type

Type THFKEY1
    iCategoryMnfCode      As Integer
End Type


'******************************************************************************
' tif_Target_Items Record Definition
'
'******************************************************************************
Type TIF
    lCode                 As Long            ' Target item definition reference
                                             ' code
    iThfCode              As Integer         ' Target header reference code
    iAvfCode              As Integer         ' Ad server vendor reference code
    sAPICode              As String * 255    ' API Code to be sent to vendor
    iUrfCode              As Integer         ' User reference code
    sUnused               As String * 20     ' Unused
End Type


Type TIFKEY0
    lCode                 As Long
End Type

Type TIFKEY1
    iThfCode              As Integer
End Type


'******************************************************************************
' VAC_Vendor_ADF Record Definition
'
'******************************************************************************
Type VAC
    iCode                 As Integer
    iAvfCode              As Integer         ' Vendor Reference ID
    iAdfCode              As Integer         ' Advertiser Reference ID
    sExternalID           As String * 70     ' Vendor assigned Code for this
                                             ' Advertiser / UUID
    sUnused               As String * 100
End Type


Type VACKEY0
    iCode                 As Integer
End Type
Type VACKEY1
    iAvfCode              As Integer
    iAdfCode              As Integer
End Type

Type VACKEY2
    iAdfCode              As Integer
End Type

'******************************************************************************
' VVC_Vendor_Vef Record Definition
'
'******************************************************************************
Type VVC
    iCode              As Integer
    iAvfCode              As Integer         ' Vendor Reference ID
    iVefCode              As Integer         ' Vehicle Reference ID
    sExternalID           As String * 70     ' Vendor assigned Code for this
                                             ' Vehicle / UUID
    sUnused               As String * 100
End Type


Type VVCKEY0
    ivccCode              As Integer
End Type
Type VVCKEY1
    iAvfCode              As Integer
    iVefCode              As Integer
End Type

Type VVCKEY2
    iVefCode              As Integer
End Type

'******************************************************************************
' VCC_Vendor_Ctr Record Definition
'
'******************************************************************************
Type VCC
    lCode                 As Long
    iAvfCode              As Integer         ' Vendor Reference ID
    lCntrNo               As Long            ' Contract Reference ID
    sExternalID           As String * 70     ' Vendor assigned Code for this
                                             ' Contract / UUID
    sUnused               As String * 100
End Type


Type VCCKEY0
    lCode                 As Long
End Type
Type VCCKEY1
    iAvfCode              As Integer
    lCntrNo               As Long
End Type

Type VCCKEY2
    lCntrNo               As Long
End Type
