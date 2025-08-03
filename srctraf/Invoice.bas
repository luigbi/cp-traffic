Attribute VB_Name = "INVOICESubs"

' Copyright 1993 Counterpoint Software ® All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Invoice.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Initialize subs and functions
Option Explicit
Option Compare Text

Public igUnpackDateError As Integer
'10016
Public bgPDFEmailTestMode As Boolean

Type EDISERVICEINFO
    tArf As ARF
    hEDI As Integer
    sPathFile As String * 250
    sUsed As String * 1
    '4/3/12
    iTotalNoInv As Integer
    sTotalAired As String * 20
End Type

Type CLUSTERERR
    lCntrNo As Long 'Contract #
    lChfCode As Long
    iFlag As Integer    'Error type: 1=National Order- Local Salesperson; 2=National Salesperson- Local Order #
End Type
Public tgClusterErr() As CLUSTERERR

Type REPMATCH
    iVefCode As Integer
    lPrice As Long
    iONoSpots As Integer    'Ordered number of spots
    iLineNo As Integer
    iPkLineNo As Integer
    sPostBy As String * 1   'C=Count; L=Line
End Type

Type SBFCNTR
    sKey As String * 80      'same 51 bytes as in typesortkey + contract code + Invoice # (8- for reprint) + Date (5)
    tSbf As SBF
    lSDate As Long          'Start date of billing period
    iCarry As Integer       'Number of missed spots from previous month
    iANoSpots As Integer    'Ordered number of spots
    lAGross As Long
    lNTRGross As Long   'Gross amount for NTR
    lNTRNet As Long
    lInvNo As Long      'Invoice number for reprint
    lInvDate As Long    'Invoice date
    lAcquisitionCost As Long
    lTax1 As Long
    lTax2 As Long
    sCashTrade As String 'L.Bianchi 07/07/2021
End Type

Public tgSbfCntr() As SBFCNTR   'Used for NTR and Rep
Public tgSbfInstall() As SBFCNTR    'Used for installment

Type CPMCNTR
    sKey As String * 80      'same 51 bytes as in typesortkey + contract code + Invoice # (8- for reprint) + Date (5)
    tPcf As PCF
    tIbf As IBF
    lSDate As Long          'Start date of billing period
    lEDate As Long
    lAGross As Long
    lCPMGross As Long   'Gross amount for NTR
    lCPMNet As Long
    lInvNo As Long      'Invoice number for reprint
    lInvDate As Long    'Invoice date
    lAcquisitionCost As Long
    lTax1 As Long
    lTax2 As Long
End Type

Public tgCPMCntr() As CPMCNTR   'Used for Podcast CPM buys
Public tmCPMCntStatus() As INVCNTRSTATUS
Public imCPMStatusConflict As Integer
Public tmInvAirNTRStatus() As INVAIRNTRSTATUS

Type READCLF
    sKey As String * 4
    iLine As Integer
    iCntRevNo  As Integer
    lRecPos As Long
End Type
'Public igShowHelpMsg As Integer
Public igUsedISR As Integer 'Used ISR (True/Flase).  This is set true
                            'only if more then 32000 records required
Type SPEEDSORTPARSE
     sVehName As String * 40
     sSdfDate As String * 5
     sSdfTime As String * 6
End Type
Type SORTSDFEXT
    sKey As String * 20
    tSdfExt As SDFEXT
End Type
Type TYPESORTNOKEY
    'sKey As String * 245    'Requires 244
                            'MnfSort #(5); Agy/City(40+5+1) or Dir Advt Name;
                            '| Advt Name (30);
                            '| Contract #(8);
                            '| If Contract by Prod(35) use Copy Product Name else Contract Product
                            '| Ordered Vehicle Name(40), if bonus special character added to force sort to end;
                            '| Line Number(4);
                            '| Week #(4);
                            '| 0 if not Extra bonus; 1 if Extra bonus
                            '| Spot Vehicle Name(40) except for InvAirOrder=O or S, then its Line Vehicle Name;
                            '| Game No (4)
                            '| Spot Date(5)
                            '| Spot Time(6)
                            '| Package Line # or zero if not package (4)
    lChfCode As Long    'Internal code number for Contract
    sInvGp As String * 1        'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    iType As Integer    '0=Sdf; 1=Smf (for missed of MG not showing); 2=Psf
    iLineNo As Integer
    iBilled As Integer  'True=Spot billed; False=Spot not billed
    sPriceType As String * 1
    lCode As Long     'SdfCode (type = 0) or SmfCode (type = 1) or PsfCode (Type = 2)
    sType As String * 1 'Vehicle type
    iVefCode As Integer 'Spot vehicle
    iSdfVefCode As Integer  'Vehicle from SDF
    iLen As Integer 'Spot length from line, required for Bonus spots on EDI
    lCntrNo As Long 'Contract number used when updating
    iBillMissed As Integer  'True=Billed Missed (Sdf.sSchStatus set to S in mInvoiceRpt)
    sSchStatus As String * 1
    sAdjAirTime As String * 6
    '5/22/14
    sXMidMessage As String * 1  'Y=Add Cross Midnight to printout and EDI
    iGenDate(0 To 1) As Integer 'Generation Date
    iGenTime(0 To 1) As Integer 'Generation time
End Type
Type TYPESORTKEY
    sKey As String * 245        '5-13-06 chg from 240 to 245 for games, requires 244
    lChfCode As Long    'Internal code number for Contract
    sInvGp As String * 1        'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    iType As Integer    '0=Sdf; 1=Smf (for missed of MG not showing); 2=Psf
    iLineNo As Integer
    iBilled As Integer  'True=Spot billed; False=Spot not billed
    sPriceType As String * 1
    lCode As Long     'SdfCode (type = 0) or SmfCode (type = 1) or PsfCode (Type = 2)
    sType As String * 1 'Vehicle type
    iVefCode As Integer 'Spot vehicle
    iSdfVefCode As Integer  'Vehicle from SDF
    iLen As Integer 'Spot length from line, required for Bonus spots on EDI
    lCntrNo As Long 'Contract number used when updating
    iBillMissed As Integer  'True=Billed Missed (Sdf.sSchStatus set to S in mInvoiceRpt)
    sSchStatus As String * 1
    sAdjAirTime As String * 6
    '5/22/14
    sXMidMessage As String * 1  'Y=Add Cross Midnight to printout and EDI
    lSbfDate As Long      'Used only with type = 10; Start date of billing period
    lCPMDate As Long        'Used only with type= 11; End date of billing perion
    iGenDate(0 To 1) As Integer 'Generation Date
    iGenTime(0 To 1) As Integer 'Generation time
End Type

Type TYPESORTISR
    sKey As String * 245       '5-13-06 chg from 240 to 245 for games, requires 244
    tSort As TYPESORTNOKEY
End Type

Type TYPESORTREPSPOTS
    sKey As String * 80 'Ordered Vehicle, Line #, Game #, Date and Time
    tSdf As SDF
End Type

Type ISR
    sKey As String * 245    '5-13-06 chg from 240 to 245 for games, requires 244
                            'MnfSort #(5); Agy/City(40+5) or Dir Advt Name;
                            '| Advt Name (30);
                            '| Contract #(8);
                            '| If Contract by Prod(35) use Copy Product Name else Contract Product
                            '| Ordered Vehicle Name(40), if bonus special character added to force sort to end;
                            '| Line Number(4);
                            '| Week #(4);
                            '| 0 if not Extra bonus; 1 if Extra bonus
                            '| Spot Vehicle Name(40) except for InvAirOrder=O or S, then its Line Vehicle Name;
                            '| Game No (4)

                            '| Spot Date(5)
                            '| Spot Time(6)
                            '| Package Line No
    lChfCode As Long    'Internal code number for Contract
    sInvGp As String * 1        'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    iType As Integer    '0=Sdf; 1=Smf (for missed of MG not showing); 2=Psf
    iLineNo As Integer
    iBilled As Integer  'True=Spot billed; False=Spot not billed
    sPriceType As String * 1
    lCode As Long     'SdfCode (type = 0) or SmfCode (type = 1) or PsfCode (Type = 2)
    sType As String * 1 'Vehicle type
    iVefCode As Integer 'Spot vehicle
    iSdfVefCode As Integer  'Vehicle from SDF
    iLen As Integer 'Spot length from line, required for Bonus spots on EDI
    lCntrNo As Long 'Contract number used when updating
    iBillMissed As Integer  'True=Billed Missed (Sdf.sSchStatus set to S in mInvoiceRpt)
    sSchStatus As String * 1
    sAdjAirTime As String * 6
    iGenDate(0 To 1) As Integer 'Generation Date
    iGenTime(0 To 1) As Integer 'Generation time
End Type
Type ISRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    iGenTime(0 To 1) As Integer 'Generation time
    sKey As String * 245        '5-13-06 chg from 240 to 245 for games    'Requires 244
End Type
Type ISRKEY1
    iGenDate(0 To 1) As Integer 'Generation Date
    iGenTime(0 To 1) As Integer 'Generation time
    iType As Integer    '0=Sdf; 1=Smf (for missed of MG not showing); 2=Psf
    lCode As Long     'SdfCode (type = 0) or SmfCode (type = 1) or PsfCode (Type = 2)
End Type
'Selected contract sort
Type SELTYPESORT
    sKey As String * 100    'Requires 91
                            'MnfSort #(5); Agy/City(40+5+1) or Dir Advt Name;
                            '| Advt Name (30);
                            '| Contract #(8);
    lChfCode As Long    'Internal code number for Contract
    iAdfCode As Integer
    lCntrNo As Long 'Contract number used when updating
    iSlfCode As Integer 'TTP 10813 - PDF invoice - selective invoice feature checklist
    sBillCycle As String * 1
    sEndDate As String * 10
    iAgfCode As Integer
    iPosted As Integer  'Indicates if Posting Completed for Contract that are obtaining Post Spots from Station Invoices
End Type
Type RVFVEF
    iVefCode As Integer     'Vehicle       Summary          Rvf: Bill         Air
                            'Conventional  Spot Vehicle          Spot Veh     Spot Veh
                            'Package-Order Pk Line Veh           Pk Line Veh  Spot Veh of Hidden
                            'Package-Aired Spot Veh              Spot Veh     Spot Veh
                            'Virtual       Spot of Virtual       Virtual Veh  Conv Veh of Virtual
                            'Template (later) same as Package-Order
    iPkLineNo As Integer
    lGsfCode As Long
    sTotalGross As String   'Total dollars within billing period
    sTotalVefDollars As String     'Total Vehicle Dollars
    sTotalVefBilledDollars As String   'Total Vehicle Billed Dollars
    lANoSpots As Long    'Aired number of Spots
    lBNoSpots As Long    'Bouns number of Spots
    lPrice As Long          'Flight price
    iLnVefCode As Integer   'Sold line vehicle (used only with cluster export)
    lAcquisitionCost As Long    'sum of the Acquisition cost
End Type
Type WKDEF
    sDays As String * 21    'Days XX XX XX XX XX XX XX
    sNoSpots As String * 3  'Number of Spots for week
    sWkPrice As String * 12    'Rate for week
    lWkSDate As Long        'Week Start Date
    lWkEDate As Long      'Week End Date
    iWkShown As Integer     'True = week shown; False = Week not shown
    sEDIDays As String * 7
    sEDIRate As String * 12
    iGameNo As Integer
    iVisitMnfCode As Integer
    iHomeMnfCode As Integer
End Type
Type UPDATEADFAGF
    iCode As Integer
    lGross As Long
    lNet As Long
    iTranDate(0 To 1) As Integer
End Type
Type PKLNGEN
    tClf As CLF
    lRecPos As Long
    lCntrNo As Long
End Type
Type RPSELINFO
    sKey As String * 100    'Advt Name "|" CntrNo "|" InvNo
    lCntrNo As Long
    lInvNo As Long
    iAdfCode As Integer
    iAgfCode As Integer
    iBillVefCode As Integer
    iAirVefCode As Integer
    iPkLineNo As Integer
    sCashTrade As String * 1
    iMnfItem As Integer '>0 => NTR
    lAirTax1 As Long
    lAirTax2 As Long
    lNTRTax1 As Long
    lNTRTax2 As Long
    iSelAllowed As Integer
    sReason As String * 50
    sBillCycle As String * 1
    sTranDate As String * 10
    lInvStartDate As Long
    lInvEndDate As Long
    iSlfCode As Integer
End Type

Type RPINFO
    lCntrNo As Long
    lInvNo As Long
    iAdfCode As Integer
    iAgfCode As Integer
    iBillVefCode As Integer
    iAirVefCode As Integer
    iPkLineNo As Integer
    sCashTrade As String * 1
    iMnfItem As Integer '==Air Time only; >0 = NTR only; -1=Combined
    lAirTax1 As Long
    lAirTax2 As Long
    lNTRTax1 As Long
    lNTRTax2 As Long
    sBillCycle As String * 1
    lInvStartDate As Long
    lInvEndDate As Long
    iSelectiveEmail As Integer 'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
End Type

Type MKTTOTAL
    iMnfVehGp3Mkt As Integer
    lONoSpots As Long
    sOMktTotalGross As String   'Total dollars within billing period
    lANoSpots As Long
    sAMktTotalGross As String   'Total dollars within billing period
End Type

Type EXPORTCOUNT
    lONoSpots As Long    'Order number of spots
    lANoSpots As Long    'Aired number of spots
    lBNoSpots As Long    'Bonus number of spots
    iCombineID As Integer
    lOGross As Long         'Ordered Gross
    iCommPct As Integer     'Agency Commission %
    lPrice As Long
    iLnVefCode As Integer
End Type

'Vehicle sort
Type VEHSORT
    sKey As String * 70 'Group Code|Veh Sort|Vehicle Name\Vehicle Code
    iCode As Integer
    iStatus As Integer  '0=Billed Ok; 1=Unbilled Dates
End Type

'Not completely posted dates
Type NOTMARKCOMPLETE
    sKey As String * 70     'Vehicle Name | Game Number | Date
    sVehName As String * 40
    iGameNo As Integer
    sDate As String * 10
End Type

Type VIEWLISTINFO
    iType As Integer    '1=EDI Call Letters missing
    lContrNo As Long
    'sErrMsg As String * 80
    'sErrMsg As String * 120 'JW allow error message to be longer
    sErrMsg As String * 512 'JW allow error message to be longer
End Type

Public tgNotMarkComplete() As NOTMARKCOMPLETE

Type AIRNTRCOMBINE
    lChfCode As Long
    lInvNo As Long
    sCashTrade As String * 1
    lIvrCode As Long
    iEDIFlag As Integer
    iArfPDFEMailCode As Integer
End Type

Type INVCNTRSTATUS
    lCntrNo As Long
    lChfCodeSchd As Long
    lChfCodeAltered As Long
    iAdfCode As Integer
    sSchStatus As String * 1
    lEarliestDate As Long
    sBillCycle As String * 1
End Type

'TTP 10515 - NTR Invoices - "NTR INVOICE and AFFIDAVIT" not displaying correct on NTR Invoices it shows "INVOICES and AFFIDAVIT" started with V81
Type INVAIRNTRSTATUS
    lCntrNo As Long
    lInvNo As Long
    bHasNTR As Boolean
    bHasAir As Boolean
    iPayeeCode As Integer 'Fix TTP 10826 / TTP 10813 - RE: v81 TTP 10826 - updated test results Issue #4
End Type

Type UNDOINFO
    lCntrNo As Long
    sType As String * 1     'I=NTR; F=Installment; R=Rep; A=Air Time
    lRvfCode As Long
    lSbfCode As Long
    lPcfCode As Long
    iPass As Integer    '0=Rvf; 1=Phf (IN); 2=Phf(HI)
    sBillCycle As String * 1
    lInvStartDate As Long
    lInvEndDate As Long
    lInvNo As Long
End Type

'Used to create advertiser separation when creating package spots
Type PSFINFO
    lDate As Long
    lTime As Long
End Type


Type CFFINVEXT
    lCode As Long
    iStartDate(0 To 1) As Integer  'Start Date of line
                                   'Date Byte 0:Day, 1:Month, followed by 2 byte year
End Type
Public Const CFFINVEXTPK As String = "LII"

Dim tmRvf As RVF                'RVF record image
Dim tmRvfSrchKey2 As LONGKEY0
Dim tmRvfSrchKey5 As RVFKEY5
Dim imRvfRecLen As Integer        'RVF record length
Dim tmUndoInfo() As UNDOINFO
Dim tmSbf As SBF                'SBF record image
Dim tmSbfSrchKey0 As SBFKEY0
Dim tmSbfSrchKey1 As LONGKEY0
Dim imSbfRecLen As Integer
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim tmClfSrchKey0 As CLFKEY0
Dim imClfRecLen As Integer
Dim hmClf As Integer
Dim tmClf As CLF
'Multi-Media Sold
Dim tmMsf As MSF
Dim imMsfRecLen As Integer
Dim tmMsfSrchKey2 As MSFKEY2
'Multi-Media Events
Dim tmMgf As MGF
Dim imMgfRecLen As Integer
Dim tmMgfSrchKey1 As MGFKEY1
'Game Schedule
Dim tmGsf As GSF
Dim tmGsfSrchKey1 As GSFKEY1
Dim imGsfRecLen As Integer
'Package spots
Dim tmPsf As SDF                'SDF record image
Dim tmPsfSrchKey4 As SDFKEY4            'SDF record image
Dim imPsfRecLen As Integer        'SDF record length
'Air Spots
Dim tmSdf As SDF                'SDF record image
Dim tmSdfSrchKey4 As SDFKEY4
Dim tmSdfSrchKey3 As LONGKEY0
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey2 As LONGKEY0   'smf key 0
Dim tmSmfSrchKey4 As SMFKEY4            'SMF record image
Dim imSmfRecLen As Integer        'SMF record length

Dim tmApf As APF        'CFF record image
Dim tmApfSrchKey5 As LONGKEY0    'CFF key record image
Dim imApfRecLen As Integer        'CFF record length

Type ACQWITHCOMMINFO
    lChfCode As Long
    lClfCode As Long
    iAirVefCode As Integer
    iOrderCount As Integer
    iAirCount As Integer
    lPrfCode As Long
    lInvNo As Long
    iMnfItem As Integer
    lSbfCode As Long
    lAcqCost As Long
End Type

Type ADVANCEBILLINFO
    lChfCode As Long
    lDate As Long
End Type

Public tgAdanceBillInfo() As ADVANCEBILLINFO
Private sbf_rst As ADODB.Recordset

Type ADVANCEBILLMERGEINFO
    lRvfInvNo As Long
    lPhfInvNo As Long
End Type

Public tgAdvanceBillMergeInfo() As ADVANCEBILLMERGEINFO
Private merge_rst As ADODB.Recordset
'******************************************************************************
' iihf_ImptInvHeader Record Definition
'
'******************************************************************************
Type IIHF
    lCode                 As Long            ' Import Invoice History Header
                                             ' internal reference code
    iVefCode              As Integer         ' Vehicle Internal code
    lChfCode              As Long            ' Contract internal code
    iInvStartDate(0 To 1) As Integer         ' Invoice month start date
    sFileName             As String * 100    ' Import File Name
    sStnEstimateNo        As String * 20     ' Station Estimate number
    sStnInvoiceNo         As String * 20     ' Station Invoice Number
    sStnContractNo        As String * 20     ' Station Contract number
    sStnAdvtName          As String * 30     ' Station Advertiser Name
    lAmfCode              As Long            ' Advertiser remap reference code
    sSourceForm           As String * 2      ' Invoice source form: M=Marketron with separate Line;R=RadioTraffic; W=WideOrbit; C=Manual Post Counts; T=Manual Post Counts/Times
    sUnused               As String * 10     ' Unused
End Type


'Type IIHFKEY0
'    lCode                 As Long
'End Type

Type IIHFKEY1
    iVefCode              As Integer
    iInvStartDate(0 To 1) As Integer
End Type

Type IIHFKEY2
    lChfCode              As Long
    iVefCode              As Integer
    iInvStartDate(0 To 1) As Integer
End Type

Type IIHFKEY3
    sFileName             As String * 100    ' Import File Name
    iInvStartDate(0 To 1) As Integer         ' Invoice month start date
End Type

'******************************************************************************
' iidf_ImptInvDetail Record Definition
'
'******************************************************************************
Type IIDF
    lCode                 As Long            ' Import Invoice spot detail
                                             ' information internal reference
                                             ' code
    lIihfCode             As Long            ' Import Invoice Header reference
                                             ' code
    sSpotMatchType        As String * 1      ' Spot Match Type (C=Spot
                                             ' Correlated; I=Import Spot
                                             ' ignored; M=Traffic spot set to
                                             ' Missed)
    lSdfCode              As Long            ' spot detail record reference code
    iStnSpotAirDate(0 To 1) As Integer       ' Station Spot Date
    iStnSpotAirTime(0 To 1) As Integer       ' Station Spot Time
    iStnSpotLen           As Integer         ' Station Spot Length
    lStnCpfCode           As Long            ' Station ISCI stored into cpf (cpfISCI)
    lStnDPStartTime       As Long            ' Station Daypart start time
    lStnDPEndTime         As Long            ' Station Daypart End time
    sStnDPDays            As String * 7      ' Station Daypart days    iOrigSpotDate(0 To 1) As Integer         ' Original Spot Date
    iOrigSpotDate(0 To 1) As Integer         ' Original Spot Date
    iOrigSpotTime(0 To 1) As Integer         ' Original Spot Time
    sAgyCompliant         As String * 1      ' Agency Compliant: A=Aired as Sold; O=Aired Outside Sold; N=Did not air
    sStnRate              As String * 8      ' Spot Rate (xxxxx.xx)
    sUnused               As String * 10     ' Unused
End Type

'******************************************************************************
' apf_Acq_Payable Record Definition
'
'******************************************************************************
Type APF
    lCode                 As Long            ' Acquisition Payable Internal
                                             ' reference code.
    iAgfCode              As Integer         ' Agency Reference code
    iAdfCode              As Integer         ' Advertiser Reference code
    lPrfCode              As Long            ' Product Reference code
    iSlfCode              As Integer         ' Salesperson reference code
    lCntrNo               As Long            ' Contract number
    lInvNo                As Long            ' Invoice number
    sAgyEst               As String * 20     ' Agency Estimate number. chfAgyEst+chfTitle
    iMnfItem              As Integer         ' NTR Type
    lSbfCode              As Long            ' NTR Type
    iInvDate(0 To 1)      As Integer         ' Invoice date
    iOrderSpotCount       As Integer         ' Order Spot Count
    iAiredSpotCount       As Integer         ' Aired Spot Count
    lAcquisitionCost      As Long            ' Acquisition cost (xxxxx.xx) taken
                                             ' from the line (cost per spot)
    iAcqCommPct           As Integer         ' Acquisition commission Percentage
                                             ' taken from vbfAcqCommPct (xxx.xx)
    iFullyPaidDate(0 To 1) As Integer        ' Fully paid date (if not fully
                                             ' paid the date will be 1/1/1970)
    sStationInvNo         As String * 20     ' Station Invoice Number
    sStationCntrNo        As String * 20     ' Station contract number
    iVefCode              As Integer         ' vehicle reference code
    sUnused               As String * 20     ' Unused
End Type


'Type APFKEY0
'    lCode                 As Long
'End Type

Type APFKEY1
    iVefCode              As Integer
    iFullyPaidDate(0 To 1) As Integer
End Type

'Type APFKEY2
'    iAdfCode              As Integer
'End Type

Type APFKEY3
    iFullyPaidDate(0 To 1) As Integer
    iVefCode              As Integer
End Type

Type APFKEY4
    lCntrNo               As Long
    iFullyPaidDate(0 To 1) As Integer
End Type

'Type APFKEY5
'    lInvNo                As Long
'End Type

Type APFKEY6
    sStationInvNo         As String * 20
End Type

Type APFKEY7
    sStationCntrNo        As String * 20
End Type

Type INVEXPORT_HEADER               '5-12-17 Invoice export feature - header into for spot & ntr files
    sInvNo As String * 10
    sInvStartDate As String * 10
    sCntrNo As String * 10
    sCntStartDate As String * 10
    sCntEndDate As String * 10
    sCashTrade As String * 1
    sAgyComm As String * 6
    sPayee As String * 40
    sAdvName As String * 30
    sProduct As String * 35
    sSlspName As String * 41
    sSlspOffice As String * 20
    sAgfCode As String * 5
    sAdfCode As String * 5
    sSlfCode As String * 5
    sTerms As String * 20
End Type

Type INVEXPORT_SPOT
    sReconciliationAmt As String * 11
    sWeekOf As String * 10
    sVehicle As String * 40
    sLen As String * 4
    sOrderedDays As String * 20
    sSpotsPerWk As String * 4
    sLine As String * 5
    sDateAired As String * 10
    sTimeAired As String * 10
    sAirStatus As String * 1
    sMGBonus As String * 1
    sMGMissedDate As String * 10
    sSpotPrice As String * 11
    sCopy As String * 20
End Type

Type INVEXPORT_NTR
    sNTRDate As String * 10
    sVehicle As String * 40
    sDescription As String * 80
    sGross As String * 11
    sNet As String * 11
End Type

'Private Type IIDFKEY0
'    lCode                 As Long
'End Type

'Private Type IIDFKEY1
'    lIihfCode             As Long
'End Type

'Private Type IIDFKEY2
'    lSdfCode             As Long
'End Type

'
'           Adjust the DP ordered times if there is are override times.
'           Test Site option to determine if applicable.  If so, check to
'           see which zone to adjust based on the vehicle options time zone table.
'           <input> ilVefCode - vehicle code
'                   slEndTime - time



'7/31/19: Remove parameter as not used and was amount was over 21,000,000.00
Public Function gUndoInvoice(llUndoStartDateStd As Long, llUndoEndDateStd As Long, llUndoStartDateCal As Long, llUndoEndDateCal As Long, llUndoStartDateWk As Long, llUndoEndDateWk As Long, tlRPInfo() As RPINFO, hlRvf As Integer, hlPhf As Integer, hlSbf As Integer, hlSdf As Integer, hlSmf As Integer, hlPsf As Integer, hlVef As Integer, hlChf As Integer, hlClf As Integer, hlMsf As Integer, hlMgf As Integer, hlGsf As Integer, hlApf As Integer, ilStdCntrRolledBack As Integer, ilCalCntrRolledBack As Integer, ilWkCntrRolledBack As Integer) As Integer  ', llReconcileAdjustment As Long) As Integer
    '                                 -------------Files with Bill flags----------------
    'Type         -------RVF-------    SDF         PSF       ----------SBF--------------    MGF
    '             sbfCode   MnfItem   Bill        Bill       TranType   IhfCode   Billed  Billed
    '
    'Air Time       0         0        Set         Set           -        -         -        -
    '
    'NTR           Set       Set        -           -            I        0        Set       -
    '
    'Rep            0         0         -           -            T        0        Set       -
    '
    'MultiMedia    Set       Set        -           -            I       Set       Set      Set
    '
    'Installment   Set        0         -           -            F        0        Set       -
    '

    'Rep SBF created when spots posted.
    'To distingush Rep from Air Time:  Test if sbfTranType = "T" defined for the contract
    'To distingush NTR and MultiMedia:  Test if sbfIhfCode defined
    '
    'If Installment Contract RVF exist and Site set to Revenue, then NTR and Air time PHF records could exist
    '
    'If Installment Contract but no RVF, then NTR and Air time PHF records could exist regardless of site option
    '
    Dim llUndoStartDate As Long
    Dim llUndoEndDate As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilOk As Integer
    Dim ilRvf As Integer
    Dim slGross As String
    Dim slNet As String
    Dim llRepUnbilledChfCode As Long
    Dim llGameInvUnbilledChfCode As Long
    Dim llNTRUnbilledChfCode As Long
    Dim llSpotsUnbilledChfCode As Long
    Dim ilPass As Integer
    Dim slType As String
    Dim slNowDate As String
    Dim llSpotInvNo As Long
    Dim llNTRInvNo As Long
    Dim ilMerge As Integer
    '6/27/18: Add airing vehicle test if not combining vehicles within a contract
    Dim ilAirVefCode As Integer
    '11/27/19:
    Dim blAnyInstallments As Boolean
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    Dim slSQLQuery As String
    Dim llRet As Long

    imRvfRecLen = Len(tmRvf)
    imSbfRecLen = Len(tmSbf)
    'imIbfRecLen = Len(tmIbf)
    imVefRecLen = Len(tmVef)
    imCHFRecLen = Len(tmChf)
    imClfRecLen = Len(tmClf)
    imSdfRecLen = Len(tmSdf)
    imSmfRecLen = Len(tmSmf)
    imPsfRecLen = Len(tmPsf)
    imGsfRecLen = Len(tmGsf)
    imMgfRecLen = Len(tmMgf)
    imApfRecLen = Len(tmApf)
    ilStdCntrRolledBack = False
    ilCalCntrRolledBack = False
    ilWkCntrRolledBack = False
    blAnyInstallments = False
    '7/31/19: Remove parameter as not used and was amount was over 21,000,000.00
    'llReconcileAdjustment = 0
    ReDim tmUndoInfo(0 To 0) As UNDOINFO
    ReDim tlRPInfoPlusMerge(0 To UBound(tlRPInfo)) As RPINFO
    For ilLoop = 0 To UBound(tlRPInfo) - 1 Step 1
        tlRPInfoPlusMerge(ilLoop) = tlRPInfo(ilLoop)
    Next ilLoop
    For ilLoop = 0 To UBound(tlRPInfo) - 1 Step 1
        For ilMerge = 0 To UBound(tgAdvanceBillMergeInfo) - 1 Step 1
            If tlRPInfo(ilLoop).lInvNo = tgAdvanceBillMergeInfo(ilMerge).lRvfInvNo Then
                tlRPInfoPlusMerge(UBound(tlRPInfoPlusMerge)) = tlRPInfo(ilLoop)
                tlRPInfoPlusMerge(UBound(tlRPInfoPlusMerge)).lInvNo = tgAdvanceBillMergeInfo(ilMerge).lPhfInvNo
                ReDim Preserve tlRPInfoPlusMerge(0 To UBound(tlRPInfoPlusMerge) + 1) As RPINFO
            End If
        Next ilMerge
    Next ilLoop
    For ilLoop = 0 To UBound(tlRPInfoPlusMerge) - 1 Step 1
        For ilPass = 0 To 2 Step 1
            tmRvfSrchKey5.lInvNo = tlRPInfoPlusMerge(ilLoop).lInvNo
            If ilPass = 0 Then
                ilRet = btrGetEqual(hlRvf, tmRvf, imRvfRecLen, tmRvfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                slType = "IN"
            Else
                ilRet = btrGetEqual(hlPhf, tmRvf, imRvfRecLen, tmRvfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilPass = 1 Then
                    slType = "IN"
                Else
                    slType = "HI"
                End If
            End If
            ilAirVefCode = tlRPInfoPlusMerge(ilLoop).iAirVefCode
            Do While (ilRet = BTRV_ERR_NONE) And (tmRvf.lInvNo = tlRPInfoPlusMerge(ilLoop).lInvNo)
                ilOk = False
                '6/27/18: Add airing vehicle test if not combining vehicles within a contract
                'If tmRvf.sInvoiceUndone <> "Y" Then
                If (tmRvf.sInvoiceUndone <> "Y") And ((tgSpf.sBCombine <> "N") Or ((tgSpf.sBCombine = "N") And (tmRvf.iAirVefCode = ilAirVefCode))) Then
                    If tmRvf.lSbfCode > 0 Then
                        tmSbfSrchKey1.lCode = tmRvf.lSbfCode
                        ilRet = btrGetEqual(hlSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If tmSbf.sTranType = "I" Then
                                'NTR
                                If tmRvf.sTranType = slType Then
                                    ilOk = True
                                    tmUndoInfo(UBound(tmUndoInfo)).sType = "I"
                                End If
                            ElseIf tmSbf.sTranType = "F" Then
                                'Installment
                                If tmRvf.sTranType = slType Then
                                    ilOk = True
                                    tmUndoInfo(UBound(tmUndoInfo)).sType = "F"
                                End If
                            End If
                        End If
                    Else
                        'Air Time or Rep
                        tmVefSrchKey.iCode = tmRvf.iBillVefCode
                        ilRet = btrGetEqual(hlVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If ((tmRvf.sTranType = slType) And (tmVef.sType <> "R")) Or ((tmRvf.sTranType = "AN") And (tmVef.sType = "R") And (tmRvf.lRefInvNo = 0)) Then
                                ilOk = True
                                If tmVef.sType = "R" Then
                                    tmUndoInfo(UBound(tmUndoInfo)).sType = "R"
                                Else
                                    tmUndoInfo(UBound(tmUndoInfo)).sType = "A"
                                End If
                            End If
                        End If
                    End If
                End If
                If ilOk Then
                    tmUndoInfo(UBound(tmUndoInfo)).lRvfCode = tmRvf.lCode
                    tmUndoInfo(UBound(tmUndoInfo)).lSbfCode = tmRvf.lSbfCode
                    tmUndoInfo(UBound(tmUndoInfo)).lPcfCode = tmRvf.lPcfCode
                    tmUndoInfo(UBound(tmUndoInfo)).lCntrNo = tmRvf.lCntrNo
                    tmUndoInfo(UBound(tmUndoInfo)).iPass = ilPass
                    tmUndoInfo(UBound(tmUndoInfo)).sBillCycle = tlRPInfoPlusMerge(ilLoop).sBillCycle
                    tmUndoInfo(UBound(tmUndoInfo)).lInvStartDate = tlRPInfoPlusMerge(ilLoop).lInvStartDate
                    tmUndoInfo(UBound(tmUndoInfo)).lInvEndDate = tlRPInfoPlusMerge(ilLoop).lInvEndDate
                    tmUndoInfo(UBound(tmUndoInfo)).lInvNo = tlRPInfoPlusMerge(ilLoop).lInvNo
                    ReDim Preserve tmUndoInfo(0 To UBound(tmUndoInfo) + 1) As UNDOINFO
                End If
                If ilPass = 0 Then
                    ilRet = btrGetNext(hlRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Else
                    ilRet = btrGetNext(hlPhf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                End If
            Loop
        Next ilPass
    Next ilLoop
    If UBound(tmUndoInfo) > 0 Then
        ArraySortTyp fnAV(tmUndoInfo(), 0), UBound(tmUndoInfo), 0, LenB(tmUndoInfo(0)), 0, -2, 0
    End If
    llRepUnbilledChfCode = -1
    llGameInvUnbilledChfCode = -1
    llNTRUnbilledChfCode = -1
    llSpotsUnbilledChfCode = -1
    llSpotInvNo = -1
    llNTRInvNo = -1
    For ilRvf = 0 To UBound(tmUndoInfo) - 1 Step 1
        tmChfSrchKey1.lCntrNo = tmUndoInfo(ilRvf).lCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmUndoInfo(ilRvf).lCntrNo) And (tmChf.sSchStatus = "A")
        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmUndoInfo(ilRvf).lCntrNo) And ((tmChf.sSchStatus <> "M") And (tmChf.sSchStatus <> "F"))
            ilRet = btrGetNext(hlChf, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
        If tmChf.sBillCycle = "C" Then
            If tmUndoInfo(ilRvf).sBillCycle = "C" Then
                llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
            Else
                llUndoStartDate = llUndoStartDateCal
                llUndoEndDate = llUndoEndDateCal
            End If
            ilCalCntrRolledBack = True
        ElseIf tmChf.sBillCycle = "W" Then
            If tmUndoInfo(ilRvf).sBillCycle = "W" Then
                llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
            Else
                llUndoStartDate = llUndoStartDateWk
                llUndoEndDate = llUndoEndDateWk
            End If
            ilWkCntrRolledBack = True
        Else
            If tmChf.sBillCycle = tmUndoInfo(ilRvf).sBillCycle Then
                llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
            Else
                llUndoStartDate = llUndoStartDateStd
                llUndoEndDate = llUndoEndDateStd
            End If
            ilStdCntrRolledBack = True
        End If
        ilAgePeriod = Month(llUndoEndDate)
        ilAgingYear = Year(llUndoEndDate)

        If tmChf.sInstallDefined = "Y" Then
            blAnyInstallments = True
            If llGameInvUnbilledChfCode <> tmChf.lCode Then
                ilRet = mUnbillGameInv(llUndoStartDate, llUndoEndDate, tmChf.lCode, hlMsf, hlMgf, hlGsf)
                llGameInvUnbilledChfCode = tmChf.lCode
            End If
            If (llNTRUnbilledChfCode <> tmChf.lCode) Or (llNTRInvNo <> tmUndoInfo(ilRvf).lInvNo) Then
                ilRet = mUnbillSpecialBill(llUndoStartDate, llUndoEndDate, tmChf.lCode, "I", hlSbf)
                llNTRUnbilledChfCode = tmChf.lCode
                llNTRInvNo = tmUndoInfo(ilRvf).lInvNo
            End If
        End If
        If (llSpotsUnbilledChfCode <> tmChf.lCode) Or (llSpotInvNo <> tmUndoInfo(ilRvf).lInvNo) Then
            ilRet = mUnbillAirSpots(llUndoStartDate, llUndoEndDate, tmChf.lCode, hlSdf, hlSmf, hlClf)
            ilRet = mClearPackageSpots(llUndoStartDate, llUndoEndDate, tmChf.lCode, hlPsf)
            llSpotsUnbilledChfCode = tmChf.lCode
            llSpotInvNo = tmUndoInfo(ilRvf).lInvNo
        End If
        If tmUndoInfo(ilRvf).lSbfCode > 0 Then
            Do
                tmSbfSrchKey1.lCode = tmUndoInfo(ilRvf).lSbfCode
                ilRet = btrGetEqual(hlSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    tmSbf.sBilled = "N"
                    ilRet = btrUpdate(hlSbf, tmSbf, imSbfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If (tmUndoInfo(ilRvf).sType = "I") And (tmSbf.iIhfCode > 0) And (llGameInvUnbilledChfCode <> tmChf.lCode) Then
                ilRet = mUnbillGameInv(llUndoStartDate, llUndoEndDate, tmChf.lCode, hlMsf, hlMgf, hlGsf)
                llGameInvUnbilledChfCode = tmChf.lCode
            End If
        Else
            If (tmUndoInfo(ilRvf).sType = "R") And (llRepUnbilledChfCode <> tmChf.lCode) Then
                ilRet = mUnbillSpecialBill(llUndoStartDate, llUndoEndDate, tmChf.lCode, "T", hlSbf)
                llRepUnbilledChfCode = tmChf.lCode
            End If
        End If
        If tmUndoInfo(ilRvf).lPcfCode > 0 Then
            slSQLQuery = "UPDATE ibf_Impression_Bill SET "
            slSQLQuery = slSQLQuery & "ibfBilled = '" & "N" & "' "
            slSQLQuery = slSQLQuery & " WHERE (ibfCntrNo = " & tmChf.lCntrNo
            slSQLQuery = slSQLQuery & " And ibfBillYear = " & ilAgingYear
            slSQLQuery = slSQLQuery & " And ibfBillMonth = " & ilAgePeriod
            slSQLQuery = slSQLQuery & ")"
            llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        End If
        Do
            tmRvfSrchKey2.lCode = tmUndoInfo(ilRvf).lRvfCode
            If tmUndoInfo(ilRvf).iPass = 0 Then
                ilRet = btrGetEqual(hlRvf, tmRvf, imRvfRecLen, tmRvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            Else
                ilRet = btrGetEqual(hlPhf, tmRvf, imRvfRecLen, tmRvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            End If
            If ilRet = BTRV_ERR_NONE Then
                'Set the Undo flag
                tmRvf.sInvoiceUndone = "Y"
                If tmUndoInfo(ilRvf).iPass = 0 Then
                    ilRet = btrUpdate(hlRvf, tmRvf, imRvfRecLen)
                Else
                    ilRet = btrUpdate(hlPhf, tmRvf, imRvfRecLen)
                End If
                gPDNToStr tmRvf.sGross, 2, slGross
                gPDNToStr tmRvf.sNet, 2, slNet
                If gStrDecToLong(slGross, 2) <> 0 Then
                    If Left$(slGross, 1) = "-" Then
                        slGross = Mid(slGross, 2)
                    Else
                        slGross = "-" & slGross
                    End If
                    gStrToPDN slGross, 2, 6, tmRvf.sGross
                End If
                If gStrDecToLong(slNet, 2) <> 0 Then
                    If Left$(slNet, 1) = "-" Then
                        slNet = Mid(slNet, 2)
                    Else
                        slNet = "-" & slNet
                    End If
                    gStrToPDN slNet, 2, 6, tmRvf.sNet
                End If
                If tmRvf.lTax1 <> 0 Then
                    tmRvf.lTax1 = -tmRvf.lTax1
                End If
                If tmRvf.lTax2 <> 0 Then
                    tmRvf.lTax2 = -tmRvf.lTax2
                End If
                If tmRvf.lAcquisitionCost <> 0 Then
                    tmRvf.lAcquisitionCost = -tmRvf.lAcquisitionCost
                End If
                tmRvf.sInvoiceUndone = "Y"
                slNowDate = Format$(gNow(), "m/d/yy")
                gPackDate slNowDate, tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1)
                tmRvf.iUrfCode = tgUrf(0).iCode
                tmRvf.lCode = 0
                If tmUndoInfo(ilRvf).iPass = 0 Then
                    '7/31/19: Remove parameter as not used and was amount was over 21,000,000.00
                    'llReconcileAdjustment = llReconcileAdjustment + tmRvf.lTax1 + tmRvf.lTax2 + gStrDecToLong(slNet, 2)
                    ilRet = btrInsert(hlRvf, tmRvf, imRvfRecLen, INDEXKEY2)
                Else
                    ilRet = btrInsert(hlPhf, tmRvf, imRvfRecLen, INDEXKEY2)
                End If
                Do
                    tmApfSrchKey5.lCode = tmRvf.lInvNo
                    ilRet = btrGetEqual(hlApf, tmApf, imApfRecLen, tmApfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hlApf)
                Loop
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    Next ilRvf
    '11/27/19: Handle cases where sbf bill flag set but not cleared (rvf record does not exist)
    If blAnyInstallments Then
        llSpotsUnbilledChfCode = -1
        llSpotInvNo = -1
        For ilRvf = 0 To UBound(tmUndoInfo) - 1 Step 1
            tmChfSrchKey1.lCntrNo = tmUndoInfo(ilRvf).lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmUndoInfo(ilRvf).lCntrNo) And (tmChf.sSchStatus = "A")
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmUndoInfo(ilRvf).lCntrNo) And ((tmChf.sSchStatus <> "M") And (tmChf.sSchStatus <> "F"))
                ilRet = btrGetNext(hlChf, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Loop
            If tmChf.sBillCycle = "C" Then
                If tmUndoInfo(ilRvf).sBillCycle = "C" Then
                    llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                    llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
                Else
                    llUndoStartDate = llUndoStartDateCal
                    llUndoEndDate = llUndoEndDateCal
                End If
            ElseIf tmChf.sBillCycle = "W" Then
                If tmUndoInfo(ilRvf).sBillCycle = "W" Then
                    llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                    llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
                Else
                    llUndoStartDate = llUndoStartDateWk
                    llUndoEndDate = llUndoEndDateWk
                End If
            Else
                If tmChf.sBillCycle = tmUndoInfo(ilRvf).sBillCycle Then
                    llUndoStartDate = tmUndoInfo(ilRvf).lInvStartDate
                    llUndoEndDate = tmUndoInfo(ilRvf).lInvEndDate
                Else
                    llUndoStartDate = llUndoStartDateStd
                    llUndoEndDate = llUndoEndDateStd
                End If
            End If
            If (llSpotsUnbilledChfCode <> tmChf.lCode) Or (llSpotInvNo <> tmUndoInfo(ilRvf).lInvNo) Then
                llSpotsUnbilledChfCode = tmChf.lCode
                llSpotInvNo = tmUndoInfo(ilRvf).lInvNo
                If tmChf.sInstallDefined = "Y" Then
                    ilRet = mUnbillSpecialBill(llUndoStartDate, llUndoEndDate, tmChf.lCode, "F", hlSbf)
                End If
            End If
        Next ilRvf
    End If
    gUndoInvoice = True
End Function

Private Function mUnbillGameInv(llUndoStartDate As Long, llUndoEndDate As Long, llChfCode As Long, hlMsf As Integer, hlMgf As Integer, hlGsf As Integer) As Integer
    Dim ilGameNo As Integer
    Dim ilRet As Integer
    Dim llGsfDate As Long

    imMsfRecLen = Len(tmMsf)
    imMgfRecLen = Len(tmMgf)
    imGsfRecLen = Len(tmGsf)
    tmMsfSrchKey2.lChfCode = llChfCode
    ilRet = btrGetEqual(hlMsf, tmMsf, imMsfRecLen, tmMsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmMsf.lChfCode = llChfCode)
        ilGameNo = 0
        Do
            tmMgfSrchKey1.lMsfCode = tmMsf.lCode
            tmMgfSrchKey1.iGameNo = ilGameNo
            ilRet = btrGetGreaterOrEqual(hlMgf, tmMgf, imMgfRecLen, tmMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmMgf.lMsfCode = tmMsf.lCode) Then
                If tmMgf.iGameNo > 0 Then
                    tmGsfSrchKey1.lghfcode = tmMsf.lghfcode
                    tmGsfSrchKey1.iGameNo = tmMgf.iGameNo
                    ilRet = btrGetEqual(hlGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    If (ilRet = BTRV_ERR_NONE) Then
                        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llGsfDate
                        If (llGsfDate >= llUndoStartDate) And (llGsfDate <= llUndoEndDate) Then
                            tmMgf.sBilled = "N"
                            ilRet = btrUpdate(hlMgf, tmMgf, imMgfRecLen)
                        End If
                    End If
                Else
                    tmMgf.sBilled = "N"
                    ilRet = btrUpdate(hlMgf, tmMgf, imMgfRecLen)
                End If
                If ilRet = BTRV_ERR_NONE Then
                    ilGameNo = tmMgf.iGameNo + 1
                ElseIf ilRet <> BTRV_ERR_CONFLICT Then
                    mUnbillGameInv = False
                    Exit Function
                End If
            Else
                Exit Do
            End If
        Loop
        ilRet = btrGetNext(hlMsf, tmMsf, imMsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mUnbillGameInv = True
End Function
Private Function mUnbillSpecialBill(llUndoStartDate As Long, llUndoEndDate As Long, llChfCode As Long, slSbfType As String, hlSbf As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long

    imSbfRecLen = Len(tmSbf)
    tmSbfSrchKey0.lChfCode = llChfCode
    gPackDateLong llUndoStartDate, tmSbfSrchKey0.iDate(0), tmSbfSrchKey0.iDate(1)
    tmSbfSrchKey0.sTranType = " "
    ilRet = btrGetGreaterOrEqual(hlSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llChfCode)
        If (tmSbf.sTranType = slSbfType) And (tmSbf.sBilled = "Y") Then
            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
            If (llDate >= llUndoStartDate) And (llDate <= llUndoEndDate) Then
                tmSbf.sBilled = "N"
                ilRet = btrUpdate(hlSbf, tmSbf, imSbfRecLen)
            End If
        End If
        ilRet = btrGetNext(hlSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mUnbillSpecialBill = True
End Function

Private Function mUnbillAirSpots(llUndoStartDate As Long, llUndoEndDate As Long, llChfCode As Long, hlSdf As Integer, hlSmf As Integer, hlClf As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llSdfDate As Long
    Dim llSmfDate As Long
    Dim ilPkLineNo As Integer
    Dim ilUpdateSpot As Integer

    imSdfRecLen = Len(tmSdf)
    imSmfRecLen = Len(tmSmf)
    imClfRecLen = Len(tmClf)
    If (tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "S") Then
        'As Order
        For llDate = llUndoStartDate To llUndoEndDate Step 1
            gPackDateLong llDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
            tmSdfSrchKey4.lChfCode = llChfCode
            ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While ilRet = BTRV_ERR_NONE
                If (tmSdf.sBill = "Y") And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                    If llSdfDate <> llDate Then
                        Exit Do
                    End If
                    If tmSdf.lChfCode <> llChfCode Then
                        Exit Do
                    End If
                    tmSdf.sBill = "N"
                    ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        gPackDateLong llDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
                        tmSdfSrchKey4.lChfCode = llChfCode
                        ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    Else
                        ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                Else
                    ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Loop
        Next llDate
        For llDate = llUndoStartDate To llUndoEndDate Step 1
            gPackDateLong llDate, tmSmfSrchKey4.iMissedDate(0), tmSmfSrchKey4.iMissedDate(1)
            tmSmfSrchKey4.lChfCode = llChfCode
            ilRet = btrGetEqual(hlSmf, tmSmf, imSmfRecLen, tmSmfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llSmfDate
                If llSmfDate <> llDate Then
                    Exit Do
                End If
                If tmSmf.lChfCode <> llChfCode Then
                    Exit Do
                End If
                Do
                    tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tmSdf.sBill = "Y") Then
                        tmSdf.sBill = "N"
                        ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = btrGetNext(hlSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next llDate
    Else
        'As Aired
        For llDate = llUndoStartDate To llUndoEndDate Step 1
            gPackDateLong llDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
            tmSdfSrchKey4.lChfCode = llChfCode
            ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While ilRet = BTRV_ERR_NONE
                If tmSdf.sBill = "Y" Then
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                    If llSdfDate <> llDate Then
                        Exit Do
                    End If
                    If tmSdf.lChfCode <> llChfCode Then
                        Exit Do
                    End If
                    ilUpdateSpot = True
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        'Test if Hidden line and line type is Virtual
                        tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                        tmClfSrchKey0.iLine = tmSdf.iLineNo
                        tmClfSrchKey0.iCntRevNo = 32000
                        tmClfSrchKey0.iPropVer = 32000
                        ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                            If tmClf.sType = "H" Then
                                'Test if Hidden line and line type is Virtual
                                tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                                tmClfSrchKey0.iLine = tmClf.iPkLineNo
                                ilPkLineNo = tmClf.iPkLineNo
                                tmClfSrchKey0.iCntRevNo = tmClf.iCntRevNo
                                tmClfSrchKey0.iPropVer = tmClf.iPropVer
                                ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = ilPkLineNo) Then
                                    If tmClf.sType = "O" Then
                                        'Check Missed date
                                        tmSmfSrchKey2.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hlSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        If ilRet = BTRV_ERR_NONE Then
                                            gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llSmfDate
                                            If (llSmfDate < llUndoStartDate) Or (llSmfDate > llUndoEndDate) Then
                                                ilUpdateSpot = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If ilUpdateSpot Then
                        tmSdf.sBill = "N"
                        ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            gPackDateLong llDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
                            tmSdfSrchKey4.lChfCode = llChfCode
                            ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Else
                            ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                    Else        '7-30-10  prevent endless loop
                        ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                Else
                    ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Loop
        Next llDate
        'Check package spot.  If Virtual, then treat as Ordered
        For llDate = llUndoStartDate To llUndoEndDate Step 1
            gPackDateLong llDate, tmSmfSrchKey4.iMissedDate(0), tmSmfSrchKey4.iMissedDate(1)
            tmSmfSrchKey4.lChfCode = llChfCode
            ilRet = btrGetEqual(hlSmf, tmSmf, imSmfRecLen, tmSmfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llSmfDate
                If llSmfDate <> llDate Then
                    Exit Do
                End If
                If tmSmf.lChfCode <> llChfCode Then
                    Exit Do
                End If
                Do
                    tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tmSdf.sBill = "Y") Then
                        'Test if Hidden line and line type is Virtual
                        tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                        tmClfSrchKey0.iLine = tmSdf.iLineNo
                        tmClfSrchKey0.iCntRevNo = 32000
                        tmClfSrchKey0.iPropVer = 32000
                        ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                            If tmClf.sType = "H" Then
                                'Test if Hidden line and line type is Virtual
                                tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                                tmClfSrchKey0.iLine = tmClf.iPkLineNo
                                ilPkLineNo = tmClf.iPkLineNo
                                tmClfSrchKey0.iCntRevNo = tmClf.iCntRevNo
                                tmClfSrchKey0.iPropVer = tmClf.iPropVer
                                ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = ilPkLineNo) Then
                                    If tmClf.sType = "O" Then
                                        tmSdf.sBill = "N"
                                        ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = btrGetNext(hlSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next llDate
    End If
    mUnbillAirSpots = True
End Function
Private Function mClearPackageSpots(llUndoStartDate As Long, llUndoEndDate As Long, llChfCode As Long, hlPsf As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    imPsfRecLen = Len(tmPsf)
    For llDate = llUndoStartDate To llUndoEndDate Step 1
        Do
            gPackDateLong llDate, tmPsfSrchKey4.iDate(0), tmPsfSrchKey4.iDate(1)
            tmPsfSrchKey4.lChfCode = llChfCode
            ilRet = btrGetEqual(hlPsf, tmPsf, imPsfRecLen, tmPsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilRet = btrDelete(hlPsf)
        Loop
    Next llDate
    mClearPackageSpots = True
End Function



Public Sub gGetAdvanceBillInfo()
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    ReDim tgAdanceBillInfo(0 To 0) As ADVANCEBILLINFO
    If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT) Then
        Exit Sub
    End If
    '3/4/19
    'Temporearily rfemove advance bill
    Exit Sub
    slSQLQuery = "Select Distinct sbfChfCode, Count(distinct sbfDate) as DateCount from SBF_Special_Billing where sbfTranType = 'F' Group By sbfChfCode having DateCount = 1"
    Set sbf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not sbf_rst.EOF
        slSQLQuery = "Select sbfDate from SBF_Special_Billing where sbfTranType = 'F' And sbfBilled <> 'Y' And sbfChfCode = " & sbf_rst!sbfChfCode
        Set tmp_rst = gSQLSelectCall(slSQLQuery)
        If Not tmp_rst.EOF Then
            tgAdanceBillInfo(UBound(tgAdanceBillInfo)).lChfCode = sbf_rst!sbfChfCode
            tgAdanceBillInfo(UBound(tgAdanceBillInfo)).lDate = gDateValue(tmp_rst!sbfDate)
            ReDim Preserve tgAdanceBillInfo(0 To UBound(tgAdanceBillInfo) + 1) As ADVANCEBILLINFO
        End If
        sbf_rst.MoveNext
    Loop
End Sub
Public Sub gGetAdvanceBillMergeInfo(llStartStdDate As Long, llEndStdDate As Long)
    Dim slSQLQuery As String
    Dim ilUpper As Integer
    ReDim tgAdvanceBillMergeInfo(0 To 0) As ADVANCEBILLMERGEINFO
    If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT) Then
        Exit Sub
    End If
    If tgSpf.sBCombine = "N" Then
        Exit Sub
    End If
    slSQLQuery = "Select Distinct rvfInvNo, phfInvNo From rvf_Receivables "
    slSQLQuery = slSQLQuery + "Inner Join phf_Payment_History On rvfCntrNo = phfCntrNo "
    slSQLQuery = slSQLQuery + "Inner Join SBF_Special_Billing on rvfSbfCode = sbfCode "
    slSQLQuery = slSQLQuery + "Where rvfTranDate >= '" + Format(llStartStdDate, sgSQLDateForm) + "' And rvfTranDate <= '" + Format(llEndStdDate, sgSQLDateForm) + "' "
    slSQLQuery = slSQLQuery + " And phfTranDate >= '" + Format(llStartStdDate, sgSQLDateForm) + "' And phfTranDate <= '" + Format(llEndStdDate, sgSQLDateForm) + "' "
    slSQLQuery = slSQLQuery + " And sbfTranType = 'F' "
    Set merge_rst = gSQLSelectCall(slSQLQuery)
    Do While Not merge_rst.EOF
        ilUpper = UBound(tgAdvanceBillMergeInfo)
        tgAdvanceBillMergeInfo(ilUpper).lRvfInvNo = merge_rst!rvfInvNo
        tgAdvanceBillMergeInfo(ilUpper).lPhfInvNo = merge_rst!phfInvNo
        ReDim Preserve tgAdvanceBillMergeInfo(0 To ilUpper + 1) As ADVANCEBILLMERGEINFO
        merge_rst.MoveNext
    Loop
End Sub

Public Sub gAdjOverrideTimes(ilVefCode As Integer, slStartTime As String, slEndTime As String)
    Dim ilTimeAdj As Integer
    Dim ilZone As Integer
    Dim ilVpf As Integer
    Dim llTime As Long
    
    'Remove the Override change for now
    'Exit Sub
    If (tgSpf.sInvSpotTimeZone = "E") Or (tgSpf.sInvSpotTimeZone = "C") Or (tgSpf.sInvSpotTimeZone = "M") Or (tgSpf.sInvSpotTimeZone = "P") Then
        ilTimeAdj = 0
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            'For ilZone = 1 To 5 Step 1
            For ilZone = LBound(tgVpf(ilVpf).sGZone) To UBound(tgVpf(ilVpf).sGZone) Step 1
                If Left$(tgVpf(ilVpf).sGZone(ilZone), 1) = tgSpf.sInvSpotTimeZone Then
                    ilTimeAdj = tgVpf(ilVpf).iGLocalAdj(ilZone)
                    Exit For
                End If
            Next ilZone
        End If
        If ilTimeAdj <> 0 Then
            llTime = gTimeToLong(slStartTime, False) + (CLng(ilTimeAdj) * 3600)
            If llTime < 0 Then
'                llTime = 86400 - llTime
                llTime = 86400 + llTime    '7-17-19 start time adjustment
            ElseIf llTime > 86399 Then
                llTime = llTime - 86400
            End If
            slStartTime = gFormatTimeLong(llTime, "A", "1")
            llTime = gTimeToLong(slEndTime, False) + (CLng(ilTimeAdj) * 3600)
            If llTime < 0 Then
'                llTime = 86400 - llTime
                llTime = 86400 + llTime    '7-17-19 start time adjustment
            ElseIf llTime > 86399 Then
                llTime = llTime - 86400
            End If
            slEndTime = gFormatTimeLong(llTime, "A", "1")
        End If
    End If
End Sub

