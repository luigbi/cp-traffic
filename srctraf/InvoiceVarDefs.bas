Attribute VB_Name = "InvoiceVarDefs"


' Proprietary Software, Do not copy
'
' File Name: InvoiceVarDef.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice support functions
Option Explicit
Option Compare Text

Public hmMsg As Integer    'd.s. 11/6/01

Public tmChf As CHF
Public hmCHF As Integer            'Contract header file handle
Public tmChfSrchKey As LONGKEY0            'CHF record image
Public tmChfSrchKey1 As CHFKEY1            'CHF record image
Public imCHFRecLen As Integer        'CHF record length

Public lmSelCntrCode() As Long     'Selected contract codes
Public tmAirNTRCombine() As AIRNTRCOMBINE

Public tmInstallCntStatus() As INVCNTRSTATUS
Public imInstallStatusConflict As Integer
Public lmInstallCntrNo As Long
Public imInstallInvoiceNo As Integer
Public imTotalInstallInvs As Integer

Public smSvStartStd As String    'Starting date for standard billing
Public smSvEndStd As String      'Ending date for standard billing
Public smSvStartCal As String    'Starting date for standard billing
Public smSvEndCal As String      'Ending date for standard billing
Public smSvStartWk As String    'Starting date for standard billing
Public smSvEndWk As String      'Ending date for standard billing
Public lmSvStartStd As Long    'Starting date for standard billing
Public lmSvEndStd As Long      'Ending date for standard billing
Public lmSvStartCal As Long    'Starting date for standard billing
Public lmSvEndCal As Long      'Ending date for standard billing
Public lmSvStartWk As Long    'Starting date for standard billing
Public lmSvEndWk As Long      'Ending date for standard billing

Public smStartStd As String    'Starting date for standard billing
Public smEndStd As String      'Ending date for standard billing
Public smStartCal As String    'Starting date for standard billing
Public smEndCal As String      'Ending date for standard billing
Public smStartWk As String    'Starting date for standard billing
Public smEndWk As String      'Ending date for standard billing
Public lmStartStd As Long    'Starting date for standard billing
Public lmEndStd As Long      'Ending date for standard billing
Public lmStartCal As Long    'Starting date for standard billing
Public lmEndCal As Long      'Ending date for standard billing
Public lmStartWk As Long    'Starting date for standard billing
Public lmEndWk As Long      'Ending date for standard billing

Public lmNTRDate As Long

Public imArfPDFEMailCode As Integer

'Invoice Report from Crystal
Public hmIvr As Integer            'Invoice Report file handle
Public tmIvr As IVR                'IVR record image
Public tmSvIvr As IVR
Public imIvrRecLen As Integer      'IVR record length
Public tmIvrSrchKey1 As LONGKEY0            'SDF record image
Public hmImr As Integer
Public tmImr As IMR
Public imImrRecLen As Integer
Public imCombineAirAndNTR As Integer
Public imAppendFutureSpots As Integer

Public bmEDIInstallBypassed As Boolean
Public lmEDIInstallBypass() As Long

Public bmEDINTRBypassed As Boolean
Public lmEDINTRBypass() As Long

Public bmEDICPMBypassed As Boolean
Public lmEDICPMBypass() As Long

Public hmPcf As Integer            'Special Billing file handle
Public tmPcf As PCF                'SBF record image
Public tmPcfSrchKey0 As LONGKEY0
Public tmPcfSrchKey1 As PCFKEY1
Public tmPcfSrchKey2 As PCFKEY2
Public tmPcfSrchKey3 As PCFKEY3
Public imPcfRecLen As Integer       'SBF record length
Public hmPcfVehCheck As Integer


Public hmIbf As Integer            'Special Billing file handle
Public tmIbf As IBF                'SBF record image
Public tmIbfSrchKey0 As LONGKEY0
Public tmIbfSrchKey1 As IBFKEY1
Public tmIbfSrchKey2 As IBFKEY2
Public tmIbfSrchKey3 As IBFKEY3
Public imIbfRecLen As Integer       'SBF record length

Public lmCPMBypassCntr() As Long    'Array of chfCode for CPM buys missing posted impressions

'5-24-13 Vehicle Features - get 4 line of address when using Vehicle as Lockbox
Public hmVff As Integer             'file handle
Public tmVff As VFF                 'record image
Public tmVffSrchKey1 As INTKEY0
Public imVffRecLen As Integer       ' record length

'Virtual Vehicle
Public hmVsf As Integer             'Virtual Vehicle file handle
Public tmVsf As VSF                 'VSF record image
Public tmVsfSrchKey As LONGKEY0             'VSF record image
Public imVsfRecLen As Integer         'VSF record length
'Vehicle Links
Public hmVLF As Integer             'Vehicle links file handle
Public imVlfRecLen As Integer       'Vehicle links record length
Public tmVlf0() As VLF              'Mon-Fri vehicle links
Public tmVlf6() As VLF              'Sat vehicle links
Public tmVlf7() As VLF              'Sun vehicle links
Public tmVlf() As VLF
'Library calendar
Public hmLcf As Integer         'Library calendar file handle
Public tmLcf As LCF
Public imLcfRecLen As Integer
Public tmLcfSrchKey As LCFKEY0
Public tmLcfSrchKey2 As LCFKEY2
Public hmLef As Integer         'Library Event file handle
Public tmLef As LEF
Public imLefRecLen As Integer
Public tmLefSrchKey As LEFKEY0

Public hmLvf As Integer         '3-23-12 need game length
Public tmLvf As LVF
Public imLvfRecLen As Integer
Public tmLvfSrchKey0 As LONGKEY0

Public tmSchChf As CHF
Public tmSchClf() As CLFLIST
Public tmSchCff() As CFFLIST
Public tmSchCgf() As CGFLIST
Public tmSchMsf() As MSFLIST
Public tmAlterChf As CHF
Public tmAlterClf() As CLFLIST
Public tmAlterCff() As CFFLIST
Public tmAlterCgf() As CGFLIST
Public tmAlterMsf() As MSFLIST

Public tmCxf As CXF            'CXF record image
Public tmCxSrchKey As LONGKEY0  'CXF key record image
Public hmCxf As Integer        'CXF Handle
Public imCxfRecLen As Integer      'CXF record length
Public smHeaderComment As String
Public hmAdf As Integer            'Advertsier name file handle
Public tmAdf As ADF                'ADF record image
Public tmAdfSrchKey As INTKEY0            'ADF record image
Public imAdfRecLen As Integer        'ADF record length
Public hmAgf As Integer            'Advertsier name file handle
Public tmAgf As AGF                'ADF record image
Public tmAgfSrchKey As INTKEY0            'ADF record image
Public imAgfRecLen As Integer        'ADF record length
Public hmPnf As Integer            'Personnel file handle
Public tmPnf As PNF                'PNF record image
Public tmPnfSrchKey As INTKEY0            'PNF record image
Public imPnfRecLen As Integer        'PNF record length
'PDF Invoice email address
Public hmPDF As Integer            'PDF Email file handle
Public tmPdf As PDF                'PDF Email record image
Public tmPDfSrchKey As INTKEY0            'PDF record image
Public imPdfRecLen As Integer        'PDF record length

Public hmSof As Integer            'Sales Office name file handle
Public tmSof As SOF                'SOLF record image
Public tmSofSrchKey As INTKEY0            'SOF record image
Public imSofRecLen As Integer        'SOF record length
Public hmSlf As Integer            'Salesperson name file handle
Public tmSlf As SLF                'SLF record image
Public tmSlfSrchKey As INTKEY0            'SLF record image
Public imSlfRecLen As Integer        'SLF record length
Public hmRdf As Integer            'Rate card program/time file handle
Public tmRdf As RDF                'RDF record image
Public imRdfRecLen As Integer        'RDF record length
Public hmPrf As Integer            'Product file handle
Public tmPrfSrchKey As PRFKEY1            'PRF record image
Public imPrfRecLen As Integer        'PRF record length
Public tmPrf As PRF
'Copy rotation record information
Public hmCrf As Integer        'Copy rotation file handle
Public tmCrfSrchKey0 As LONGKEY0 'CRF key record image
'Public tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Public tmCrfSrchKey4 As CRFKEY4 'CRF key record image
Public imCrfRecLen As Integer  'CRF record length
Public tmCrf As CRF            'CRF record image
'Copy instruction record information
Public hmCnf As Integer        'Copy instruction file handle
Public tmCnfSrchKey As CNFKEY0 'CNF key record image
Public imCnfRecLen As Integer  'CNF record length
Public tmCnf As CNF            'CNF record image
'Copy Vehicle
Public hmCvf As Integer        'Copy inventory file handle
'Copy inventory
Public hmCif As Integer        'Copy inventory file handle
Public tmCif As CIF            'CIF record image
Public tmCifSrchKey As LONGKEY0 'CIF key record image
Public imCifRecLen As Integer     'CIF record length
' Copy Combo Inventory File
Public hmCcf As Integer        'Copy Combo Inventory file handle
Public tmCcf As CCF            'CCF record image
Public imCcfRecLen As Integer     'CCF record length
'  Copy Product/Agency File
Public hmCpf As Integer        'Copy Product/Agency file handle
Public tmCpf As CPF            'CPF record image
Public tmCpfSrchKey As LONGKEY0 'CPF key record image
Public imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Public hmTzf As Integer        'Time Zone Copy file handle
Public tmTzf As TZF            'TZF record image
Public tmTzfSrchKey As LONGKEY0 'TZF key record image
Public imTzfRecLen As Integer     'TZF record length
Public hmMnf As Integer            'MultiName file handle
Public tmMnf As MNF                'MNF record image
Public tmMnfSrchKey As INTKEY0            'MNF record image
Public imMnfRecLen As Integer        'MNF record length
Public hmArf As Integer            'Name/Address file handle
Public tmArf As ARF                'ARF record image
Public tmEDIArf As ARF                'ARF record image
Public tmArfSrchKey As INTKEY0            'ARF record image
Public imArfRecLen As Integer        'ARF record length

Public hmSbf As Integer            'Special Billing file handle
Public hmSbfVehCheck As Integer
Public tmSbf As SBF                'SBF record image
Public tmSbfSrchKey0 As SBFKEY0
Public tmSbfSrchKey1 As LONGKEY0
Public tmSbfSrchKey2 As SBFKEY2
Public tmSbfSrchKey3 As SBFKEY3
Public imSbfRecLen As Integer       'SBF record length

Public hmSmf As Integer            'MG and outside Times file handle
Public tmSmf As SMF                'SMF record image
Public tmSmfSrchKey1 As LONGKEY0            'SMF record image
Public tmSmfSrchKey2 As LONGKEY0            'SMF record image
Public tmSmfSrchKey4 As SMFKEY4            'SMF record image
Public imSmfRecLen As Integer        'SMF record length


Public hmEDI As Integer   'From file hanle
Public smEDIFile As String
Public smEDIByVehSort() As SORTCODE
Public smEDIRecords() As String   'Create for a contract then written and cleared
'4/3/12: require totals by service
Public imEDIIndex As Integer


Public hmRvf As Integer            'Receivable file handle
Public tmRvf As RVF                'RVF record image
Public tmRvfSrchKey1 As RVFKEY1            'RVF record image (Advertiser code)
Public tmRvfSrchKey2 As LONGKEY0            'RVF record image (Advertiser code)
Public tmRvfSrchKey3 As RVFKEY3            'RVF record image (Advertiser code)
Public tmRvfSrchKey4 As RVFKEY4            'RVF record image (Advertiser code)
Public tmRvfSrchKey5 As RVFKEY5
Public imRvfRecLen As Integer        'RVF record length
Public tmPkRvf() As RVFVEF

Public hmPhf As Integer            'Receivable History file handle- same structure as rvf

Public tmAcqWithCommInfo() As ACQWITHCOMMINFO

Public tmPDFAdf As ADF
Public tmPDFAgf As AGF


Public smInvSpotTimeZone As String

Public hmInvExportSpots As Integer     '5-11-17 Export filename for spot file (using invexportparameters feature)
Public hmInvExportNTR As Integer       '5-11-17 Export filename for NTR file (using invexportparameters feature)


Public tmRPInfo() As RPINFO        'Reprint Information
Public lmRPTax1 As Long
Public lmRPTax2 As Long
Public tmRPSelInfo() As RPSELINFO
