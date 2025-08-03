Attribute VB_Name = "modImport"
'******************************************************
'*  modImport - various global declarations for importing
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Public igImportSelection As Integer    '0=Global format; 1= mai format; 2=Log Spots; 3= Affiliate Spots

Type ALT
    lCode              As Long               ' Affiliate Link Table Auto
                                             ' Increment
    lAstCode           As Long               ' Affiliate Spot Reference code
    lLinkToAstCode     As Long               ' Link to Affiliate Spot Table
                                             ' (Defined for MG and Replacement)
    sAiredISCI         As String * 20        ' Aired ISCI if different from what
                                             ' was exported
    iMnfMissed         As Integer            ' Missed Reason
    iAdfCode              As Integer         ' Advertiser reference code.  Used
                                             ' with Marketron import.
    iMissedDate(0 To 1)   As Integer         ' Missed date used with Marketron
                                             ' Import (Default 1/1/1970)
    iMGDate(0 To 1)       As Integer         ' MG Date used with Marketron
                                             ' Import (Default 1/1/1970)
    sUnused            As String * 10
End Type


Type AGREEID
    lCode As Long
    lAgreementID As Long
    iShttCode As Integer
    iVefCode As Integer
    lOnAir As Long
    lOffAir As Long
    lDropDate As Long
    lEndDate As Long
End Type

Type MISSINGATT
    iShttCode As Integer
    iVefCode As Integer
    lCount As Long
End Type

Type LSTMYLINFO
    iType As Integer
    sLogTime As String * 12
    lLogTime As Long
    iHour As Integer
    iLen As Integer
    iWkNo As Integer
    iBreakNo As Integer
    iPositionNo As Integer
    sZone As String * 3
    iAnfCode As Integer
    lCode As Long
    iHourBreakNo As Integer
End Type

Type LOGSPOTINFO
    lCode As Long
    iVefCode As Integer
    lLogTime As Long
    lCntrNo As Long
End Type

Type PledgeInfo
    ilDays(0 To 6) As Integer
    iFdStatus As Integer
    lPdSTime As Long
    lCode As Long
End Type

'    "attCode","Advt","Prod","PledgeStartDate1","PledgeEndDate","PledgeStartTime","PledgeEndTime",
'    "SpotLen","Cart","ISCI","CreativeTitle","astCode","ActualAirDate1","ActualAirTime1","statusCode",
'    "FeedDate","FeedTime","RecType", "MRReason", "OrgAstCode", "NewAstCode"

Type AIRSPOTINFO
    lAtfCode As Long
    sAdvt As String * 30
    sProd As String * 35
    sPledgeStartDate1 As String * 10
    sPledgeEndDate As String * 10
    sPledgeStartTime As String * 10
    sPledgeEndTime As String * 10
    iSpotLen As Integer
    sCart As String * 12
    sISCI As String * 20
    sCreativeTitle As String * 30
    lAstCode As Long
    sActualAirDate1 As String * 10
    sActualAirTime1 As String * 10
    sStatusCode As String * 2
    sFeedDate As String * 10
    sFeedTime As String * 10
    'D.S. 03/08/12 4 new fields below
    'D.S. 06/07/16 Changed sRecType from 1 to 2 bytes
    sRecType As String * 2
    iMissedReason As Integer
    lOrgAstCode As Long
    lNewAstCode As Long
    'end new fields
    sEndDate As String * 10
    iVefCode As Integer
    iShfCode As Integer
    iFound As Integer
    sStartDate As String * 10
    iUpdateComplete As Integer
    bIsciChngFlag As Boolean
    lgsfCode As Long
    sVendorSource As String
End Type

Type MARKETRONSPOTINFO
    iSpotLen As Integer
    sCart As String * 12
    sISCI As String * 20
    sCreativeTitle As String * 30
    lAstCode As Long
    sActualAirDate1 As String * 10
    sActualAirTime1 As String * 8
    sStatusCode As String * 1
    sFeedDate As String * 10
    iFound As Integer
    sSignature As String * 200
End Type

Type BIASTATIONINFO
    sCallLetters As String * 26
    sMarketName As String * 60
    iRank As Integer
    sOwnerName As String * 60
    sFormat As String * 40
End Type

Type BIAREPORTINFO
    sCallLetters As String * 26
    lPermStationNo As Long
    sReportInfo As String * 200
End Type

Type BIAREGIONSET
    sCategory As String * 1
    iFromCode As Integer
    iToCode As Integer
End Type

Type UPDATESTATION
    iCode As Integer
    sCallLetters As String * 10
    lID As Long
    sFrequency As String * 6
    sTerritory As String * 20
    sArea As String * 40
    sFormat As String * 60
    iDMARank As Integer
    sDMARank As String * 4
    sDMAMarket As String * 60
    sCityLicense As String * 40
    sCountyLicense As String * 40
    sStateLicense As String * 2
    sOwner As String * 60
    sOperator As String * 40
    iMSARank As Integer
    sMSARank As String * 4
    sMSAMarket As String * 60
    sMarketRep As String * 20
    sServiceRep As String * 20
    sZone As String * 1
    sOnAir As String * 3
    sCommercial As String * 1
    sDaylight As String * 3
    lXDSStationID As Long
    sXDSStationID As String * 10
    sIPumpID As String * 10
    sSerial1 As String * 10
    sSerial2 As String * 10
    sUsedAgreement As String * 3
    sUsedXDS As String * 3
    sUsedWegener As String * 3
    sUsedOLA As String * 3
    sMoniker As String * 40
    lWatts As Long
    sWatts As String * 12
    sHistoricalDate As String * 10
    sTransactID As String * 5
    lP12Plus As Long
    sP12Plus As String * 12
    sWebAddress As String * 90
    sWebPassword As String * 10
    sMailAddress1 As String * 40
    sMailAddress2 As String * 40
    sMailCity As String * 40
    sMailState As String * 2
    sMailZip As String * 20
    sMailCountry As String * 40
    sPhysicalAddress1 As String * 40
    sPhysicalAddress2 As String * 40
    sPhysicalCity As String * 40
    sPhysicalState As String * 2
    sPhysicalZip As String * 20
    sPhone As String * 20
    sFax As String * 20
    sPersonName(0 To 4) As String * 82
    sPersonTitle(0 To 4) As String * 40
    sPersonPhone(0 To 4) As String * 20
    sPersonFax(0 To 4) As String * 20
    sPersonEMail(0 To 4) As String * 70
    sPersonAffLabel(0 To 4) As String * 3
    sPersonISCIExport(0 To 4) As String * 3
    sPersonAffEMail(0 To 4) As String * 3
End Type
    
Type NEWNAMESIMPORTED
    sNewName As String * 60
    iRank As Integer    'DMA and MSA Market rank
    lUpdateStationIndex As Long
    lReplaceCode As Long
    iCount As Integer
End Type

Public tgNewNamesImported() As NEWNAMESIMPORTED
Public igNewNamesImportedType As Integer '1 = DMA Market; 2 = MSA Market; 3 = Owner; 4= Format
Public igNewNamesImportedReturn As Integer   '0=Cancelled; 1=Ok

Type FORMATLINKINFO
    iCode As Integer
    sExtFormatName As String * 60
    iIntFmtCode As Integer
End Type
