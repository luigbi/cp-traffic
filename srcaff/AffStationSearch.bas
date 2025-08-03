Attribute VB_Name = "modStationSearch"
'******************************************************
'*  modContact - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Public Const MIDGREENCOLOR = &H669933    'RGB(51, 153, 102)
Public Const BROWNCOLOR = 39372  'RGB(204, 153, 00)
Public Const LIGHTGREENCOLOR = &HCCFFCC 'RGB(204, 255, 204)
Public Const ORANGECOLOR = 39423    ' Hex value treated as integer negative number &H99FF       'RGB(255, 153, 0)
Public Const VIOLETCOLOR = &H800080    'RGB(128, 0, 128)
Public Const LIGHTMAGENTACOLOR = &HFFD1F7    'RGB(247, 209, 255)
Public Const LIGHTBLUECOLOR = 16763904   'RGB(0, 204, 255)

'Retain the numbers as those are retains in saved filters
'The value should be one greater than index into tgFilterTypes
Public Const SFAREA = 1
Public Const SFCALLLETTERSCHGDATE = 2
Public Const SFCALLLETTERS = 3
Public Const SFCITYLIC = 4
Public Const SFCOMMERCIAL = 5
Public Const SFCOUNTYLIC = 6
Public Const SFDAYLIGHT = 7
Public Const SFDMA = 8
Public Const SFFORMAT = 9
Public Const SFFREQ = 10
Public Const SFHISTSTARTDATE = 11
Public Const SFPERMID = 12
Public Const SFMAILADDRESS = 13
Public Const SFMARKETREP = 14
Public Const SFMONIKER = 15
Public Const SFMSA = 16
Public Const SFONAIR = 17
Public Const SFOPERATOR = 18
Public Const SFOWNER = 19
Public Const SFP12PLUS = 20
Public Const SFPHONE = 21
Public Const SFPHYADDRESS = 22
Public Const SFSERIAL = 23
Public Const SFSERVICEREP = 24
Public Const SFSTATELIC = 25
Public Const SFEMAIL = 26
Public Const SFISCI = 27
Public Const SFLABEL = 28
Public Const SFPERSONNEL = 29
Public Const SFAGREEMENT = 30
Public Const SFWEGENER = 31
Public Const SFXDS = 32
Public Const SFTERRITORY = 33
Public Const SFZONE = 34
Public Const SFENTERPRISEID = 35
Public Const SFVEHICLEACTIVE = 36
Public Const SFVEHICLEALL = 46
Public Const SFWEBADDRESS = 37
Public Const SFWEBPW = 38
Public Const SFXDSID = 39
Public Const SFZIP = 40
Public Const SFMULTICAST = 41
Public Const SFSISTER = 42
Public Const SFWATTS = 43
Public Const SFDMARANK = 44
Public Const SFMSARANK = 45
'6048
Public Const SFEMAILADDRESS = 47
'
Public Const SFDUE = 48
'5/6/18: Add Vendor
Public Const SFLOGDELIVERY = 49
Public Const SFAUDIODELIVERY = 50
'6/23/19
Public Const SFSERVICEAGREEMENT = 51
'Rules:
'Format or'd with other formats
'Owner or'd with other owners
'Vehicle or'd with other vehicles
'DMA, MSA and Zip or'd with other DMA, MSA, Zip
'Not Equal is And with all
'Contains is And with all
'Any item from one or group above is AND with other or groups
'i.e. Format is AND with Vehicle
'     Format is AND with Owner
'     Vehicle is AND with owner

Type FILTERTYPES
    sFieldName As String * 40
    iSelect As Integer
    sContainAllowed As String * 1
    sEqualAllowed As String * 1
    sRangeAllowed As String * 1
    sNoEqualAllowed As String * 1
    sGreaterOrEqual As String * 1
    sCntrlType As String * 1     'E=Edit Box; L=List Box; T=Toggle
    iCountGroup As Integer  '0=All those field Name that represent geographic area (station, DMA, MSA, Zip, State,..); 1-N=Unique counts like format, Daylight, Market Rep,...
End Type

Type FILTERDEF
    iSelect As Integer  '0=DMA; 1=Format; 2=MSA; 3=Owner; 4=Vehicle; 5=Zip
    iOperator As Integer    '0=Contains; 1=Equal; 2=Not Equal; 3=Range
    lFromValue As Long  'Select: 0; 1; 2; 3; and 4 and and operator <> 0
    sFromValue As String    'Select Select 5 or Operator 0
    lToValue As Long  'Select: 0; 1; 2; 3; and 4 and and operator <> 0
    sToValue As String    'Select Select 5 or Operator 0
    lFitCode As Long
    iCountGroup As Integer  '0=All those field Name that represent geographic area (station, DMA, MSA, Zip, State,..); 1-N=Unique counts like format, Daylight, Market Rep,...
                            'Same group treats as OR's, different groups treated as AND's
    iFirstFilterLink As Integer
    sCntrlType As String * 1
End Type

Type FILTERLINK
    lFilterDefIndex As Long
    lNotFilterDefIndex As Long
    lNextAnd As Long
End Type

'5/6/18: Set to max number of filters defined in mPopFilterTypes
Public tgFilterTypes(0 To 51) As FILTERTYPES

Public tgFilterDef() As FILTERDEF
'Public tgNotFilterDef() As FILTERDEF 'Contains and the Not Equals
Public igFilterReturn As Integer    'True = Done pressed, False=Cancel pressed
Public sgFilterName As String
Public lgFhtCode As Long
Public igFilterChgd As Integer

Type FILTERCOUNT
    iCount As Integer
    iType As Integer    '0=Format; 1=Owner; 2=Vehicle; 3=DMA or MSA or Zip
End Type

'6/23/19
Public bgServiceAgreementExist As Boolean
Public bgIncludeServiceAgreement As Boolean

Public igContactEmailShttCode As Integer
Public igCommentShttCode() As Integer

Private smWeek1 As String
Private smWeek54 As String
Private rst_Cptt As ADODB.Recordset


Public Sub gBuildStationCount(slWeek1 As String, slWeek54 As String)
    Dim slCPTTSQLQuery As String
    Dim slEndWeek As String
    Dim ilShttCode As Integer
    Dim ilCount As Integer
    
    If (gDateValue(slWeek1) = gDateValue(smWeek1)) And (gDateValue(slWeek54) = gDateValue(smWeek54)) And (Not gFileChgd("cptt.mkd")) Then
        Exit Sub
    End If
    slEndWeek = DateAdd("d", -14, slWeek1)
    ReDim tgStationCount(0 To 10000) As SHTTINFO1
    ilCount = 0
    ilShttCode = -1
    slCPTTSQLQuery = "Select cpttShfCode, cpttVefcode, Count(*) as Due from CPTT "
    slCPTTSQLQuery = slCPTTSQLQuery & " Where cpttPostingStatus <= 1 and "
    slCPTTSQLQuery = slCPTTSQLQuery & " cpttStartDate >= '" & Format(slWeek54, sgSQLDateForm) & " ' and cpttStartdate <= '" & Format(slEndWeek, sgSQLDateForm) & "'"
    slCPTTSQLQuery = slCPTTSQLQuery & " Group By cpttShfCode, cpttVefCode Order By cpttShfCode, Due Desc"
    'Set rst_Cptt = cnn.Execute(slCPTTSQLQuery)
    Set rst_Cptt = gSQLSelectCall(slCPTTSQLQuery)
    Do While Not rst_Cptt.EOF
        If (ilShttCode <> rst_Cptt!cpttshfcode) Then
            tgStationCount(ilCount).shttCode = rst_Cptt!cpttshfcode
            tgStationCount(ilCount).shttTimeZone = rst_Cptt!Due
            ilShttCode = rst_Cptt!cpttshfcode
            ilCount = ilCount + 1
            If ilCount > UBound(tgStationCount) Then
                ReDim Preserve tgStationCount(0 To ilCount + 1000) As SHTTINFO1
            End If
        End If
        rst_Cptt.MoveNext
    Loop
    ReDim Preserve tgStationCount(0 To ilCount) As SHTTINFO1
    smWeek1 = slWeek1
    smWeek54 = slWeek54
    gFileChgdUpdate "cptt.mkd", False
End Sub

Public Function gBinarySearchStationCount(ilCode As Integer) As String
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    gBinarySearchStationCount = ""
    llMin = LBound(tgStationCount)
    llMax = UBound(tgStationCount) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tgStationCount(llMiddle).shttCode Then
            'found the match
            If Val(tgStationCount(llMiddle).shttTimeZone) > 0 Then
                gBinarySearchStationCount = Trim$(tgStationCount(llMiddle).shttTimeZone)
            End If
            Exit Function
        ElseIf ilCode < tgStationCount(llMiddle).shttCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    Exit Function
    
End Function
