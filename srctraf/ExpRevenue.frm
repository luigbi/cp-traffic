VERSION 5.00
Begin VB.Form ExpRevenue 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   7095
   Begin VB.ComboBox cbcVehGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3810
      TabIndex        =   15
      Top             =   1110
      Width           =   1500
   End
   Begin VB.TextBox lacVehGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2505
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Vehicle Group"
      Top             =   1185
      Width           =   1215
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1125
      MaxLength       =   9
      TabIndex        =   13
      Top             =   1125
      Width           =   1185
   End
   Begin VB.OptionButton rbcYearType 
      Caption         =   "Standard"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1470
      TabIndex        =   11
      Top             =   825
      Width           =   1350
   End
   Begin VB.OptionButton rbcYearType 
      Caption         =   "Corporate"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5595
      Top             =   360
   End
   Begin VB.TextBox edcSelCFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1275
      MaxLength       =   3
      TabIndex        =   0
      Top             =   435
      Width           =   615
   End
   Begin VB.TextBox edcSelCFrom1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2745
      MaxLength       =   4
      TabIndex        =   1
      Top             =   435
      Width           =   615
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6390
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5670
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6045
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   2355
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Top             =   2355
      Width           =   1050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contract #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1185
      Width           =   915
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   75
      Width           =   1380
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2145
      TabIndex        =   8
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Month"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   420
      TabIndex        =   6
      Top             =   1770
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   2310
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "ExpRevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file handle
Dim lmSingleCntr As Long
Dim smAirOrder As String            'billing type from site
Dim imMajorVG As Integer
Dim imMinorVG As Integer
Dim imVGGroupSelected As Integer        'vehicle group index selected fromlist box

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control

Dim tmAdjust() As ADJUSTLIST         'list of MGs $ and vehicles moved to
Dim imUpperAdjust As Integer         'running count of count of MGs built per contract

Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF

Dim tmCntAllYear() As ALLPIFPCTYEAR      'all participant % for all vehicles for a contract for 12 months (1 or more vehicles each could have
                                         '1 or more participants
Dim tmOneVehAllYear() As ALLPIFPCTYEAR       'ss mnf code, mnfgroup, 1 year percentages for 1 vehicle (1 or more participants)
Dim tmOnePartAllYear As ONEPARTYEAR      'ss mnf code, mnfgroup, 1 years percentages for 1 participant

Dim tmSofList() As SOFLIST

Dim hmMnf As Integer            'List file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey0 As INTKEY0     'MNF key image
Dim tmMnf As MNF
Dim tmNTRMNF() As MNF           'NTR types
Dim tmVGMNF() As MNF            'vehicle groups

Dim hmVef As Integer            'Vehicle file handle
Dim imVefRecLen As Integer      'VEF record length
Dim tmVef As VEF

Dim hmVsf As Integer            'Vehicle options file handle
Dim imVsfRecLen As Integer      'VSF record length
Dim tmVsf As VSF


Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlf As SLF

Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim imAdfRecLen As Integer        'ADF record length

Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer
Dim tmAgfSrchKey0 As INTKEY0

Dim hmRvf As Integer            'Receivables file handle
Dim tmTRvf() As RVFCNTSORT
Dim tmRvf As RVF                '


Dim hmPrf As Integer            'Product Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length

Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey As SDFKEY0     'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0   'SDF record image (SDF code as keyfield)

Dim hmSmf As Integer            'MG and outside Times file handle
Dim tmSmf As SMF                'Spots MG record image
Dim tmSmfSrchKey As SMFKEY0     'SMF record image
Dim imSmfRecLen As Integer      'SMF record length

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer      'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmSof As Integer            'Sales Office line file handle
Dim tmSofSrchKey As INTKEY0     'SOF record image
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF

Dim hmSbf As Integer            'Special billing (NTR) file handle
Dim imSbfRecLen As Integer      'SBF record length
Dim tmSbf As SBF

Dim lmSlfSplit() As Long           '4-20-00 slsp slsp share %
Dim imSlfCode() As Integer             '4-20-00

'Index zero ignored in all the below arrays
Dim lmSlspRemGross(0 To 12) As Long 'remaining $ after all slsp has been split, last one gets pennies
Dim lmSlspRemNet(0 To 12) As Long
Dim lmSlspRemTNet(0 To 12) As Long

'Index zero ignored in all the below arrays
Dim lmGrossDollar(0 To 12) As Long  'calc $ for 1 slsp/participant
Dim lmNetDollar(0 To 12) As Long
Dim lmTNetDollar(0 To 12) As Long

Dim lmPartRemGross(0 To 12) As Long
Dim lmPartRemNet(0 To 12) As Long
Dim lmPartRemTNet(0 To 12) As Long

Dim lmPartGross(0 To 12) As Long
Dim lmPartNet(0 To 12) As Long
Dim lmPartTNet(0 To 12) As Long

Dim lmProjectGross(0 To 12) As Long
Dim lmProjectNet(0 To 12) As Long
Dim lmProjectTNet(0 To 12) As Long

Dim tmSBFAdjust() As ADJUSTLIST

Dim tmExportInfo As REVENUEEXPORT
Private Type REVENUEEXPORT          'format of output record
    sContract As String * 9
    sAdvtName As String * 30
    sProduct As String * 35
    sAgency As String * 40
    sVehicle As String * 40
    sVehicleGroup As String * 20
    sSlsp As String * 40
    sOffice As String * 20
    sSalesSource As String * 20
    sParticipant As String * 20
    sCashTrade As String * 1            'c = cash , t = trade
    sAirTimeNTR As String * 8           'a = airtime, n = NTR , H = NTR hard cost
    sPolitical As String * 1            'Y = political, N = non-polit
    sGross(0 To 12) As String * 15  'Index zero ignored
    sNet(0 To 12) As String * 15      'Index zero ignored
    sNetNet(0 To 12) As String * 15   'Index zero ignored
    sType As String * 20                'contract type:  standard, remnant, PI, DR, etc.
End Type

Dim tmInputInfo As REVENUEINPUT     'info required to create an export record
Private Type REVENUEINPUT
    lCashGross(0 To 12) As Long         '1 year of gross, Index zero ignored
    lCashNet(0 To 12) As Long           '1 year of net, Index zero ignored
    lCashTNet(0 To 12) As Long          '1 year of triple net, Index zero ignored
    lTradeGross(0 To 12) As Long         '1 year of gross, Index zero ignored
    lTradeNet(0 To 12) As Long           '1 year of net, Index zero ignored
    lTradeTNet(0 To 12) As Long          '1 year of triple net, Index zero ignored
    iSlfCode As Integer
    iVefCode As Integer             'vehicle code
    iMnfCode As Integer             'participant code
    sCashTrade As String * 1        'cash/trade flag (from rec or contract)
    sAirTimeNTR As String * 1       'A = Air time, N = NTR, H = hardcost
    iVefGroup As Integer            'vehicle group code
End Type

Const LBONE = 1

'
'        verify input parameters for Year, Month
'        <input> ilMonth - month # entered (if text entered, converted to month inx)
'        <return> igMonth - relative month index to start of year (corporate starts
'                 with month other than Jan)
Private Sub mGetAdjustedMonth(ilMonth As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slMonth                                                 *
'******************************************************************************************

Dim slStr As String
Dim ilSaveMonth As Integer
Dim slMonthInYear As String * 36

        slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"


        If rbcYearType(0).Value Then                         'corporate months
            'convert the month name to the correct relative month # of the corp calendar
            'i.e. if 10 entered and corp calendar starts with oct, the result will be july (10th month of corp cal)
            slStr = edcSelCFrom.Text
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth <> 0 Then                           'input is text month name,
                slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                igMonthOrQtr = gGetCorpMonthNoFromMonthName(slMonthInYear, slStr)         'getmonth # relative to start of corp cal
            Else
                igMonthOrQtr = Val(slStr)
            End If

        Else
            igMonthOrQtr = Val(ilMonth)                       'put month entered in global variable
        End If
    Exit Sub
End Sub
'
'                   mAdjustMissedMGs - Subtract missed spots (missed, cancel, hidden)
'                                Count MGs where they air
'                               Billed and booked report
'                   <input> llstdstartDates() - array of 13 start dates of the 12 months to gather
'                           ilFirstProjInx - index of first month to start projection (earlier is from receivables)
'                           llStartAdjust -  Earliest date to start searching for missed, etc.
'                           llEndAdjust - latest date to stop searchng for missed, etc.
'                           ilSubMissed - true if subtract $ from missed vehicle for missed, cancel, hidden spots
'                           ilCountMGs - true if subtract $ from missed vehicle, move $ to makegood vehicle
'                           ilHowManyPer - # of periods to gather
'                                           (for billed & booked its 12,
'                                           for Sales Comparisons its max 3)
'                           ilAdjustMissedForMG - true to subtract out the missed part of the mg,
'                                       false to ignore the missed part of the MG.
'                                       When gathering spots, accumlating spot $ will be short if
'                                       the missed part is subtracted out
'
'       3/97 dh Comment out code to test the MNF Missed Reasons - field contains
'               value whether or not to bill the mg, bill the missed, etc.
'       3-15-00 Speed up the gathering of SMF spots; obtain the flight instead of going through the
'               generalized routines (which reads & rereads SMF) and could also have caused looping when
'               an SMF error existed.
'               Also, when retrieving the SMF, since the key is missed date, start searching for the spots
'               based on the start of the reporting period minus 2 months, as opposed to the beginning (zero).
'       8-17-06 count mg where they air is not checked for air time billing.  It was always included
'       11-12-07 add flag to test to adjust the missed portion of the makegood
Private Sub mAdjustMissedMGs(llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilSubMissed As Integer, ilCountMGs As Integer, ilHowManyPer As Integer, ilAdjustMissedForMG As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilListIndex                   ilFoundOption             *
'*                                                                                        *
'******************************************************************************************

Dim slDate As String
ReDim ilDate(0 To 1) As Integer           'converted date for earliest start date for sdf keyread
Dim ilRet As Integer
Dim llDate As Long
Dim ilMonthInx As Integer
Dim ilFoundMonth As Integer
Dim ilDoAdjust As Integer
Dim slPrice As String           'rate from flight
Dim llActPrice As Long          'rate from flight
Dim ilFoundVef As Integer
Dim ilTemp As Integer
'   rules to subtract missed spots
'
'                     Case 1 "S"        Case 2 "O"              Case 3 "A"
'                   Bill as Order    Bill as Order             As Aired
'                   update Order      Update aired             Update aired
'Package Lines      ignore missed    Ignore missed             Ignore missed
'Hidden Lines       ignore missed    Ignore missed             ignore missed
'Standard Lines     ignore missed    Answer from user input    Answer from user input
'
'If tmClf.sType <> "O" And tmClf.sType <> "A" And ilSubMissed And smAirOrder <> "S" Then   'possibly adjust missed for standard lines                             'subtract missed, cancel, hidden spots?
                                                        'hidden & package lines should never subtract missed spots
If tmClf.sType = "S" And ilSubMissed And smAirOrder <> "S" Then   'possibly adjust missed for standard lines                             'subtract missed, cancel, hidden spots?
    tmSdfSrchKey.iVefCode = tmClf.iVefCode
    tmSdfSrchKey.lChfCode = tmChf.lCode
    tmSdfSrchKey.iLineNo = tmClf.iLine
    tmSdfSrchKey.lFsfCode = 0
    slDate = Format$(llStartAdjust, "m/d/yy")
    gPackDate slDate, ilDate(0), ilDate(1)
    tmSdfSrchKey.iDate(0) = ilDate(0)
    tmSdfSrchKey.iDate(1) = ilDate(1)
    tmSdfSrchKey.sSchStatus = ""
    tmSdfSrchKey.iTime(0) = 0
    tmSdfSrchKey.iTime(1) = 0
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmClf.iVefCode) And (tmSdf.lChfCode = tmChf.lCode) And (tmSdf.iLineNo = tmClf.iLine)
        'If (tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "U" Or tmSdf.sSchStatus = "R" Or slPass = "H") Then
        If (tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "U" Or tmSdf.sSchStatus = "R" Or tmSdf.sSchStatus = "H") Then
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            llDate = gDateValue(slDate)
            If llDate > llEndAdjust Then
                Exit Do
            End If
            'spot is OK, adjust the $
            ilFoundMonth = False
            For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthInx
            If ilFoundMonth Then
                'Temporarily commented out to bypass testing the Missed reason flags for billing,
                'Always go thru the adjustments
                 ilDoAdjust = True
                'found a month that it falls within, should the missed be billed?
                'ilDoAdjust = False
                'For ilLoop = LBound(tmMnfList) To UBound(tmMnfList) Step 1
                '    If ((tmMnfList(ilLoop).iMnfCode = tmSdf.iMnfMissed And tmMnfList(ilLoop).iBillMissMG <= 1) Or (tmSdf.iMnfMissed = 0)) Then
                '    '1=bill mg, nc missed , 0 = nothing answered, default to same as 1
                '        ilDoAdjust = True
                '        Exit For
                '    End If
                'Next ilLoop
            End If
            If ilDoAdjust Then
                ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)

                If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                    'lmProjectTNet(ilMonthInx) = lmProjectTNet(ilMonthInx) - tmClf.lAcquisitionCost     'subtr missed $, since they won't be invoiced
                    llActPrice = gStrDecToLong(slPrice, 2)
                    lmProjectGross(ilMonthInx) = lmProjectGross(ilMonthInx) - llActPrice     'subtr missed $, since they won't be invoiced
                End If

            End If
        End If                          'sschstatus = C, M, H, U, R
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop                                        'while BTRV_err none and contracts & lines match
End If
'Rules for Count mgs where they air
'                     Case 1 "S"        Case 2 "O"              Case 3 "A"
'                   Bill as Order    Bill as Order             As Aired
'                   update Order      Update aired             Update aired
'Package Lines      ignore MGs       Ignore mgs                Same as Case 2 (hidden Line)
'Hidden Lines       ignore mgs       Count MGs where they      Same as Case 2 (hidden line)
'                                    air($ sub from month of
'                                    missed spots vehicle and
'                                    moved to the msised month
'                                    of the mgs spots vehicle.
'Standard Lines     ignore mgs       Same as case 2 except      Same as Case 2 (std Line)
'                                    $ in month vehicle moved to.
If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "O" And tmClf.sType = "S" And ilCountMGs) Or (smAirOrder = "A" And ilCountMGs) Then        '8-17-06 test to count mg where they air for aired billing
    'Do the "Outs" and "MGs"
    tmSmfSrchKey.lChfCode = tmChf.lCode
    tmSmfSrchKey.iLineNo = tmClf.iLine
    slDate = Format$(llStartAdjust - 60, "m/d/yy")   'cant use start of report period because when looking for the SMF by missed date key, the missed spot
                                                    'could be prior to the reporting period , so back it up 2 months
    gPackDate slDate, ilDate(0), ilDate(1)         '3-13-01
    tmSmfSrchKey.iMissedDate(0) = ilDate(0)
    tmSmfSrchKey.iMissedDate(1) = ilDate(1)
    'tmSmfSrchKey.iMissedDate(0) = 0               '3-13-01 use start of period minus 2 months to adjust for missed spots
    'tmSmfSrchKey.iMissedDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmChf.lCode) And (tmSmf.iLineNo = tmClf.iLine)
        'test dates later in SDF
        'Find associated spot in SDF
        If (tmSmf.sSchStatus = "O" Or tmSmf.sSchStatus = "G") Then
            tmSdfSrchKey3.lCode = tmSmf.lSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
            If (ilRet = BTRV_ERR_NONE) Then
                'spot is OK, assume to adjust $ to where spot was aired  (as aired billing)
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate

                'use Month spot moved to, or month missed from?
                If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "A" And tmClf.sType <> "S") Then
                    'use original missed date
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                End If
                llDate = gDateValue(slDate)
                ilFoundMonth = False
                For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                    If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                        ilFoundMonth = True
                        Exit For
                    End If
                Next ilMonthInx
                ilDoAdjust = False
                If ilFoundMonth Then
                    'Temporarily commented out to bypass testing the Missed reason flags for billing,
                    'Always go thru the adjustments
                     ilDoAdjust = True

                    'found a month that it falls within, should the mg be billed?
                    'ilDoAdjust = False
                    'For ilLoop = LBound(tmMnfList) To UBound(tmMnfList) Step 1
                    '    If ((tmMnfList(ilLoop).iMnfCode = tmSdf.iMnfMissed) And (tmMnfList(ilLoop).iBillMissMG <= 1 Or tmMnfList(ilLoop).iBillMissMG = 3)) Or (tmSdf.iMnfMissed = 0) Then
                    '    '1=bill mg, nc missed  , 3 = bill both missed & mg
                    '        ilDoAdjust = True
                    '        Exit For
                    '    End If
                    'Next ilLoop
                End If
                If ilDoAdjust Then
                    ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)

                    If (InStr(slPrice, ".") <> 0) Then         'found spot cost and its a selected vehicle

                        ilFoundVef = False
                        'setup vehicle that spot was moved to
                        For ilTemp = LBound(tmAdjust) To UBound(tmAdjust) - 1 Step 1
                            If tmAdjust(ilTemp).iVefCode = tmSdf.iVefCode Then
                                ilFoundVef = True
                                Exit For
                            End If
                        Next ilTemp
                        If Not (ilFoundVef) Then
                            tmAdjust(imUpperAdjust).iVefCode = tmSdf.iVefCode
                            ilTemp = imUpperAdjust
                            imUpperAdjust = imUpperAdjust + 1
                            ReDim Preserve tmAdjust(0 To imUpperAdjust) As ADJUSTLIST

                        End If

                        'mg $ - if ordered (update ordered or aired), put mg in same month it was ordered
                        'if update as aired, put mg where it ran
                        llActPrice = gStrDecToLong(slPrice, 2)
                        'tmAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + tmClf.lAcquisitionCost
                        tmAdjust(ilTemp).lProject(ilMonthInx) = tmAdjust(ilTemp).lProject(ilMonthInx) + llActPrice    'add back in the mg that is invoiced
                        'now do the missed portion of the mg
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                        llDate = gDateValue(slDate)
                        ilFoundMonth = False
                        For ilMonthInx = 1 To ilHowManyPer Step 1         'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                ilFoundMonth = True
                                Exit For
                            End If
                        Next ilMonthInx
                        If ilFoundMonth And ilAdjustMissedForMG Then    'subtract out the missed portion of the mg?
                            If ilMonthInx >= ilFirstProjInx Then            'only adjust if its in the future
                                'lmProjectTNet(ilMonthInx) = lmProjectTNet(ilMonthInx) - tmClf.lAcquisitionCost
                                lmProjectGross(ilMonthInx) = lmProjectGross(ilMonthInx) - llActPrice    'adjust the missed portion of the mg

                            End If
                        End If
                    End If
                Else                    'month not found for mg, find missed part of adjustment
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                    llDate = gDateValue(slDate)
                    ilFoundMonth = False
                    For ilMonthInx = 1 To ilHowManyPer Step 1         'loop thru months to find the match
                        If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                            ilFoundMonth = True
                            Exit For
                        End If
                    Next ilMonthInx
                    If ilFoundMonth And ilAdjustMissedForMG Then    'subtract out the missed portion of the mg?
                        'If tmMnfList(ilLoop).iBillMissMG <= 1 Or tmSdf.iMnfMissed = 0 Then     'nc missed, bill mg
                        ilDoAdjust = False
                        'For ilLoop = LBound(tmMnfList) To UBound(tmMnfList) Step 1
                            'If ((tmMnfList(ilLoop).iMnfCode = tmSdf.iMnfMissed) And (tmMnfList(ilLoop).iBillMissMG <= 1 Or tmMnfList(ilLoop).iBillMissMG = 3)) Or (tmSdf.iMnfMissed = 0) Then
                            '1=bill mg, nc missed  , 3 = bill both missed & mg
                                ilDoAdjust = True
                        '        Exit For
                            'End If
                        'Next ilLoop
                        If ilDoAdjust Then
                            ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)

                            If (InStr(slPrice, ".") <> 0) Then         'found spot cost
                                llActPrice = gStrDecToLong(slPrice, 2)
                                If ilMonthInx >= ilFirstProjInx Then            'only adjust if its in the future
                                    'lmProjectTNet(ilMonthInx) = lmProjectTNet(ilMonthInx) - tmClf.lAcquisitionCost
                                    lmProjectGross(ilMonthInx) = lmProjectGross(ilMonthInx) - llActPrice    'adjust the missed portion of the mg
                                End If
                            End If
                        End If
                    End If
                End If
            End If                          'btrv_err_none
        End If                              'schstatus = O,G
        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End If
Exit Sub
End Sub
'
'           mSplit - convert Receivables packed decimal to string for
'           math computations
'           <input>  llProcessPct - revenue share %
'                    llTransGross -  gross to split
'                    llTransNet - net to split
'                    llTransTNet - tnet to split
'                    ilReverseFlag - true to negate and subtract $
'           <oupput> llGrossDollars - slsp portion gross
'                    llNetDollars - slsp portion of net
'                    llTNetDollars - slsp portion of Tnet
Private Sub mSplitPast(llProcessPct As Long, llTransGross As Long, llTransNet As Long, llTransTNet As Long, llGrossDollar As Long, llNetDollar As Long, llTNetDollar As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slAcquisitionCost             llAcquisitionCost             ilLoopOnMonths            *
'*                                                                                        *
'******************************************************************************************

    Dim slPct As String
    Dim slAmount As String
    Dim slDollar As String

        slPct = gLongToStrDec(llProcessPct, 4)           'slsp split share

         slAmount = gLongToStrDec(llTransGross, 2)
         slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
         llGrossDollar = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)

         slAmount = gLongToStrDec(llTransNet, 2)
         slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
         llNetDollar = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)

         slAmount = gLongToStrDec(llTransTNet, 2)
         slDollar = gMulStr(slPct, slAmount)
         llTNetDollar = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)

        Exit Sub
End Sub
'
'               mGetParticipantSplits - build the table of participant splits
'           <input> llstdStartDate - effective start date to gather from PIF
'           <output> tmPifKey() - array of vehicles and indices into tmPifPct array
'                   tmPifPct() - array of participant percentages by vehicle
Private Sub mGetAllParticipantSplits(llStdStartDate As Long)
    gCreatePIFForRpts llStdStartDate, tmPifKey(), tmPifPct(), ExpRevenue
End Sub
'
'           mSplit - convert Receivables packed decimal to string for
'           math computations
'           <input>  llProcessPct - revenue share %
'                    llTransGross -  gross to split
'                    llTransNet - net to split
'                    llTransTNet - tnet to split
'                    ilReverseFlag - true to negate and subtract $
'           <oupput> llGrossDollars - slsp portion gross
'                    llNetDollars - slsp portion of net
'                    llTNetDollars - slsp portion of Tnet
Private Sub mSplitFuture(llProcessPct() As Long, llTransGross() As Long, llTransNet() As Long, llTransTNet() As Long, llGrossDollar() As Long, llNetDollar() As Long, llTNetDollar() As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slAcquisitionCost             llAcquisitionCost                                       *
'******************************************************************************************

    Dim slPct As String
    Dim slAmount As String
    Dim slDollar As String
    Dim ilLoopOnMonths As Integer


        For ilLoopOnMonths = 1 To 12
            If llProcessPct(ilLoopOnMonths) > 0 Then
                slPct = gLongToStrDec(llProcessPct(ilLoopOnMonths), 4)
                slAmount = gLongToStrDec(llTransGross(ilLoopOnMonths), 2)
                slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
                llGrossDollar(ilLoopOnMonths) = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)

                slAmount = gLongToStrDec(llTransNet(ilLoopOnMonths), 2)
                slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
                llNetDollar(ilLoopOnMonths) = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)

                slAmount = gLongToStrDec(llTransTNet(ilLoopOnMonths), 2)
                slDollar = gMulStr(slPct, slAmount)
                llTNetDollar(ilLoopOnMonths) = gLongToStrDec((gStrDecToLong(slDollar, 2) \ 100), 0)
            Else
                llGrossDollar(ilLoopOnMonths) = 0
                llNetDollar(ilLoopOnMonths) = 0
                llTNetDollar(ilLoopOnMonths) = 0
            End If
        Next ilLoopOnMonths
        Exit Sub
End Sub
'
'           Open all applicable files required for Exporting Revenue
'           If standard export, need actuals from receivables/history
'           If corporate export, only retrieve adjustments from receivables/history,
'           and obtain all revenue from contract files on an averaging week basis
'
Private Function mOpenExportFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer
ReDim tmNTRMNF(0 To 0) As MNF
ReDim tmVGMNF(0 To 0) As MNF

    ilError = False
    On Error GoTo mOpenExportFilesErr

    slTable = "SBF"
    hmSbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSbfRecLen = Len(tmSbf)

    slTable = "CHF"
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCHFRecLen = Len(tmChf)

    slTable = "SLF"
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSlfRecLen = Len(tmSlf)

    slTable = "VEF"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVefRecLen = Len(tmVef)

    slTable = "MNF"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMnfRecLen = Len(tmMnf)

    slTable = "SOF"
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSofRecLen = Len(tmSof)

    slTable = "CLF"
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imClfRecLen = Len(tmClf)

    slTable = "CFF"
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCffRecLen = Len(tmCff)

    slTable = "AGF"
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAgfRecLen = Len(tmAgf)

    slTable = "SDF"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSdfRecLen = Len(tmSdf)

    slTable = "SMF"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSmfRecLen = Len(tmSmf)

    slTable = "VSF"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVsfRecLen = Len(tmVsf)

    slTable = "ADF"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAdfRecLen = Len(tmAdf)

    slTable = "PRF"
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imPrfRecLen = Len(tmPrf)


    ilRet = gObtainMnfForType("I", "", tmNTRMNF())      'ntr types
    ilRet = gObtainMnfForType("H", "", tmVGMNF())      'vehicle groups

    mOpenExportFiles = ilError      'return any error flag
    Exit Function

mOpenExportFilesErr:
    ilError = True
    gBtrvErrorMsg ilRet, "mOpenExportFiles (OpenError) #" & str(ilRet) & ": " & slTable, ExpRevenue

    Resume Next
End Function
Private Sub cmcCancel_Click()

       If imExporting Then
           imTerminate = True
           Exit Sub
       End If
       mTerminate

End Sub
Private Sub cmcExport_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilYear                                                                                *
'******************************************************************************************
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim ilError As Integer
    ReDim llStartDates(0 To 13) As Long      '12 months of std/corp start dates, index zero ignored
    Dim llPacingDate As Long
    Dim ilLastBilledInx As Integer              'index in array of startdates for the year (into llstartdates)
    Dim llLastBilled As Long                    'last billed date (if std, need to get phf/rvf prior to last billed date)
    Dim ilMonth As Integer                      'starting month entered
    Dim slStr As String
    ReDim tmSofList(0 To 0) As SOFLIST
    ReDim tmTRvf(0 To 0) As RVFCNTSORT

       lacInfo(0).Visible = False
       lacInfo(1).Visible = False
       If imExporting Then
           Exit Sub
       End If
       On Error GoTo ExportError
       
        If Not Len(edcSelCFrom.Text) > 0 Then
            ''MsgBox "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
            gAutomationAlertAndLogHandler "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
            edcSelCFrom.SetFocus
            Exit Sub
        End If

        If Not Len(edcSelCFrom1.Text) > 0 Then
            ''MsgBox "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            gAutomationAlertAndLogHandler "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            edcSelCFrom1.SetFocus
            Exit Sub
        End If

        'validity check the month input
        slStr = edcSelCFrom.Text                    'month in text form (jan..dec)
        gGetMonthNoFromString slStr, ilMonth          'getmonth #
        If ilMonth = 0 Then                                 'input isn't text month name, try month #
            ilMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                Beep
                edcSelCFrom.SetFocus
                Exit Sub
            End If
        End If
        'validity check the year input
        slStr = edcSelCFrom1.Text
        igYear = gVerifyYear(slStr)
        If igYear = 0 Then
            Beep
            edcSelCFrom1.SetFocus                 'invalid year
            Exit Sub
        End If

        mGetAdjustedMonth ilMonth

        slToFile = sgExportPath & "REVENUE " & Trim$(edcSelCFrom.Text) & Trim$(edcSelCFrom1.Text) & ".CSV"
        If DoesFileExist(slToFile) Then
            Kill slToFile
        End If
        If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
            slToFile = sgExportPath & slToFile
            End If
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            'hmTo = FreeFile
            'Open slToFile For Append As hmTo
            ilRet = gFileOpen(slToFile, "Append", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         Else
            ilRet = 0
            'hmTo = FreeFile
            'Open slToFile For Output As hmTo
            ilRet = gFileOpen(slToFile, "Output", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         End If

        imExporting = True
        ilError = mOpenExportFiles()      'open all applicable files
        If ilError Then
            imTerminate = True
            Exit Sub
        End If

        On Error GoTo 0
        Screen.MousePointer = vbHourglass
        imExporting = True
        
        gLogMsg "", "ExportRevenue.txt", False
        gLogMsg "Revenue Export for " & Trim$(tgSpf.sGClient) & " : " & edcSelCFrom.Text & " " & edcSelCFrom1.Text, "ExportRevenue.txt", False
        gLogMsg "* StartMonth = " & edcSelCFrom.Text, "ExportRevenue.txt", False
        gLogMsg "* StartYear = " & edcSelCFrom1.Text, "ExportRevenue.txt", False
        If rbcYearType(0).Value = True Then
            gLogMsg "* Calendar = Corporate", "ExportRevenue.txt", False
        Else
            gLogMsg "* Calendar = Standard", "ExportRevenue.txt", False
        End If
        gLogMsg "* VehicleGroup = " & cbcVehGroup.Text, "ExportRevenue.txt", False
        gLogMsg "* Contract# = " & edcContract.Text, "ExportRevenue.txt", False
        
        lmSingleCntr = Val(edcContract)
        imVGGroupSelected = ExpRevenue!cbcVehGroup.ListIndex     '6-13-02
        imMajorVG = gFindVehGroupInx(imVGGroupSelected, tgVehicleSets1())
        imMinorVG = 0       'unused but common to subrooutine that gets the vehicle grous


        'create the header record containing client name, date/timegenerated, start month entered, last bill date
        mCreateHeader

        'On Error GoTo cmcExportErr
        gObtainSOF hmSof, tmSofList()   'get the sales offices and sales sources

        llPacingDate = 0                'pacing not applicable for this feature
        If rbcYearType(0).Value = True Then         'corporate
            gSetupBOBDates 1, llStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, igMonthOrQtr  'build array of corp start & end dates
            mGetAllParticipantSplits llStartDates(1)
            ilRet = mLoadRvf(llStartDates(1), llStartDates(13))   'get adjustments, merch/promo (for T-NET)
            If ilRet = False Then
                Exit Sub
            End If
            mPastRevenue llStartDates(), ilLastBilledInx
            mFutureRevenue llStartDates(), llLastBilled
         Else    'standard
            gSetupBOBDates 2, llStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, igMonthOrQtr  'build array of std start & end dates
            mGetAllParticipantSplits llStartDates(1)
            ilRet = mLoadRvf(llStartDates(1), llStartDates(13))  'get adjustments, merch & Promo
            If ilRet = False Then
                Exit Sub
            End If
            gLogMsg "Getting Past Revenue..", "ExportRevenue.txt", False
            mPastRevenue llStartDates(), ilLastBilledInx
            gLogMsg "Getting Future Revenue..", "ExportRevenue.txt", False
            mFutureRevenue llStartDates(), llLastBilled
        End If

        If ilRet = False Then        'error will be an error code
            lacInfo(0).Caption = "Export Failed"
            gLogMsg "Export failed: #" & Trim$(str$(ilRet)), "ExportRevenue.txt", False
        Else
            lacInfo(0).Caption = "Export Successfully Completed"
            gLogMsg "Export Successfully Completed, Export Files: " & slToFile, "ExportRevenue.txt", False
        End If
        lacInfo(1).Caption = "Export Files: " & slToFile

        lacInfo(0).Visible = True
        lacInfo(1).Visible = True
        Close hmTo

        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
        cmcExport.Enabled = False
        Screen.MousePointer = vbDefault
        imExporting = False

        Exit Sub
'cmcExportErr:
'        ilRet = Err.Number
'        Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub Form_Activate()

       If Not imFirstActivate Then
           DoEvents    'Process events so pending keys are not sent to this
           Me.KeyPreview = True
           Exit Sub
       End If
       imFirstActivate = False
       DoEvents    'Process events so pending keys are not sent to this
       Me.KeyPreview = True
       Me.Refresh
       edcSelCFrom.SetFocus

End Sub
Private Sub Form_Deactivate()

       Me.KeyPreview = False

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

       If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
           gFunctionKeyBranch KeyCode
       End If

End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

       sgDoneMsg = CmdStr
       igChildDone = True
       Cancel = 0

End Sub
Private Sub Form_Load()

    mInit
    If imTerminate Then
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmAdf)
    btrDestroy hmRvf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmPrf
    btrDestroy hmSbf
    btrDestroy hmVsf
    btrDestroy hmAdf

    Erase tmTRvf
    Erase tmPifPct
    Erase tmPifKey
    Erase tmCntAllYear
    Erase tmOneVehAllYear
    Erase tmSofList
    Erase tmAdjust
    Erase tmNTRMNF
    Erase tmVGMNF
    Erase lmSlfSplit
    Erase imSlfCode
    Erase tmSBFAdjust

    Set ExpRevenue = Nothing   'Remove data segment

End Sub

Private Sub rbcYearType_Click(Index As Integer)
    Dim Value As Integer
    Value = rbcYearType(Index).Value

End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStdDate                     ilError                                                 *
'******************************************************************************************


    Dim ilRet As Integer

        imTerminate = False
        imFirstActivate = True
        Screen.MousePointer = vbHourglass
        imExporting = False
        imFirstFocus = True
        imBypassFocus = False
        lmTotalNoBytes = 0
        lmProcessedNoBytes = 0

        'default to corporate calendar if using it; otherwise disable it
        If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
            rbcYearType(0).Enabled = False
            rbcYearType(0).Value = False
            rbcYearType(1).Value = True
        Else
            ilRet = gObtainCorpCal()
            If rbcYearType(0).Value Then
                rbcYearType_Click 0
            Else
                rbcYearType(0).Value = True
            End If
        End If

        gPopVehicleGroups ExpRevenue!cbcVehGroup, tgVehicleSets1(), True

        gCenterStdAlone ExpRevenue
        Screen.MousePointer = vbDefault
        gAutomationAlertAndLogHandler ""
        gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
        
        Exit Sub

End Sub

Private Sub mTerminate()


        Screen.MousePointer = vbDefault
        igManUnload = YES
        Unload ExpRevenue
        igManUnload = NO

End Sub
'
'               mLoadRVF - retrieve Invoice and/or Adjustments for dates equal/prior
'               to last billing date
'               <input> llStartDate - earliest phf/rvf transactions to retrieve
'                       llEndDate - latest phf/rvf transactions to retrieve
'
Private Function mLoadRvf(llStartDate As Long, llEndDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slStr1                        ilYear                    *
'*  llRvfLoop                                                                             *
'******************************************************************************************

    Dim ilRet As Integer

    Dim tlTranType As TRANTYPES
    Dim ilWhichDate As Integer      '0=use tran date, 1 = use date entered
    Dim slStartDate As String
    Dim slEndDate As String

    mLoadRvf = True
    tlTranType.iNTR = True
    tlTranType.iAirTime = True
    tlTranType.iInv = True              'invoices
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iMerch = True            'get merch & promo to calc T-Net amts
    tlTranType.iPromo = True
    tlTranType.iAdj = True              'adjustments

    'If rbcYearType(0).Value = True Then     'corporate, only retrieve the adjustments
    '    tlTranType.iInv = False
    'End If

    ilWhichDate = 0                     'default to use tran date vs date entered
    On Error GoTo mLoadRvfErr
    slStartDate = Format(llStartDate, "m/d/yy")
    slEndDate = Format(llEndDate, "m/d/yy")
    ilRet = gObtainPhfRvfforSort(ExpRevenue, slStartDate, slEndDate, tlTranType, tmTRvf(), ilWhichDate)
    If ilRet = False Then
        mLoadRvf = False
        Exit Function
    Else
        If UBound(tmTRvf) > 2 Then
            ArraySortTyp fnAV(tmTRvf(), 0), UBound(tmTRvf), 0, LenB(tmTRvf(1)), 0, Len(tmTRvf(1).sKey), 0
        End If
    End If

    Exit Function
mLoadRvfErr:
    gDbg_HandleError "Messages: mLoadRvf"
    Resume Next
End Function


Private Sub mPastRevenue(llStartDates() As Long, ilLastBilledInx As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llAcqAmt                      ilMatchSOFCode                ilOKtoSeeVeh              *
'*  ilFoundOption                 ilSaveLastBilledInx           ilSlfInx                  *
'*  llPastGross                   llPastNet                     llPastTNet                *
'*                                                                                        *
'******************************************************************************************

Dim llRvfLoop As Long
Dim ilLoop As Integer
Dim llDate As Long
Dim slStr As String
Dim slCode As String
Dim llTransNet As Long
Dim llTransGross As Long
Dim llTransTNet As Long
Dim ilUseSlsComm As Integer
Dim llSlfSplit() As Long           'slsp slsp share %
Dim ilSlfCode() As Integer
Dim llSlfSplitRev() As Long
Dim ilMatchCntr As Integer
Dim ilMatchSSCode As Integer
Dim ilFoundMonth As Integer
Dim ilRet As Integer
Dim ilSlfRecd As Integer
Dim ilMonthNo As Integer
ReDim ilProdPct(0 To 1) As Integer  'index zero is ignored
ReDim ilMnfGroup(0 To 1) As Integer 'index zero is ignored
ReDim ilMnfSSCode(0 To 1) As Integer    'index zero is ignored
Dim ilTemp As Integer
Dim ilMnfSubCo As Integer
Dim llProcessPct As Long
Dim ilUse100pct As Integer
Dim llGrossDollar As Long
Dim llNetDollar As Long
Dim llTNetDollar As Long
Dim ilHowManyDefined As Integer
Dim ilLoopOnParts As Integer
Dim ilReverseSign As Integer
'participants share transaction
Dim llPartGross As Long
Dim llPartNet As Long
Dim llPartTNet As Long
'running total less all the slsp splits to see whats left so that the
'last participant/slsp can get remaining pennies
Dim llSlspRemGross As Long
Dim llSlspRemNet As Long
Dim llSlspRemTNet As Long
'running total less all the participant splits to see whats left so that the
'last participant/slsp can get remaining pennies
Dim llPartRemGross As Long
Dim llPartRemNet As Long
Dim llPartRemTNet As Long
Dim ilTempSlf() As Integer
Dim llTempSlfSplit() As Long
Dim ilIsItHardCost As Integer



        For llRvfLoop = LBound(tmTRvf) To UBound(tmTRvf) - 1 Step 1
            tmRvf = tmTRvf(llRvfLoop).tRVF

            'llTransNet, llTransGross, llTransTNet are the starting point to calc each of the slsp splits calcs from
            gPDNToLong tmRvf.sNet, llTransNet
            gPDNToLong tmRvf.sGross, llTransGross
            llTransTNet = tmRvf.lAcquisitionCost

            If tgSpf.sSEnterAgeDate = "E" Then      'use entered date or ageing date
                gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
            Else
                slCode = Trim$(str$(tmRvf.iAgePeriod) & "/15/" & Trim$(str$(tmRvf.iAgingYear)))
                slStr = gObtainEndStd(slCode)
            End If
            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
            'valid record must be an "Invoice", History Invoices, or Adjustment types, non-zero amount, and transaction date within the start date of the
            'cal year and end date of the current cal month requested
            'Merchandising and Promotions contracts for net-net options to be subtracted from the net amounts
            If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI" Or Left$(tmRvf.sTranType, 1) = "A") And (Trim$(tmRvf.sType) = "" Or tmRvf.sType = "A") And (llTransNet <> 0) And (llDate >= llStartDates(1))) Then             'looking for Invoice types only

                If tmChf.lCntrNo <> tmRvf.lCntrNo Then
                    'get contract from history or rec file
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

                    '9-19-06 when there is no sch lines and need to process merchandising for t-net, the contract may not be scheduled (schstatus = "N");  need to process those contracts
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" And tmChf.sSchStatus <> "N")
                         ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop

                    gFakeChf tmRvf, tmChf       'create a bare-bone header if the contract doesnt exist

                    ilMatchCntr = True
                    'exclude psa/promo, contracts not fully or manually sched or proposals
                    If tmChf.lCntrNo <> tmRvf.lCntrNo Or tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" And tmChf.sSchStatus <> "N" Or (lmSingleCntr > 0 And lmSingleCntr <> tmRvf.lCntrNo) Or tmChf.sType = "M" Or tmChf.sType = "S" Then
                        ilMatchCntr = False
                    End If

                End If
                If ilMatchCntr Then

                    'determine the month that this transaction falls within
                    ilFoundMonth = False
                    For ilMonthNo = 1 To 12 Step 1         'loop thru months to find the match
                        If llDate >= llStartDates(ilMonthNo) And llDate < llStartDates(ilMonthNo + 1) Then
                            'if this transaction if for a month in the future, ignore the AN (adjustments ) and
                            'only process the Merchandising and Promotions for net-net/triple net options.
                            If ilMonthNo <= ilLastBilledInx Then        'ok, prior to last billed month
                                If rbcYearType(1).Value = True Then               'for std option, need the IN and AN
                                    If (tmRvf.sCashTrade = "C" Or tmRvf.sCashTrade = "T" Or tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P") And (tmRvf.sTranType = "AN" Or tmRvf.sTranType = "IN" Or tmRvf.sTranType = "HI") Then     '9-26-06 include HI for the past
                                        ilFoundMonth = True
                                    End If
                                Else                        'for corporate, only AN since the revenue is driven by lines
                                                            '6-19-08 adjusting acquisition , if merch/promo need to process
                                    If (tmRvf.sCashTrade = "C" Or tmRvf.sCashTrade = "T") And (tmRvf.sTranType = "AN") Or ((tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P") And (tmRvf.sTranType = "IN" Or tmRvf.sTranType = "AN")) Then
                                        ilFoundMonth = True
                                    End If
                                End If
                                Exit For
                            Else                                        'after last billed month, ok to use if Merch/promo
                                If (tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P") And (tmRvf.sTranType = "AN" Or tmRvf.sTranType = "IN" Or tmRvf.sTranType = "HI") Then
                                    ilFoundMonth = True
                                End If
                                Exit For
                            End If

                        End If
                    Next ilMonthNo
                    If ilFoundMonth Then

                        ilReverseSign = mReverseSign(llTransGross, llTransNet, llTransTNet)

                        ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
                        ReDim ilSlfCode(0 To 9) As Integer             '4-20-00
                        ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

                        'set the slsp used for this contract.  If subcompany used, more than 1 subcompany can be defined on a contract.
                        'Be sure to use the correct vehicles slsp for the proper subcompany
                        ilUseSlsComm = False
                        ilMnfSubCo = gGetSubCmpy(tmChf, ilSlfCode(), llSlfSplit(), tmRvf.iAirVefCode, ilUseSlsComm, llSlfSplitRev())                                         '4-6-00
                        'ilslfCode() has the array of slsp matching the vehicles subcompany (if applicable); otherwise,
                        'its the same list of slsp defined in the header
                        'some of the Slsp do not have any percentages defined, so they must be ignored.
                        'when splitting $, extra pennies may exist so the last slsp with a % defined
                        'must get the extra pennies.
                        ReDim ilTempSlf(0 To 0) As Integer
                        ReDim llTempSlfSplit(0 To 0) As Long
                        ilHowManyDefined = 0
                        'determine the actual number of slsp to process
                        'check for both slsp code and split % because they may not all be defined
                        'with percentages
                        'i.e. slsp A:  50%
                        '      slsp b:  0
                        '      slsp c:  50%

                        For ilTemp = 0 To 9
                            If ilSlfCode(ilTemp) > 0 And llSlfSplit(ilTemp) > 0 Then        'both slsp & % split defined
                                ilTempSlf(ilHowManyDefined) = ilSlfCode(ilTemp)
                                llTempSlfSplit(ilHowManyDefined) = llSlfSplit(ilTemp)
                                ilHowManyDefined = ilHowManyDefined + 1
                                ReDim Preserve ilTempSlf(0 To ilHowManyDefined) As Integer
                                ReDim Preserve llTempSlfSplit(0 To ilHowManyDefined) As Long
                            End If
                        Next ilTemp
                        If ilHowManyDefined = 0 Then        'nothing found with slsp %, force to first slsp at 100%
                            ilTempSlf(0) = ilSlfCode(0)
                            llTempSlfSplit(0) = 10000
                            ReDim Preserve ilTempSlf(0 To 1) As Integer
                            ReDim Preserve llTempSlfSplit(0 To 1) As Long
                        End If

                        'llSlsp fields are the running totals minus each slsp split to see what left after all slsp have
                        'been processed.  Need to give last slsp the extra pennies
                        llSlspRemGross = llTransGross
                        llSlspRemNet = llTransNet
                        llSlspRemTNet = llTransTNet

                        'For slsp, create as many as 10 records per trans (up to 10 split slsp) per trans.
                        For ilLoop = 0 To UBound(ilTempSlf) - 1 Step 1           'loop based on report option

                            ilMatchSSCode = 0
                            llProcessPct = 0
                            For ilSlfRecd = LBound(tgMSlf) To UBound(tgMSlf)
                                If ilTempSlf(ilLoop) = tgMSlf(ilSlfRecd).iCode Then
                                    tmSlf = tgMSlf(ilSlfRecd)
                                    ilMatchSSCode = mFindMatchSS(tmSofList(), tmSlf.iSofCode)       'get the matching sales souce
                                    'may need the subcompany to determine which slsp to pick up for the transactions vehicle
                                    Exit For
                                End If
                            Next ilSlfRecd
                            If ilMatchSSCode = 0 Then           'no more slsp to process as no matching sales source
                                Exit For
                            End If
                            llProcessPct = llTempSlfSplit(ilLoop)  'llSlfSplit(ilLoop)

                            If llProcessPct > 0 Then            'split the slsp
                                mSplitPast llProcessPct, llTransGross, llTransNet, llTransTNet, llGrossDollar, llNetDollar, llTNetDollar

                                'see whats left after $ have been distributed for last slsp
                                llSlspRemGross = llSlspRemGross - llGrossDollar
                                llSlspRemNet = llSlspRemNet - llNetDollar
                                llSlspRemTNet = llSlspRemTNet - llTNetDollar

                                'adjust for last slsp, remaining pennies
                                If ilLoop = UBound(ilTempSlf) - 1 Then
                                    llGrossDollar = llGrossDollar + llSlspRemGross
                                    llNetDollar = llNetDollar + llSlspRemNet
                                    llTNetDollar = llTNetDollar + llSlspRemTNet
                                End If

                                'lmPart fields are the running totals minus each participant split to see what left after all slsp have
                                'been processed.  Need to give last participant the extra pennies
                                llPartRemGross = llGrossDollar
                                llPartRemNet = llNetDollar
                                llPartRemTNet = llTNetDollar

                                ilUse100pct = False             ' dont use 100% for participant share, search for the participants %
                                'split the participants
                                gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iTranDate(), tmPifKey(), tmPifPct(), ilUse100pct
                                For ilLoopOnParts = LBONE To UBound(ilMnfGroup)
                                    For ilTemp = 1 To 12 Step 1               'init the years $ buckets for the next participant
                                        tmInputInfo.lCashGross(ilTemp) = 0
                                        tmInputInfo.lCashNet(ilTemp) = 0
                                        tmInputInfo.lCashTNet(ilTemp) = 0
                                        tmInputInfo.lTradeGross(ilTemp) = 0
                                        tmInputInfo.lTradeNet(ilTemp) = 0
                                        tmInputInfo.lTradeTNet(ilTemp) = 0
                                    Next ilTemp
                                    llProcessPct = CLng(ilProdPct(ilLoopOnParts)) * 100
                                    If llProcessPct > 0 Then
                                        mSplitPast llProcessPct, llGrossDollar, llNetDollar, llTNetDollar, llPartGross, llPartNet, llPartTNet

                                        'see whats left after $ have been distributed for last participant
                                        llPartRemGross = llPartRemGross - llPartGross
                                        llPartRemNet = llPartRemNet - llPartNet
                                        llPartRemTNet = llPartRemTNet - llPartTNet

                                        'adjust for last participant, remaining pennies
                                        If ilLoopOnParts = UBound(ilMnfGroup) Then
                                            llPartGross = llPartGross + llPartRemGross
                                            llPartNet = llPartNet + llPartRemNet
                                            llPartTNet = llPartTNet + llPartRemTNet
                                        End If

                                        tmInputInfo.iSlfCode = tmSlf.iCode
                                        If ilReverseSign Then
                                            If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then    'merch or promo, consider it for netnet only
                                                'merchandising record should never have an acquisition value with it
                                                 tmInputInfo.lCashTNet(ilMonthNo) = tmInputInfo.lCashNet(ilMonthNo) - llPartNet      'accumulate the $ to sub for netnet (merch $)
                                            ElseIf tmRvf.sCashTrade = "C" Then
                                                'acquisition values can be with regular cash/trade IN/AN transactions
                                                tmInputInfo.lCashGross(ilMonthNo) = tmInputInfo.lCashGross(ilMonthNo) - llPartGross
                                                tmInputInfo.lCashNet(ilMonthNo) = tmInputInfo.lCashNet(ilMonthNo) - llPartNet
                                                tmInputInfo.lCashTNet(ilMonthNo) = tmInputInfo.lCashTNet(ilMonthNo) - llPartTNet
                                            Else
                                                tmInputInfo.lTradeGross(ilMonthNo) = tmInputInfo.lTradeGross(ilMonthNo) - llPartGross
                                                tmInputInfo.lTradeNet(ilMonthNo) = tmInputInfo.lTradeNet(ilMonthNo) - llPartNet
                                                tmInputInfo.lTradeTNet(ilMonthNo) = tmInputInfo.lTradeTNet(ilMonthNo) - llPartTNet
                                            End If
                                        Else
                                            If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then    'merch or promo, consider it for netnet only
                                                'merchandising record should never have an acquisition value with it
                                                 tmInputInfo.lCashTNet(ilMonthNo) = tmInputInfo.lCashNet(ilMonthNo) + llPartNet      'accumulate the $ to sub for netnet (merch $)
                                            ElseIf tmRvf.sCashTrade = "C" Then
                                                'acquisition values can be with regular cash/trade IN/AN transactions
                                                tmInputInfo.lCashGross(ilMonthNo) = tmInputInfo.lCashGross(ilMonthNo) + llPartGross
                                                tmInputInfo.lCashNet(ilMonthNo) = tmInputInfo.lCashNet(ilMonthNo) + llPartNet
                                                tmInputInfo.lCashTNet(ilMonthNo) = tmInputInfo.lCashTNet(ilMonthNo) + llPartTNet
                                            Else
                                                tmInputInfo.lTradeGross(ilMonthNo) = tmInputInfo.lTradeGross(ilMonthNo) + llPartGross
                                                tmInputInfo.lTradeNet(ilMonthNo) = tmInputInfo.lTradeNet(ilMonthNo) + llPartNet
                                                tmInputInfo.lTradeTNet(ilMonthNo) = tmInputInfo.lTradeTNet(ilMonthNo) + llPartTNet
                                            End If
                                        End If
                                        tmInputInfo.iMnfCode = ilMnfGroup(ilLoopOnParts)            'participant code
                                        tmInputInfo.iVefCode = tmRvf.iAirVefCode
                                        gGetVehGrpSets tmInputInfo.iVefCode, imMinorVG, imMajorVG, ilTemp, tmInputInfo.iVefGroup
                                        If imMajorVG = 1 Then               'participant vehicle group selected
                                            tmInputInfo.iVefGroup = tmInputInfo.iMnfCode
                                        End If
                                        tmInputInfo.sCashTrade = tmRvf.sCashTrade
                                        If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then        'assume cash if the transaction is the merchadising portion
                                            tmInputInfo.sCashTrade = "C"
                                        End If
                                        tmInputInfo.sAirTimeNTR = "A"           'assume air time
                                        If tmRvf.lSbfCode > 0 Then              'if sbf pointer, its an NTR
                                            'determine hard cost
                                            ilIsItHardCost = gIsItHardCost(tmRvf.iMnfItem, tmNTRMNF())
                                            If ilIsItHardCost Then
                                                tmInputInfo.sAirTimeNTR = "H"
                                            Else
                                                tmInputInfo.sAirTimeNTR = "N"
                                            End If
                                        End If
                                        mCreateExportRecord
                                    End If
                                Next ilLoopOnParts
                            End If                              'llProcessPct > 0
                        Next ilLoop                             'loop for 10 slsp possible splits, or 3 possible owners, otherwise loop once
                    End If                                      'if foundmonth
                End If
            End If                                          'contr # doesnt match or not a fully sched contr
        Next llRvfLoop          '03-13-01

        Erase ilTempSlf, llTempSlfSplit
        Erase ilSlfCode, llSlfSplit, llSlfSplitRev
        Exit Sub
End Sub
'
'               mFindMatchSS - get the Sales Source from the Sales Office
'               <input> Salesperson sales office
'                   array of selling offices and sales sources
'                   sof code to match
'               <return> - sales source mnf code
'
'
Private Function mFindMatchSS(tlSofList() As SOFLIST, ilSofCode As Integer) As Integer
Dim ilLoop As Integer
Dim ilMatchSSCode As Integer

        'associated sales source
        ilMatchSSCode = 0
        For ilLoop = LBound(tlSofList) To UBound(tlSofList)
            If tlSofList(ilLoop).iSofCode = ilSofCode Then
                ilMatchSSCode = tlSofList(ilLoop).iMnfSSCode          'Sales source
                Exit For
            End If
        Next ilLoop
        mFindMatchSS = ilMatchSSCode
End Function
'           mReverseSign - always work with postitive amounts
'           <input> llTransGross - receivables gross amt
'                   llTransNet - receivables net amt
'                   llTransTNet - receivables acquisition amt
'           <output> llTransGross - gross amt of transaction (possibly sign reversed)
'                    llTransNet - Net amount of tran (possibly sign reversed)
'                    llTransAct - acquisition amt of trans (possibly sign reversed)
'           <return> true if sign reversed
Private Function mReverseSign(llTransGross As Long, llTransNet As Long, llTransTNet As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slAmount                                                                              *
'******************************************************************************************

Dim ilReverseFlag As Integer


'          If (tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P") Then
'              'Reverse the sign on the gross & net fields, then store it back into the RVf record
'              'so it properly subtracts this amount from commissions
'              'If the transaction is negative, need to give commision back, so add it in.
'              'if the transaction is positive, its not counted as commissionable
'              ilReverseFlag = True
'          End If
          If llTransNet < 0 Then      'always work with positive amounts (if already positive, leave it)
              'retain reversal flag if its an adjustment (AN)
'              If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then
'                  ilReverseFlag = False   'make sure that the merchandise/promotion amt is added back in
'              Else
                  ilReverseFlag = True
'              End If
              llTransGross = -llTransGross
              llTransNet = -llTransNet
              llTransTNet = -llTransTNet

          End If
          mReverseSign = ilReverseFlag
End Function
'
'               mCreateExportRecord - obtain all the associated files
'               and build into an export comma delimited record
'               <input>  llGross - array of 12 months gross amts
'                        llNet - array of 12 months net amts
'                        llTNet - array of 12 months triple net amts
'                        ilVefCode - vehicle code
'                        ilSlfCode - slsp code (can have up to 10 splits on a contract)
'                        ilMnfCode - participant split
'               <output>  none
'
Public Sub mCreateExportRecord()
Dim slStr As String
Dim ilLoop As Integer
Dim ilError As Integer
Dim ilIsItPolitical As Integer
Dim ilRet As Integer
Dim slRecord As String
Dim llNetNet As Long
Dim ilLoopOnCT As Integer
Dim llTempNet As Long
Dim llTempTNet As Long

        tmExportInfo.sContract = Trim$(str$(tmChf.lCntrNo))         'Contract #

        'contract type
        'C=Standard; V=Reservation; T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo

        If tmChf.sType = "V" Then
            tmExportInfo.sType = "Reservation"
        ElseIf tmChf.sType = "T" Then
            tmExportInfo.sType = "Remnant"
        ElseIf tmChf.sType = "R" Then
            tmExportInfo.sType = "Direct Response"
        ElseIf tmChf.sType = "Q" Then
            tmExportInfo.sType = "Per Inquiry"
        ElseIf tmChf.sType = "S" Then
            Exit Sub
        ElseIf tmChf.sType = "P" Then
            Exit Sub
        Else
            tmExportInfo.sType = "Standard"
        End If

        'Advertiser Name
        ilLoop = gBinarySearchAdf(tmChf.iAdfCode)
        If ilLoop <> -1 Then
            tmExportInfo.sAdvtName = tgCommAdf(ilLoop).sName
            ilIsItPolitical = gIsItPolitical(tgCommAdf(ilLoop).iCode)           'its a political, include this contract?
            If ilIsItPolitical Then                 'its a political
                tmExportInfo.sPolitical = "Y"
            Else                                    'not a plitical
                tmExportInfo.sPolitical = "N"
            End If
        Else
            tmExportInfo.sAdvtName = "Unknown Advertiser"
            tmExportInfo.sPolitical = "N"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": " & "Invalid Advertiser ID : " & Trim$(str(tmChf.iAdfCode)), "ExportRevenue.txt", False
        End If

        'Agency Name, test for Direct advertiser
        If tmChf.iAgfCode > 0 Then          'test for direct advertiser
            ilLoop = gBinarySearchAgf(tmChf.iAgfCode)
            If ilLoop <> -1 Then
                tmExportInfo.sAgency = tgCommAgf(ilLoop).sName
            Else
                ilError = True
                gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": " & "Invalid Agency ID : " & Trim$(str(tmChf.iAgfCode)), "ExportRevenue.txt", False
            End If
        Else
            tmExportInfo.sAgency = "Direct"
        End If

        tmExportInfo.sProduct = tmChf.sProduct          'contract product

        'Vehicle Name
        If tmInputInfo.iVefCode > 0 Then
            ilLoop = gBinarySearchVef(tmInputInfo.iVefCode)
            If ilLoop <> -1 Then
                tmExportInfo.sVehicle = Trim$(tgMVef(ilLoop).sName)
            Else
                ilError = True
                gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": " & Trim$(str(tmInputInfo.iVefCode)), "ExportRevenue.txt", False
             End If
        Else
            tmExportInfo.sVehicle = "Unknown Vehicle"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": Invlid vehicle- " & Trim$(str(tmInputInfo.iVefCode)), "ExportRevenue.txt", False
        End If

        'Salesperson
        If tmInputInfo.iSlfCode > 0 Then
            ilLoop = gBinarySearchSlf(tmInputInfo.iSlfCode)
            If ilLoop <> -1 Then
                tmExportInfo.sSlsp = Trim$(tmSlf.sFirstName) + " " + Trim$(tmSlf.sLastName)
            Else
                ilError = True
                gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": " & Trim$(str(tmInputInfo.iSlfCode)), "ExportRevenue.txt", False
             End If
        Else
            tmExportInfo.sSlsp = "Unknown Salesperson"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": Invalid Salesperson -" & Trim$(str(tmInputInfo.iSlfCode)), "ExportRevenue.txt", False
        End If

        'Sales office
        tmSofSrchKey.iCode = tgMSlf(ilLoop).iSofCode
        ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmExportInfo.sOffice = tmSof.sName
        Else
            tmExportInfo.sOffice = "Unknown Office"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": Invalid Sales Office- " & Trim$(str(tgMSlf(ilLoop).iSofCode)), "ExportRevenue.txt", False
        End If

        'Sales source
        tmMnfSrchKey0.iCode = tmSof.iMnfSSCode
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmExportInfo.sSalesSource = tmMnf.sName
        Else
            tmExportInfo.sSalesSource = "Unknown Sales Source"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & ": Invalid Sales Source- " & Trim$(str(tmSof.iMnfSSCode)), "ExportRevenue.txt", False
        End If

        'Participant
        tmMnfSrchKey0.iCode = tmInputInfo.iMnfCode
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmExportInfo.sParticipant = tmMnf.sName
        Else
            tmExportInfo.sSalesSource = "Unknown Participant"
            ilError = True
            gLogMsg "Contract # " & Trim$(tmChf.lCntrNo) & " for " & Trim$(tmExportInfo.sVehicle) & ": Invalid Participant- " & Trim$(str(tmInputInfo.iMnfCode)), "ExportRevenue.txt", False
        End If

        'vehicle group
        If tmInputInfo.iVefGroup > 0 Then
            For ilLoop = LBound(tmVGMNF) To UBound(tmVGMNF) - 1
                If tmVGMNF(ilLoop).iCode = tmInputInfo.iVefGroup Then
                    tmExportInfo.sVehicleGroup = tmVGMNF(ilLoop).sName
                    Exit For
                End If
            Next ilLoop
        Else
            tmExportInfo.sVehicleGroup = ""
        End If

        For ilLoopOnCT = 1 To 2             'loop for cash and trade
            slRecord = Trim$(tmExportInfo.sContract) & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sAdvtName) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sProduct) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sAgency) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sVehicle) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sVehicleGroup) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sSlsp) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sOffice) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sSalesSource) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sParticipant) & """" & ","
            If ilLoopOnCT = 1 Then
                slRecord = slRecord & """" & "C" & """" & ","
            Else
                slRecord = slRecord & """" & "T" & """" & ","
            End If
            If tmInputInfo.sAirTimeNTR = "A" Then
                tmExportInfo.sAirTimeNTR = "Airtime"
            ElseIf tmInputInfo.sAirTimeNTR = "H" Then
                tmExportInfo.sAirTimeNTR = "Hardcost"
            Else
                tmExportInfo.sAirTimeNTR = "NTR"
            End If
            slRecord = slRecord & """" & Trim$(tmExportInfo.sAirTimeNTR) & """" & ","
            slRecord = slRecord & """" & Trim$(tmExportInfo.sPolitical) & """" & ","

            llTempNet = 0
            llTempTNet = 0
            If ilLoopOnCT = 1 Then
                'loop thru the cash net and tnet dollars to see if theres anything to create; ignore if all zeros
                For ilLoop = 1 To 12
                    llTempNet = llTempNet + tmInputInfo.lCashNet(ilLoop)
                    llTempTNet = llTempTNet + tmInputInfo.lCashTNet(ilLoop)
                Next ilLoop
                If llTempNet + llTempTNet <> 0 Then      'dont create a record if there is no $
                    For ilLoop = 1 To 12            'put out 12 months of gross $
                        slStr = gLongToStrDec(tmInputInfo.lCashGross(ilLoop), 2)
                        tmExportInfo.sGross(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    For ilLoop = 1 To 12            'put out 12 months of net $
                        slStr = gLongToStrDec(tmInputInfo.lCashNet(ilLoop), 2)
                        tmExportInfo.sNet(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    For ilLoop = 1 To 12            'put out 12 months of netnet$
                        llNetNet = tmInputInfo.lCashNet(ilLoop) - tmInputInfo.lCashTNet(ilLoop)
                        slStr = gLongToStrDec(llNetNet, 2)
                        tmExportInfo.sNetNet(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    slRecord = slRecord & """" & Trim$(tmExportInfo.sType) & """"

                    Print #hmTo, slRecord
                    For ilLoop = 1 To 12 Step 1               'init the years $ buckets for the next participant
                        tmInputInfo.lCashGross(ilLoop) = 0
                        tmInputInfo.lCashNet(ilLoop) = 0
                        tmInputInfo.lCashTNet(ilLoop) = 0
                    Next ilLoop
                End If
            Else
                'loop thru the cash net and tnet dollars to see if theres anything to create; ignore if all zeros
                For ilLoop = 1 To 12
                    llTempNet = llTempNet + tmInputInfo.lTradeNet(ilLoop)
                    llTempTNet = llTempTNet + tmInputInfo.lTradeTNet(ilLoop)
                Next ilLoop
                If llTempNet + llTempTNet <> 0 Then      'dont create a record if there is no $
                    For ilLoop = 1 To 12            'put out 12 months of gross $
                        slStr = gLongToStrDec(tmInputInfo.lTradeGross(ilLoop), 2)
                        tmExportInfo.sGross(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    For ilLoop = 1 To 12            'put out 12 months of net $
                        slStr = gLongToStrDec(tmInputInfo.lTradeNet(ilLoop), 2)
                        tmExportInfo.sNet(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    For ilLoop = 1 To 12            'put out 12 months of netnet$
                        llNetNet = tmInputInfo.lTradeNet(ilLoop) - tmInputInfo.lTradeTNet(ilLoop)
                        slStr = gLongToStrDec(llNetNet, 2)
                        tmExportInfo.sNetNet(ilLoop) = gRoundStr(slStr, ".01", 2)
                        'slRecord = slRecord & """" & Trim$(slStr) & """" & ","
                        slRecord = slRecord & Trim$(slStr) & ","
                    Next ilLoop

                    slRecord = slRecord & """" & Trim$(tmExportInfo.sType) & """"

                    Print #hmTo, slRecord
                    For ilLoop = 1 To 12 Step 1               'init the years $ buckets for the next participant
                        tmInputInfo.lTradeGross(ilLoop) = 0
                        tmInputInfo.lTradeNet(ilLoop) = 0
                        tmInputInfo.lTradeTNet(ilLoop) = 0
                    Next ilLoop
                End If
            End If

        Next ilLoopOnCT

End Sub
Public Sub mCreateHeader()
Dim slRecord As String
Dim slDateGenned As String
Dim slTimeGenned As String
Dim slLastInvDate As String
Dim slYearType As String
Dim llDate As Long

        On Error GoTo mCreateHeaderErr:

        '3-4-09 create column headings
        slRecord = """" & "Contract #" & """" & ","
        slRecord = slRecord & """" & "Advertiser" & """" & ","
        slRecord = slRecord & """" & "Product" & """" & ","
        slRecord = slRecord & """" & "Agency #" & """" & ","
        slRecord = slRecord & """" & "Vehicle" & """" & ","
        slRecord = slRecord & """" & "Vehicle Group" & """" & ","
        slRecord = slRecord & """" & "Salesperson" & """" & ","
        slRecord = slRecord & """" & "Sales Source" & """" & ","
        slRecord = slRecord & """" & "Sales Office" & """" & ","
        slRecord = slRecord & """" & "Participant" & """" & ","
        slRecord = slRecord & """" & "Cash/Trade" & """" & ","
        slRecord = slRecord & """" & "AirTime/NTR/HC" & """" & ","
        slRecord = slRecord & """" & "Political" & """" & ","
        slRecord = slRecord & """" & "Gross Per 1" & """" & ","
        slRecord = slRecord & """" & "Gross Per 2" & """" & ","
        slRecord = slRecord & """" & "Gross Per 3" & """" & ","
        slRecord = slRecord & """" & "Gross Per 4" & """" & ","
        slRecord = slRecord & """" & "Gross Per 5" & """" & ","
        slRecord = slRecord & """" & "Gross Per 6" & """" & ","
        slRecord = slRecord & """" & "Gross Per 7" & """" & ","
        slRecord = slRecord & """" & "Gross Per 8" & """" & ","
        slRecord = slRecord & """" & "Gross Per 9" & """" & ","
        slRecord = slRecord & """" & "Gross Per 10" & """" & ","
        slRecord = slRecord & """" & "Gross Per 11" & """" & ","
        slRecord = slRecord & """" & "Gross Per 12" & """" & ","
        slRecord = slRecord & """" & "Net Per 1" & """" & ","
        slRecord = slRecord & """" & "Net Per 2" & """" & ","
        slRecord = slRecord & """" & "Net Per 3" & """" & ","
        slRecord = slRecord & """" & "Net Per 4" & """" & ","
        slRecord = slRecord & """" & "Net Per 5" & """" & ","
        slRecord = slRecord & """" & "Net Per 6" & """" & ","
        slRecord = slRecord & """" & "Net Per 7" & """" & ","
        slRecord = slRecord & """" & "Net Per 8" & """" & ","
        slRecord = slRecord & """" & "Net Per 9" & """" & ","
        slRecord = slRecord & """" & "Net Per 10" & """" & ","
        slRecord = slRecord & """" & "Net Per 11" & """" & ","
        slRecord = slRecord & """" & "Net Per 12" & """" & ","
        slRecord = slRecord & """" & "TNet Per 1" & """" & ","
        slRecord = slRecord & """" & "TNet Per 2" & """" & ","
        slRecord = slRecord & """" & "TNet Per 3" & """" & ","
        slRecord = slRecord & """" & "TNet Per 4" & """" & ","
        slRecord = slRecord & """" & "TNet Per 5" & """" & ","
        slRecord = slRecord & """" & "TNet Per 6" & """" & ","
        slRecord = slRecord & """" & "TNet Per 7" & """" & ","
        slRecord = slRecord & """" & "TNet Per 8" & """" & ","
        slRecord = slRecord & """" & "TNet Per 9" & """" & ","
        slRecord = slRecord & """" & "TNet Per 10" & """" & ","
        slRecord = slRecord & """" & "TNet Per 11" & """" & ","
        slRecord = slRecord & """" & "TNet Per 12" & """" & ","
        slRecord = slRecord & """" & "Type" & """"
        Print #hmTo, slRecord

        slDateGenned = Format$(gNow(), "m/d/yy")
        slTimeGenned = Format$(gNow(), "h:mm:ssAM/PM")
        gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llDate
        slLastInvDate = Format(llDate, "m/d/yy")
        If rbcYearType(0).Value Then
            slYearType = "Corporate"
        Else
            slYearType = "Standard"
        End If

        slRecord = Trim$(tgSpf.sGClient) & " "
        slRecord = slRecord & "Generated: " & Trim$(slDateGenned) & " @"
        slRecord = slRecord & Trim$(slTimeGenned) & " for "
        slRecord = slRecord & Trim$(slYearType) & " "
        slRecord = slRecord & Trim$(edcSelCFrom.Text) & " " & Trim$(edcSelCFrom1.Text) & " "
        slRecord = slRecord & "Last Bill Date: " & Trim$(slLastInvDate)

        Print #hmTo, slRecord
        Exit Sub

mCreateHeaderErr:
        gDbg_HandleError "Messages: mCreateHeader"
        Exit Sub

End Sub

Public Sub mFutureRevenue(llStartDates() As Long, llLastBilled As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilMatchSOFCode                ilTemp                        tmOneVefAllYear           *
'*                                                                                        *
'******************************************************************************************

Dim ilCurrentRecd As Long       'current index to process from tlChfAdvtExt array
Dim ilLoop As Integer
Dim slStr As String
Dim ilMatchCntr As Integer
Dim ilMatchSSCode As Integer
Dim ilRet As Integer
Dim slTempStart As String       'earliest date to obtain contracts: defaulted to 30 days prior to last date billed
Dim slTempEnd As String         'latest date to obtain contracts; defaulted to 90 days past 12th months requested
Dim slCntrStatus As String      'contract status to include:  Sch Hold/Order, Unsch Hold/Order
Dim slCntrType As String        'contract types to include:  all except PSA/Promo
Dim ilHOState As Integer        'State of contract, most recent revision
Dim llContrCode As Long         'contract code to read contract
Dim ilAdjust As Integer
Dim ilAdjustMissedForMG As Integer  'default to true
Dim ilFirstProjInx As Integer       'index of first month to project (1-12)
Dim slPctTrade As String        '% of trade : ie. 40.00
Dim ilStartCorT As Integer      '1 = cash, 2 = trade
Dim ilEndCorT As Integer        '1 = cash, 2 = trade
Dim tlSBFType As SBFTypes       'types of SBF records to retrieve:  NTR, Hard copy, Installment
Dim slStdStart As String        'first date of month requested
Dim llStdStart As Long          'first date of month requested
Dim slStdEnd As String          'last date of 12 months requested
Dim llStdEnd As Long            'last date of 12 months requested
ReDim tlSbf(0 To 0) As SBF      'array of NTR / installment records
Dim ilSubMissed As Integer      'subtract missed ; default to false
Dim ilCountMGs As Integer       'count mgs where they air:  default to true
Dim tlChfAdvtExt() As CHFADVTEXT    'array of contracts based on first date to project thru end of the 12 months
Dim ilContinue As Integer
Dim ilSaveNTRFlag As Integer
Dim ilVehLoop As Integer
Dim tlSplitInfo As splitinfo    'info from spot adjustments, NTR records or sch line to send to common rtn
Dim slCashAgyComm As String     'agy commission % (i.e. 15.00, 00.00)
Dim ilClf As Integer            'schedule line index
Dim llProjectSpots(0 To 12) As Long 'unused, reqd for flight subrtn, index zero ignored
Dim llProjectRC(0 To 12) As Long    'unused, reqd for flight subrtn, index zero ignored
ReDim tmCntAllYear(0 To 1) As ALLPIFPCTYEAR     'participant info for contract, index zero ignored

        tlSBFType.iNTR = True
        tlSBFType.iInstallment = False
        tlSBFType.iImport = False

        slCntrStatus = "HOGN"           'statuses: hold, order, unsch hold, uns order
        slCntrType = ""
        ilHOState = 2

        'default subtract missed spots and count mg where they air
        ilSubMissed = False          'subt misses
        ilCountMGs = True            'count mgs
        smAirOrder = tgSpf.sInvAirOrder         'bill as ordred, aired
        If smAirOrder = "S" Or rbcYearType(0).Value Then      'bill as ordered (update as order) or Corporate, no adjustments for makegoods/missed
            ilAdjust = False                    'always ignore missed and makegoods
        Else
            ilAdjust = True                     'missed or mg must be adjusted
        End If

        ilAdjustMissedForMG = True
        If rbcYearType(0).Value Then           'corp , look at all orders for entire year
            ilFirstProjInx = 1
        Else       'std, decide what month to start looking at orders for future
            '7-12-01 Billed & Booked Slsp Comm comes thru here.  Do nothing with that report for now.
            If (tgSpf.iPkageGenMeth = 1) Then           'calc virtual pkg by line & use airing Lines?
                ilFirstProjInx = 1                                   'if so, force to gather all information for contract because of balancing issue with receivables
            Else
                For ilLoop = 1 To 12 Step 1
                    If llLastBilled > llStartDates(ilLoop) And llLastBilled < llStartDates(ilLoop + 1) Then
                        ilFirstProjInx = ilLoop + 1
                        slStdStart = Format$(llStartDates(ilFirstProjInx), "m/d/yy")
                        Exit For
                    End If
                Next ilLoop
                If ilFirstProjInx = 0 Then
                    ilFirstProjInx = 1                          'all projections, no actuals
                End If
                If llLastBilled >= llStartDates(13) Then   'all data was in the past only, dont do contracts
                    Exit Sub
                End If
            End If
        End If

        slStdStart = Format$(llStartDates(1), "m/d/yy")       'assume first date of proj is the quarter entered
        slStdEnd = Format$(llStartDates(13), "m/d/yy")

        llStdStart = llStartDates(ilFirstProjInx)  'first date to project
        llStdEnd = llStartDates(13)                'end date to project
        slTempStart = Format$((gDateValue(slStdStart) - 30), "m/d/yy")
        slTempEnd = Format$((gDateValue(slStdEnd) + 90), "m/d/yy")

        'obtain contracts to process
        ilRet = gObtainCntrForDate(ExpRevenue, slTempStart, slTempEnd, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())

        For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1

            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChf, tgClfCT(), tgCffCT())
            tmChf = tmChf
            ilMatchCntr = True

            'exclude psa/promo, contracts not fully or manually sched or proposals
            If (lmSingleCntr > 0 And lmSingleCntr <> tmChf.lCntrNo) Or tmChf.sType = "M" Or tmChf.sType = "S" Then
                ilMatchCntr = False
            End If

            If ilMatchCntr Then
                mSetupCTAgyComm slCashAgyComm, slPctTrade, ilStartCorT, ilEndCorT

                ilContinue = False
                ilSaveNTRFlag = tlSBFType.iNTR
                ReDim tmSBFAdjust(0 To 0) As ADJUSTLIST             'build new for every contract
                If (tmChf.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) <> INSTALLMENTREVENUEEARNED) Then     'its an install contract, if install method is bill as aired, get the NTRs from SBF                                                                 'if install bill method is Invoiced, get the installment records from SBF
                    'install method is invoiced, get the future from Installment records
                    tlSBFType.iNTR = False
                    tlSBFType.iInstallment = True
                    ilContinue = True
                    ilRet = gObtainSBF(ExpRevenue, hmSbf, tmChf.lCode, slStdStart, slStdEnd, tlSBFType, tlSbf(), 0)
                    'Build array of the vehicles and their NTR $ into tmSBFAdjust array
                    gSbfAdjustForInstall tlSbf(), tmSBFAdjust(), llStartDates(), ilFirstProjInx, llStdStart, llStdEnd, 12, tmChf.iAgfCode
                Else            'all NTR items in the future
                    ilRet = gObtainSBF(ExpRevenue, hmSbf, tmChf.lCode, slStdStart, slStdEnd, tlSBFType, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use
                    'Build array of the vehicles and their NTR $ into tmSBFAdjust array
                    gSbfAdjustForNTR tlSbf(), tmSBFAdjust(), llStartDates(), ilFirstProjInx, llStdStart, llStdEnd, 12, tmNTRMNF()
                    ilContinue = True
                End If
                If ilContinue Then
                    'loop on sbf to process each $ and split by slsp, then participant, then cash / trade
                    For ilVehLoop = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1      '11-06-06 chg to ubound -1
                        tlSplitInfo.iStartCorT = ilStartCorT
                        tlSplitInfo.iEndCorT = ilEndCorT
                        tlSplitInfo.iFirstProjInx = ilFirstProjInx
                        tlSplitInfo.iVefCode = tmSBFAdjust(ilVehLoop).iVefCode

                        tlSplitInfo.sPctTrade = slPctTrade
                        If ((tmChf.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED)) Or tmChf.sInstallDefined <> "Y" Then     'its an install contract, if install method is bill as aired (separate inv from revenue), get the NTRs from SBF                                                                 'if install bill method is Invoiced, get the installment records from SBF
                            'this is revenue for NTR / Hard Cost in the future
                            tlSplitInfo.sTradeAgyComm = tmSBFAdjust(ilVehLoop).sAgyComm
                            If tmSBFAdjust(ilVehLoop).sAgyComm = "Y" Then
                                tlSplitInfo.sCashAgyComm = slCashAgyComm
                            Else
                                tlSplitInfo.sCashAgyComm = "00.00"
                            End If
                            tlSplitInfo.sNTR = "Y"
                            tlSplitInfo.iHardCost = tmSBFAdjust(ilVehLoop).iIsItHardCost      'true false if hard cost
                            tlSplitInfo.iMnfNTRItemCode = tmSBFAdjust(ilVehLoop).iMnfItem       '11-06-06 NTR Item type from MNF
                        Else           'otherwise its billed as invoiced and need to get future from NTR and lines
                            'this is SBF records for Installment in the future
                            'tlSplitInfo.sCashAgyComm = slCashAgyComm
                            If tmSBFAdjust(ilVehLoop).sAgyComm = "Y" Then
                                tlSplitInfo.sCashAgyComm = slCashAgyComm
                            Else
                                tlSplitInfo.sCashAgyComm = "00.00"
                            End If
                            tlSplitInfo.sTradeAgyComm = tmChf.sAgyCTrade        '12-23-02
                            tlSplitInfo.sNTR = "N"
                            tlSplitInfo.iHardCost = False
                            tlSplitInfo.iMnfNTRItemCode = 0         '11-06-06 not an NTR
                        End If

                        For ilLoop = 1 To 12
                            lmProjectGross(ilLoop) = tmSBFAdjust(ilVehLoop).lProject(ilLoop)
                            lmProjectTNet(ilLoop) = tmSBFAdjust(ilVehLoop).lAcquisitionCost(ilLoop)
                            'calc net if agency commissionable or its trade with agency commissionable
                            If (tmChf.iAgfCode > 0 And tmSBFAdjust(ilVehLoop).sAgyComm = "Y") Then
                                slStr = gLongToStrDec(lmProjectGross(ilLoop), 2)
                                slStr = gRoundStr(gMulStr(slStr, gSubStr("100.00", slCashAgyComm)), "1", 0)
                                lmProjectNet(ilLoop) = Val(slStr)
                            Else        'no agency comm
                                lmProjectNet(ilLoop) = lmProjectGross(ilLoop)
                            End If
                        Next ilLoop
                        mSetupSlspForFuture tlSplitInfo, llStartDates()     'distribute $ for slsp/participant
                    Next ilVehLoop
                End If
                tlSBFType.iNTR = ilSaveNTRFlag
                tlSBFType.iInstallment = False  'installment flag to retrieve SBF records are only set when getting installment records for the future based on install method "Invoiced"

                'Insure the common monthly buckets are initialized for the schedule lines
                For ilLoop = 1 To 12
                    lmProjectGross(ilLoop) = 0
                    lmProjectNet(ilLoop) = 0
                    lmProjectTNet(ilLoop) = 0
                Next ilLoop

                'process the schedule lines if air time included and its not a installment contract, or
                'air time included and its an installment method is Aired (separate invoicing from revenue); get the
                'future from schedule lines.  If installment whose method is invoiced (inv = revenue),
                If (tmChf.sInstallDefined <> "Y") Or (tmChf.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED) Then                    '11-25-02 include air time (vs NTR)?

                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        tmClf = tgClfCT(ilClf).ClfRec

                        'ignore any type of package lines
                        If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
                            gBuildFlightInfo ilClf, llStartDates(), ilFirstProjInx, 13, lmProjectGross(), llProjectSpots(), llProjectRC(), lmProjectTNet(), 1, tgClfCT(), tgCffCT()

                            'Adjust with misses and makegoods for the line?  No adjustments for corporate option
                            If ilAdjust Then        'retain the miss/mgs in a tmAdjustList array
                                mAdjustMissedMGs llStartDates(), ilFirstProjInx, llStdStart, llStdEnd, ilSubMissed, ilCountMGs, 12, ilAdjustMissedForMG
                            End If

                            'Distribute the sched line $
                            tlSplitInfo.iMatchSSCode = ilMatchSSCode
                            tlSplitInfo.iStartCorT = ilStartCorT
                            tlSplitInfo.iEndCorT = ilEndCorT
                            tlSplitInfo.iFirstProjInx = ilFirstProjInx
                            tlSplitInfo.iVefCode = tmClf.iVefCode
                            tlSplitInfo.sPctTrade = slPctTrade
                            tlSplitInfo.sCashAgyComm = slCashAgyComm
                            tlSplitInfo.sTradeAgyComm = tmChf.sAgyCTrade
                            tlSplitInfo.sNTR = "A"
                            tlSplitInfo.iHardCost = False
                            tlSplitInfo.iMnfNTRItemCode = 0

                            For ilLoop = 1 To 12
                                'lmProjectGross & lmprojectTNet have are distributed from schedule line
                                'calc net if agency commissionable or its trade with agency commissionable
                                If tmChf.iAgfCode > 0 Then
                                    slStr = gLongToStrDec(lmProjectGross(ilLoop), 2)
                                    slStr = gRoundStr(gMulStr(slStr, gSubStr("100.00", slCashAgyComm)), "1", 0)
                                    lmProjectNet(ilLoop) = Val(slStr)
                                Else        'no agency comm
                                    lmProjectNet(ilLoop) = lmProjectGross(ilLoop)
                                End If
                            Next ilLoop
                            mSetupSlspForFuture tlSplitInfo, llStartDates()

                            'process the makegoods for this line - there can be multiple vehicles that spots were moved  to.  Each
                            'new vehicle must be retrieved and it's splits determined
                            For ilVehLoop = 0 To imUpperAdjust - 1 Step 1
                                If tmAdjust(ilVehLoop).iVefCode > 0 Then
                                    tlSplitInfo.iMatchSSCode = ilMatchSSCode
                                    tlSplitInfo.iStartCorT = ilStartCorT
                                    tlSplitInfo.iEndCorT = ilEndCorT
                                    tlSplitInfo.iFirstProjInx = ilFirstProjInx
                                    tlSplitInfo.iVefCode = tmAdjust(ilVehLoop).iVefCode
                                    tlSplitInfo.sPctTrade = slPctTrade
                                    tlSplitInfo.sCashAgyComm = slCashAgyComm
                                    tlSplitInfo.sTradeAgyComm = tmChf.sAgyCTrade
                                    tlSplitInfo.sNTR = "A"              'airtime vs NTR
                                    tlSplitInfo.iHardCost = False
                                    tlSplitInfo.iMnfNTRItemCode = 0
                                    For ilLoop = 1 To 12
                                        lmProjectGross(ilLoop) = tmAdjust(ilVehLoop).lProject(ilLoop)
                                        'calc net if agency commissionable or its trade with agency commissionable
                                        If (tmChf.iAgfCode > 0) Then
                                            slStr = gLongToStrDec(lmProjectGross(ilLoop), 2)
                                            slStr = gRoundStr(gMulStr(slStr, gSubStr("100.00", slCashAgyComm)), "1", 0)
                                            lmProjectNet(ilLoop) = Val(slStr)
                                        Else        'no agency comm
                                            lmProjectNet(ilLoop) = lmProjectGross(ilLoop)
                                        End If
                                    Next ilLoop
                                End If
                                mSetupSlspForFuture tlSplitInfo, llStartDates()     'distribute $ for slsp/participant
                            Next ilVehLoop
                            ReDim tmAdjust(0 To 0) As ADJUSTLIST             'prepare list of mgs
                            imUpperAdjust = 0
                        End If
                    Next ilClf                          'next schedule line - for advt & slsp the entire contr is accumulated before writing to GRF
                End If                              'include Air Time vs NTR
            End If                                  'exclude all promotions, merchandising - only include trade if requested
         Next ilCurrentRecd          '03-13-01

        Erase tlChfAdvtExt, tlSbf

        Exit Sub
End Sub
'
'               mSetupSlspForFuture - determine how many slsp splits
'               to share with the contract.  Determine if using Sub-company
'               and obtain the slsp for the proper sub-company
'
Private Sub mSetupSlspForFuture(tlSplitInfo As splitinfo, llStartDates() As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoopOn12Months              llGrossDollar                 llNetDollar               *
'*  llTNetDollar                                                                          *
'******************************************************************************************

ReDim ilSlfCode(0 To 9) As Integer
ReDim llSlfSplit(0 To 9) As Long
ReDim llSlfSplitRev(0 To 9) As Long     'not used in this app, but needed for common routine
Dim ilHowManyDefined As Integer
Dim ilTempSlf() As Integer
Dim llTempSlfSplit() As Long
Dim ilTemp As Integer
Dim ilUseSlsComm As Integer
Dim ilMatchSSCode As Integer
Dim ilLoopOnSlsp As Integer
Dim llTempPct(0 To 12) As Long  'index zero ignored
Dim ilStartCorT As Integer
Dim ilEndCorT As Integer
Dim ilCorT As Integer
Dim slCTSplit As String
Dim llCTGross As Long
Dim llCTNet As Long
Dim llCTTNet As Long
Dim slDollar As String
Dim slGross As String
Dim slNet As String
Dim slTNet As String
Dim slCashAgyComm As String
Dim ilMnfSubCo As Integer
Dim ilFound As Integer
Dim ilLoop As Integer
Dim slPctTrade As String
Dim ilSlfRecd As Integer
Dim ilLoopOnMonths As Integer
Dim ilLoopOnParts As Integer


        'set the slsp used for this contract.  If subcompany used, more than 1 subcompany can be defined on a contract.
        'Be sure to use the correct vehicles slsp for the proper subcompany
        ilUseSlsComm = False
        ilMnfSubCo = gGetSubCmpy(tmChf, ilSlfCode(), llSlfSplit(), tlSplitInfo.iVefCode, ilUseSlsComm, llSlfSplitRev())                                         '4-6-00
        'ilslfCode() has the array of slsp matching the vehicles subcompany (if applicable); otherwise,
        'its the same list of slsp defined in the header
        'some of the Slsp do not have any percentages defined, so they must be ignored.
        'when splitting $, extra pennies may exist so the last slsp with a % defined
        'must get the extra pennies.

        ReDim ilTempSlf(0 To 0) As Integer
        ReDim llTempSlfSplit(0 To 0) As Long
        ilHowManyDefined = 0
        'determine the actual number of slsp to process
        'check for both slsp code and split % because they may not all be defined
        'with percentages
        'i.e. slsp A:  50%
        '      slsp b:  0
        '      slsp c:  50%
        For ilTemp = 0 To 9
            If ilSlfCode(ilTemp) > 0 And llSlfSplit(ilTemp) > 0 Then        'both slsp & % split defined
                ilTempSlf(ilHowManyDefined) = ilSlfCode(ilTemp)
                llTempSlfSplit(ilHowManyDefined) = llSlfSplit(ilTemp)
                ilHowManyDefined = ilHowManyDefined + 1
                ReDim Preserve ilTempSlf(0 To ilHowManyDefined) As Integer
                ReDim Preserve llTempSlfSplit(0 To ilHowManyDefined) As Long
            End If
        Next ilTemp

        If ilHowManyDefined = 0 Then        'nothing found with slsp %, force to first slsp at 100%
            ilTempSlf(0) = ilSlfCode(0)
            llTempSlfSplit(0) = 10000
            ReDim Preserve ilTempSlf(0 To 1) As Integer
            ReDim Preserve llTempSlfSplit(0 To 1) As Long
        End If

        ilStartCorT = tlSplitInfo.iStartCorT
        ilEndCorT = tlSplitInfo.iEndCorT
        slPctTrade = tlSplitInfo.sPctTrade
        slCashAgyComm = tlSplitInfo.sCashAgyComm

        'save the original numbers so whatever portion of the $ are distributed are given to the last slsp/participant
        For ilLoopOnMonths = 1 To 12
            lmSlspRemGross(ilLoopOnMonths) = lmProjectGross(ilLoopOnMonths)
            lmSlspRemNet(ilLoopOnMonths) = lmProjectNet(ilLoopOnMonths)
            lmSlspRemTNet(ilLoopOnMonths) = lmProjectTNet(ilLoopOnMonths)
        Next ilLoopOnMonths
        'create as many as 10 slsp records; each split by the number of participants
        'for each vehicle
        'For slsp, create as many as 10 records per trans (up to 10 split slsp) per trans.
        For ilLoopOnSlsp = 0 To UBound(ilTempSlf) - 1 Step 1           'loop based on report option

            ilMatchSSCode = 0
            For ilSlfRecd = LBound(tgMSlf) To UBound(tgMSlf)
                If ilTempSlf(ilLoopOnSlsp) = tgMSlf(ilSlfRecd).iCode Then
                    tmSlf = tgMSlf(ilSlfRecd)
                    ilMatchSSCode = mFindMatchSS(tmSofList(), tmSlf.iSofCode)       'get the matching sales souce
                    'may need the subcompany to determine which slsp to pick up for the transactions vehicle
                    Exit For
                End If
            Next ilSlfRecd
            If ilMatchSSCode = 0 Then           'no more slsp to process as no matching sales source
                Exit For
            End If

            gInitCntPartYear hmVsf, tmChf, ilMatchSSCode, llStartDates(), tmCntAllYear(), tmPifKey(), tmPifPct()

            'Needed to get the sales source before we can get the participants
            'Get participants for the vehicle.  They are dated and may have varying revenue splits.
            ilFound = gInitVehAllYearPcts(ilMatchSSCode, tlSplitInfo.iVefCode, tmCntAllYear(), tmOneVehAllYear()) 'only those participants for the contracts matching SS have
            'been built into the tmCntAllYear array.  Each entry is for a different participant of the sales source
            If Not ilFound Then     'no matching vehicle found,must be a mg or outside vehicle that is not on the contract
                gGetOneVehAllYearForMG tlSplitInfo.iVefCode, llStartDates(), ilMatchSSCode, tmPifKey(), tmPifPct(), tmCntAllYear()
                ilFound = gInitVehAllYearPcts(ilMatchSSCode, tlSplitInfo.iVefCode, tmCntAllYear(), tmOneVehAllYear()) 'only those participants for the contracts matching SS have
                If Not ilFound Then         'force to 100%
                    tmOneVehAllYear(LBound(tmOneVehAllYear)).AllYear.iVefCode = tlSplitInfo.iVefCode
                    tmOneVehAllYear(LBound(tmOneVehAllYear)).AllYear.iSSMnfCode = ilMatchSSCode
                    For ilLoop = 1 To 12
                        tmOneVehAllYear(LBound(tmOneVehAllYear)).AllYear.iPct(ilLoop) = 10000
                    Next ilLoop
                    tmOneVehAllYear(LBound(tmOneVehAllYear)).AllYear.iMnfGroup = 0        'unknown participant
                End If
            End If
            'slsp have same rev splits all year (no dated revenue splits)
            For ilLoopOnMonths = 1 To 12
                llTempPct(ilLoopOnMonths) = llTempSlfSplit(ilLoopOnSlsp)
            Next ilLoopOnMonths

            'split the Slsp $, send the % split, original gross, net, Tnet amounts.  returned gross, net and tnet values based on %
            mSplitFuture llTempPct(), lmProjectGross(), lmProjectNet(), lmProjectTNet(), lmGrossDollar(), lmNetDollar(), lmTNetDollar()

            'accumulate the $ amount distributed so that whatever is remaining is also distributed
            For ilLoopOnMonths = 1 To 12
                'see whats left after $ have been distributed for last slsp
                lmSlspRemGross(ilLoopOnMonths) = lmSlspRemGross(ilLoopOnMonths) - lmGrossDollar(ilLoopOnMonths)
                lmSlspRemNet(ilLoopOnMonths) = lmSlspRemNet(ilLoopOnMonths) - lmNetDollar(ilLoopOnMonths)
                lmSlspRemTNet(ilLoopOnMonths) = lmSlspRemTNet(ilLoopOnMonths) - lmTNetDollar(ilLoopOnMonths)
            Next ilLoopOnMonths

            'adjust for last slsp, remaining pennies
            If ilLoopOnSlsp = UBound(ilTempSlf) - 1 Then
                For ilLoopOnMonths = 1 To 12
                    lmGrossDollar(ilLoopOnMonths) = lmGrossDollar(ilLoopOnMonths) + lmSlspRemGross(ilLoopOnMonths)
                    lmNetDollar(ilLoopOnMonths) = lmNetDollar(ilLoopOnMonths) + lmSlspRemNet(ilLoopOnMonths)
                    lmTNetDollar(ilLoopOnMonths) = lmTNetDollar(ilLoopOnMonths) + lmSlspRemTNet(ilLoopOnMonths)
                Next ilLoopOnMonths
            End If

            'lmPart fields are the running totals minus each participant split to see what left after all slsp have
            'been processed.  Need to give last participant the extra pennies
            For ilLoopOnMonths = 1 To 12
                lmPartRemGross(ilLoopOnMonths) = lmGrossDollar(ilLoopOnMonths)
                lmPartRemNet(ilLoopOnMonths) = lmNetDollar(ilLoopOnMonths)
                lmPartRemTNet(ilLoopOnMonths) = lmTNetDollar(ilLoopOnMonths)
            Next ilLoopOnMonths


            For ilLoopOnParts = LBound(tmOneVehAllYear) To UBound(tmOneVehAllYear)
                tmOnePartAllYear = tmOneVehAllYear(ilLoopOnParts).AllYear      'someparticipants are not defined properly, cant do -1
                If tmOnePartAllYear.iMnfGroup > 0 Then      'ignore if null participant
                    'current participants split for 12 months
                    For ilLoopOnMonths = 1 To 12
                        llTempPct(ilLoopOnMonths) = tmOnePartAllYear.iPct(ilLoopOnMonths)
                        llTempPct(ilLoopOnMonths) = llTempPct(ilLoopOnMonths) * 100
                    Next ilLoopOnMonths

                    'Split the Participant $
                    mSplitFuture llTempPct(), lmGrossDollar(), lmNetDollar(), lmTNetDollar(), lmPartGross(), lmPartNet(), lmPartTNet()
                    'see whats left after $ have been distributed for last participant
                    For ilLoopOnMonths = 1 To 12
                        lmPartRemGross(ilLoopOnMonths) = lmPartRemGross(ilLoopOnMonths) - lmPartGross(ilLoopOnMonths)
                        lmPartRemNet(ilLoopOnMonths) = lmPartRemNet(ilLoopOnMonths) - lmPartNet(ilLoopOnMonths)
                        lmPartRemTNet(ilLoopOnMonths) = lmPartRemTNet(ilLoopOnMonths) - lmPartTNet(ilLoopOnMonths)
                    Next ilLoopOnMonths


                    For ilLoopOnMonths = 1 To 12
                        For ilCorT = ilStartCorT To ilEndCorT
                            If ilCorT = 1 Then
                                slCTSplit = gSubStr("100.", slPctTrade)
                                slDollar = gLongToStrDec(lmPartGross(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slGross = gRoundStr(slDollar, "1", 0)
                                'slNet = gRoundStr(gDivStr(gMulStr(slGross, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
                                slDollar = gLongToStrDec(lmPartNet(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slNet = gRoundStr(slDollar, "1", 0)

                                'T-Net amount
                                slDollar = gLongToStrDec(lmPartTNet(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slTNet = gRoundStr(slDollar, "1", 0)

                            Else
                                slCTSplit = slPctTrade
                                slDollar = gLongToStrDec(lmPartGross(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slGross = gRoundStr(slDollar, "1", 0)
                                'slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
                                slDollar = gLongToStrDec(lmPartNet(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slNet = gRoundStr(slDollar, "1", 0)

                                'T-Net amount
                                slDollar = gLongToStrDec(lmPartTNet(ilLoopOnMonths), 0)
                                slDollar = gDivStr(gMulStr(slDollar, slCTSplit), "100")              'slsp gross
                                slTNet = gRoundStr(slDollar, "1", 0)

                            End If
                            llCTGross = Val(slGross)
                            llCTNet = Val(slNet)
                            llCTTNet = Val(slTNet)
                            'adjust for last participant, remaining pennies
                            If ilLoopOnParts = UBound(tmOneVehAllYear) - 1 Then
                                llCTGross = llCTGross + lmPartRemGross(ilLoopOnMonths)
                                llCTNet = llCTNet + lmPartRemNet(ilLoopOnMonths)
                                llCTTNet = llCTTNet + lmPartRemTNet(ilLoopOnMonths)
                            End If

                            tmInputInfo.iSlfCode = tmSlf.iCode
                            tmInputInfo.iMnfCode = tmOneVehAllYear(ilLoopOnParts).AllYear.iMnfGroup             'participant code
                            tmInputInfo.iVefCode = tlSplitInfo.iVefCode
                            gGetVehGrpSets tmInputInfo.iVefCode, imMinorVG, imMajorVG, ilTemp, tmInputInfo.iVefGroup
                            If imMajorVG = 1 Then               'participant vehicle group selected
                                tmInputInfo.iVefGroup = tmInputInfo.iMnfCode
                            End If

                            If ilCorT = 1 Then
                                tmInputInfo.sCashTrade = "C"
                                tmInputInfo.lCashGross(ilLoopOnMonths) = tmInputInfo.lCashGross(ilLoopOnMonths) + llCTGross
                                tmInputInfo.lCashNet(ilLoopOnMonths) = tmInputInfo.lCashNet(ilLoopOnMonths) + llCTNet
                                tmInputInfo.lCashTNet(ilLoopOnMonths) = tmInputInfo.lCashTNet(ilLoopOnMonths) + llCTTNet
                            Else    'ilCorT = 2
                                tmInputInfo.sCashTrade = "T"
                                tmInputInfo.lTradeGross(ilLoopOnMonths) = tmInputInfo.lTradeGross(ilLoopOnMonths) + llCTGross
                                tmInputInfo.lTradeNet(ilLoopOnMonths) = tmInputInfo.lTradeNet(ilLoopOnMonths) + llCTNet
                                tmInputInfo.lTradeTNet(ilLoopOnMonths) = tmInputInfo.lTradeTNet(ilLoopOnMonths) + llCTTNet
                            End If
                            If tlSplitInfo.sNTR = "Y" And Not tlSplitInfo.iHardCost Then
                                tmInputInfo.sAirTimeNTR = "NTR"
                            ElseIf tlSplitInfo.sNTR = "Y" And tlSplitInfo.iHardCost Then
                                tmInputInfo.sAirTimeNTR = "HardCost"           'air time/ntr flag
                            Else
                                tmInputInfo.sAirTimeNTR = "Airtime"
                            End If
                        Next ilCorT
                    Next ilLoopOnMonths
                    mCreateExportRecord                 'create a record for 1 slsp/participant for the year (max 2 records, 1 for cash, 1 for trade)
                End If
            Next ilLoopOnParts

        Next ilLoopOnSlsp                             'loop for 10 slsp possible splits, or 3 possible owners, otherwise loop once
        For ilLoopOnMonths = 1 To 12
            lmSlspRemGross(ilLoopOnMonths) = 0
            lmSlspRemNet(ilLoopOnMonths) = 0
            lmSlspRemTNet(ilLoopOnMonths) = 0
            lmPartRemGross(ilLoopOnMonths) = 0
            lmPartRemNet(ilLoopOnMonths) = 0
            lmPartRemTNet(ilLoopOnMonths) = 0
            lmProjectGross(ilLoopOnMonths) = 0
            lmProjectNet(ilLoopOnMonths) = 0
            lmProjectTNet(ilLoopOnMonths) = 0
            'init 1 Salespersons share
            lmGrossDollar(ilLoopOnMonths) = 0
            lmNetDollar(ilLoopOnMonths) = 0
            lmTNetDollar(ilLoopOnMonths) = 0
            'init 1 participants share
            lmPartGross(ilLoopOnMonths) = 0
            lmPartNet(ilLoopOnMonths) = 0
            lmPartTNet(ilLoopOnMonths) = 0
        Next ilLoopOnMonths
        Erase ilTempSlf, llTempSlfSplit
        Exit Sub
End Sub
'               mSetupCTAgyComm - determine if contract is agency commissionable.  If it is, it will
'               apply to air time schedule lines as NTR has its own agency comm flag
'               <input> - none (assume contract tmChf has been read)
'               <output> slCashAgycomm = percent of agency comm (i.e. 15.00, 0.00)
'                        ilCorT: 1 = cash only or combined cash/trade)
'                        ilCorT:  2 = trade only or combined cash/trade
'
Private Sub mSetupCTAgyComm(slCashAgyComm As String, slPctTrade As String, ilStartCorT As Integer, ilEndCorT As Integer)
Dim ilRet As Integer

                If (tmChf.iAgfCode > 0) Then   'if direct advert or showing acquisition costs, dont take any agency comm out
                    tmAgfSrchKey0.iCode = tmChf.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                    Else
                        slCashAgyComm = ".00"
                    End If
                Else
                    slCashAgyComm = ".00"
                End If

                slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                If tmChf.iPctTrade = 0 Then                     'setup loop to do cash & trade
                    ilStartCorT = 1
                    ilEndCorT = 1
                ElseIf tmChf.iPctTrade = 100 Then
                    ilStartCorT = 2
                    ilEndCorT = 2
                Else
                    ilStartCorT = 1
                    ilEndCorT = 2
                End If
End Sub
