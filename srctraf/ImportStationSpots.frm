VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form ImportStationSpots 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin VB.Frame frcCntr 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3210
      Left            =   -90
      TabIndex        =   20
      Top             =   3555
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CommandButton cmcCntrMatch 
         Appearance      =   0  'Flat
         Caption         =   "&Match"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3285
         TabIndex        =   22
         Top             =   2880
         Width           =   1005
      End
      Begin VB.CommandButton cmcCntrCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   5055
         TabIndex        =   21
         Top             =   2880
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCntrStation 
         Height          =   570
         Left            =   60
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   195
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1005
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCntrNetwork 
         Height          =   1860
         Left            =   60
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   915
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   10
         Cols            =   10
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin VB.Label lacPossibleCntr 
         Alignment       =   2  'Center
         Caption         =   "Unposted and Manually Posted Contracts Sold"
         Height          =   180
         Left            =   1530
         TabIndex        =   29
         Top             =   750
         Width           =   4410
      End
      Begin VB.Label lacSelectedUnresolvedImport 
         Alignment       =   2  'Center
         Caption         =   "Imported Station Invoice- Unresolved"
         Height          =   180
         Left            =   2160
         TabIndex        =   28
         Top             =   0
         Width           =   4125
      End
   End
   Begin VB.TextBox edcProcessing 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   705
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1770
      Visible         =   0   'False
      Width           =   8085
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   495
      Top             =   3990
   End
   Begin VB.Frame frcDetail 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3210
      Left            =   120
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   9030
      Begin VB.CommandButton cmcReconcile 
         Appearance      =   0  'Flat
         Caption         =   "Reconcile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4890
         TabIndex        =   14
         Top             =   1710
         Width           =   1440
      End
      Begin VB.CommandButton cmcUndoReconcile 
         Appearance      =   0  'Flat
         Caption         =   "Undo Reconcile"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6615
         TabIndex        =   13
         Top             =   2880
         Width           =   1440
      End
      Begin VB.CommandButton cmcReturn 
         Appearance      =   0  'Flat
         Caption         =   "&Return"
         Height          =   285
         Left            =   4710
         TabIndex        =   11
         Top             =   2895
         Width           =   1440
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNetworkSpots 
         Height          =   1395
         Left            =   60
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   2461
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStationSpots 
         Height          =   1380
         Left            =   4590
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   2434
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMatchedSpots 
         Height          =   1050
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1800
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   1852
         _Version        =   393216
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   13
      End
      Begin VB.Label lacDetail 
         Alignment       =   2  'Center
         Height          =   210
         Left            =   60
         TabIndex        =   19
         Top             =   -15
         Width           =   8295
      End
      Begin VB.Label lacReconciledSpots 
         Caption         =   "Reconciled Spots"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1590
         Width           =   4395
      End
      Begin VB.Label lacStnSpots 
         Caption         =   "Unreconciled Station Spots"
         Height          =   210
         Left            =   4590
         TabIndex        =   17
         Top             =   75
         Width           =   4395
      End
      Begin VB.Label lacNetSpots 
         Caption         =   "Unreconciled Network Spots"
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   75
         Width           =   4395
      End
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   165
      Left            =   9045
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   3885
      Width           =   75
   End
   Begin VB.Frame frcResult 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3360
      Left            =   75
      TabIndex        =   3
      Top             =   300
      Width           =   9195
      Begin VB.CommandButton cmcUndoMatch 
         Appearance      =   0  'Flat
         Caption         =   "&Undo Match"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6180
         TabIndex        =   25
         Top             =   2805
         Width           =   1155
      End
      Begin VB.CommandButton cmcCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   285
         Left            =   4605
         TabIndex        =   10
         Top             =   2805
         Width           =   1155
      End
      Begin VB.CommandButton cmcDetail 
         Appearance      =   0  'Flat
         Caption         =   "&Detail"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2835
         TabIndex        =   9
         Top             =   2805
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMatchedResult 
         Height          =   1890
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1425
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   3334
         _Version        =   393216
         Rows            =   10
         Cols            =   18
         FixedRows       =   2
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   18
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNoMatchedResult 
         Height          =   975
         Left            =   75
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   210
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   10
         Cols            =   10
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin VB.Label lacUnresolvedImports 
         Alignment       =   2  'Center
         Caption         =   "Imported Station Invoices- Unresolved"
         Height          =   180
         Left            =   3120
         TabIndex        =   27
         Top             =   0
         Width           =   3510
      End
      Begin VB.Label lacResolvedImports 
         Alignment       =   2  'Center
         Caption         =   "Imported Station Invoices- Resolved"
         Height          =   180
         Left            =   3780
         TabIndex        =   26
         Top             =   1185
         Width           =   2985
      End
   End
   Begin VB.Label lacScreen 
      Caption         =   "Import Station Spots"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   1965
   End
End
Attribute VB_Name = "ImportStationSpots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ImportStationSpots.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImportStationSpots.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
Dim smImportFiles() As String

'Program library dates Field Areas
Dim imFirstActivate As Integer

Dim hmCHF As Integer    'file handle
Dim imCHFRecLen As Integer  'Record length
Dim tmChf As CHF
Dim tmChfSrchKey0 As LONGKEY0
Dim tmChfSrchKey1 As CHFKEY1
Dim tmChfSrchKey4 As CHFKEY4

Dim hmClf As Integer        'Contract line file handle
Dim tmClfSrchKey0 As CLFKEY0 'CLF key record image
Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
Dim tmClfSrchKey2 As LONGKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image

Dim hmCff As Integer
Dim tmCff As CFF        'CFF record image
Dim tmCffSrchKey0 As CFFKEY0    'CFF key record image
Dim tmCffSrchKey1 As LONGKEY0
Dim imCffRecLen As Integer        'CFF record length

Dim hmSdf As Integer    'Demo Book Name file handle
Dim tmSdf As SDF
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey1 As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey7 As SDFKEY7
Dim imSdfRecLen As Integer        'Sdf record length


Dim hmSmf As Integer    'Demo Book Name file handle
Dim tmSmf As SMF
Dim tmSmfSrchKey2 As LONGKEY0
Dim imSmfRecLen As Integer        'Sdf record length

Dim hmCif As Integer
Dim tmCif As CIF        'Rvf record image
Dim tmCifSrchKey As LONGKEY0
Dim tmCifSrchKey2 As LONGKEY0
Dim imCifRecLen As Integer        'RvF record length

Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim tmCpfSrchKey1 As CPFKEY1 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length

Dim hmSsf As Integer        'Spot summary file handle
Dim lmSsfDate(0 To 6) As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos(0 To 6) As Long  'Record positions
Dim tmSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday; 1 for tuesday;...; 6 for sunday)
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
Dim imSelectedDay As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim imVpfIndex As Integer   'Vehicle option index

Dim hmCrf As Integer
Dim hmGhf As Integer
Dim hmGsf As Integer

Dim tmRdf As RDF

Dim hmSxf As Integer

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imCtrlKey As Integer
Dim imMRLastResultColSorted As Integer
Dim imMRLastResultSort As Integer
Dim imNMRLastResultColSorted As Integer
Dim imNMRLastResultSort As Integer
Dim imMSLastResultColSorted As Integer
Dim imMSLastResultSort As Integer
Dim imCLastResultColSorted As Integer
Dim imCLastResultSort As Integer
Dim lmMRRowSelected As Long
Dim lmNMRRowSelected As Long
Dim lmNRowSelected As Long
Dim lmSRowSelected As Long
Dim lmMRowSelected As Long
Dim lmCRowSelected As Long
Dim imTaxChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer
Dim lmInvStartDate As Long
Dim hmMatch As Integer
Dim hmUnMatch As Integer
Dim lmLastStdMnthBilled As Long
Dim bmInPreview As Boolean
Dim bmFrcDetailAdjusted As Boolean

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer
Dim smGrossNet As String

Dim smInvoiceNumber As String
Dim smContractNumber As String
Dim smAdvertiserName As String
Dim smEstimateNumber As String
Dim smIihfFileName As String
Dim smSourceForm As String  'M=Marketron; W=WideOrbit; R=RadioTraffic, N=NaturalLog, T=Manual Post Date/Time, C=Manual Post Counts, P=Post via Post Log, Blank=Not determined
Dim smCallLetters As String
Dim smNetContractNumber As String

Private Type UNPOSTEDCNTRINFO
    lChfCode As Long
    sSourceForm As String * 2
End Type
Private tmUnpostedCntrInfo() As UNPOSTEDCNTRINFO

Private Type IMPORTSPOTINFO
    lAirDate As Long
    lAirTime As Long
    iLen As Integer
    sISCI As String * 20
    lRate As Long
    lDPStartTime As Long
    lDPEndTime As Long
    sDPDays As String * 7
    bMatched As Boolean
End Type
Private tmImportSpotInfo() As IMPORTSPOTINFO

Private Type STNMATCHINFO
    iLen As Integer
    iAirWeek(0 To 5) As Integer
End Type
Private tmStnMatchInfo() As STNMATCHINFO

Private Type NETSPOTINFO
    sKey As String * 40
    tSdf As SDF
    bMatched As Boolean
    iLen As Integer
    lAcqCost As Long
    lCffIndex As Integer
End Type
Private tmNetSpotInfo() As NETSPOTINFO

Private Type CFFINFO
    lCode As Long
    lMoDate As Long
    lStartDate As Long
    lEndDate As Long
    sDW As String * 1
    iSpotsWk As Integer 'Spots per week, if zero, then daily
    iDay(0 To 6) As Integer
End Type
Private tmCffInfo() As CFFINFO

Private Type MATCHCNTR
    lChfCode As Long
    lSpotCount As Long
End Type
Private tmMatchCntr() As MATCHCNTR
Private Type MATCHCNTRLEN
    lChfCode As Long
    iLen As Integer
    iAirWeek(0 To 5) As Integer
    iSchWeek(0 To 5) As Integer
End Type
Private tmMatchCntrLen() As MATCHCNTRLEN

'******************************************************************************
' amf_Advt_Map Record Definition
'
'******************************************************************************
Private Type AMF
    lCode                 As Long            ' Advertiser Station To Traffic
                                             ' mapping internal reference code
    iVefCode              As Integer         ' Vehicle reference code of the
                                             ' station
    sStationAdvtName      As String * 40     ' Station Advertiser name that does
                                             ' not match one in traffic system
    iAdfCode              As Integer         ' Traffic advertiser substitute
                                             ' reference code
    sUnused               As String * 10     ' Unused
End Type


'Private Type AMFKEY0
'    lCode                 As Long
'End Type

Private Type AMFKEY1
    iVefCode              As Integer
End Type

Private Type AMFKEY2
    iVefCode              As Integer
    sStationAdvtName      As String * 40
End Type

Dim hmAmf As Integer    'Demo Book Name file handle
Dim tmAmf As AMF
Dim tmAmfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmAmfSrchKey2 As AMFKEY2
Dim imAmfRecLen As Integer        'Sdf record length

Dim hmIihf As Integer
Dim tmIihf As IIHF        'CFF record image
Dim tmIihfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmIihfSrchKey1 As IIHFKEY1    'CFF key record image
Dim tmIihfSrchKey2 As IIHFKEY2    'CFF key record image
Dim tmIihfSrchKey3 As IIHFKEY3    'CFF key record image
Dim imIihfRecLen As Integer        'CFF record length

Dim hmIidf As Integer
Dim tmIidf As IIDF        'CFF record image
Dim tmIidfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmIidfSrchKey1 As LONGKEY0    'CFF key record image
Dim tmIidfSrchKey2 As LONGKEY0    'CFF key record image
Dim imIidfRecLen As Integer
Dim tmIidfDetail() As IIDF

Dim hmApf As Integer
Dim tmApf As APF        'CFF record image
Dim tmApfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmApfSrchKey4 As APFKEY4
Dim tmApfSrchKey7 As APFKEY7
Dim imApfRecLen As Integer        'CFF record length

'Result Grids
Const NMRSTATIONINDEX = 0
Const NMRADVERTISERINDEX = 1
Const NMRESTIMATEINDEX = 2
Const NMRCONTRACTINDEX = 3
Const NMRINVOICEINDEX = 4
Const NMRSTATUSINDEX = 5
Const NMRCOUNTINDEX = 6
Const NMRFILENAMEINDEX = 7
Const NMRSORTINDEX = 8
Const NMRIIHFCODEINDEX = 9

Const MRNETADVERTISERINDEX = 0
Const MRNETESTIMATEINDEX = 1
Const MRNETCONTRACTINDEX = 2
Const MRSTNSTATIONINDEX = 3
Const MRSTNADVERTISERINDEX = 4
Const MRSTNESTIMATEINDEX = 5
Const MRSTNCONTRACTINDEX = 6
Const MRSTNINVOICEINDEX = 7
Const MRSTATUSINDEX = 8
Const MRMATCHCOUNTINDEX = 9
Const MRNETCOUNTINDEX = 10
Const MRSTNCOUNTINDEX = 11
Const MRCOMPLIANTINDEX = 12
Const MRFILENAMEINDEX = 13
Const MRSELECTEDINDEX = 14
Const MRCHFCODEINDEX = 15
Const MRIIHFCODEINDEX = 16
Const MRSORTINDEX = 17

'Detail Grids
Const NDPDAYSINDEX = 0
Const NDPTIMEINDEX = 1
Const NLENGTHINDEX = 2
Const NACQRATEINDEX = 3
Const NDATESINDEX = 4
Const NSELECTEDINDEX = 5
Const NSDFCODEINDEX = 6

Const SDATEINDEX = 0
Const STIMEINDEX = 1
Const SLENGTHINDEX = 2
Const SACQRATEINDEX = 3
Const SISCIINDEX = 4
Const SSELECTEDINDEX = 5
Const SIIDFCODEINDEX = 6

Const MNETDPDAYSINDEX = 0
Const MNETDPTIMEINDEX = 1
Const MNETDATESINDEX = 2
Const MSTNDATEINDEX = 3
Const MSTNTIMEINDEX = 4
Const MSTNCOMPLIANTINDEX = 5
Const MLENGTHINDEX = 6
Const MACQRATEINDEX = 7
Const MSTNISCIINDEX = 8
Const MSELECTEDINDEX = 9
Const MSDFCODEINDEX = 10
Const MIIDFCODEINDEX = 11
Const MSSORTINDEX = 12

Const CADVERTISERINDEX = 0
Const CESTIMATEINDEX = 1
Const CCONTRACTINDEX = 2
Const CPRODUCTINDEX = 3
Const CORDERCOUNTINDEX = 4
Const CSTATUSINDEX = 5
Const CPREVIEWINDEX = 6
Const CSELECTEDINDEX = 7
Const CSORTINDEX = 8
Const CCHFCODEINDEX = 9

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub cmcCntrCancel_Click()
    frcCntr.Visible = False
    frcResult.Visible = True
End Sub

Private Sub cmcCntrMatch_Click()
    'Add code to generate match
    mMousePointer grdCntrStation, grdCntrNetwork, vbHourglass
    mProcessUserMatch
    mMousePointer grdCntrStation, grdCntrNetwork, vbDefault
    frcCntr.Visible = False
    frcResult.Visible = True
End Sub

Private Sub cmcDetail_Click()
    If lmMRRowSelected < grdMatchedResult.FixedRows Then
        Exit Sub
    End If
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbHourglass
    imMSLastResultColSorted = -1
    imMSLastResultSort = -1
    mPopDetail
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbDefault
    If Not bmFrcDetailAdjusted Then
        frcDetail.Width = frcDetail.Width + 270
        grdNetworkSpots.Left = 60
        grdStationSpots.Left = grdNetworkSpots.Left + grdNetworkSpots.Width + 120
        grdMatchedSpots.Left = grdNetworkSpots.Left
        bmFrcDetailAdjusted = True
    End If
    frcDetail.Visible = True
    frcResult.Visible = False
End Sub

Private Sub cmcReconcile_Click()
    mReconcile
End Sub

Private Sub cmcReturn_Click()
    Dim ilRet As Integer
    
    If bmInPreview Then
        frcCntr.Visible = True
        mSetDetailButtons
        frcDetail.Visible = False
        bmInPreview = False
    Else
        ilRet = mUpdateApf(lmMRRowSelected)
        frcDetail.Visible = False
        frcResult.Visible = True
    End If
End Sub

Private Sub cmcUndoMatch_Click()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llNoMatchRow As Long
    Dim llIihfCode As Long
    Dim slSchDate As String
    Dim ilGameNo As Integer
    Dim llCntrNo As Long
    
    If lmMRRowSelected < grdMatchedResult.FixedRows Then
        Exit Sub
    End If
    ilRet = MsgBox("Are you sure that you want to Undo the Match", vbYesNo + vbQuestion, "Undo Match")
    If ilRet = vbNo Then
        Exit Sub
    End If
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbHourglass
    llIihfCode = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRIIHFCODEINDEX))
    tmIihfSrchKey0.lCode = llIihfCode
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'Only want "I" left
    tmIihf.lChfCode = 0
    ilRet = btrUpdate(hmIihf, tmIihf, imIihfRecLen)
    tmIidfSrchKey1.lCode = llIihfCode
    ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmIidf.lIihfCode = llIihfCode)
        If tmIidf.sSpotMatchType = "C" Then
            tmSdfSrchKey3.lCode = tmIidf.lSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                imSelectedDay = gWeekDayStr(slSchDate)
                ilGameNo = 0
                ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                tmSdfSrchKey3.lCode = tmIidf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    tmSdf.iDate(0) = tmIidf.iOrigSpotDate(0)
                    tmSdf.iDate(1) = tmIidf.iOrigSpotDate(1)
                    tmSdf.iTime(0) = tmIidf.iOrigSpotTime(0)
                    tmSdf.iTime(1) = tmIidf.iOrigSpotTime(1)
                    tmSdf.sPtType = "0"
                    tmSdf.lCopyCode = 0
                    tmSdf.iRotNo = 0
                    ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                End If
            End If
            tmIidf.sSpotMatchType = "I"
            tmIidf.lSdfCode = 0
            gPackDate "1/1/1970", tmIidf.iOrigSpotDate(0), tmIidf.iOrigSpotDate(1)
            gPackTime "12AM", tmIidf.iOrigSpotTime(0), tmIidf.iOrigSpotTime(1)
            tmIidf.sAgyCompliant = "N"
            ilRet = btrUpdate(hmIidf, tmIidf, imIidfRecLen)
            'ReDim tmNetSpotInfo(0 To 1) As NETSPOTINFO
            'tmNetSpotInfo(0).tSdf = tmSdf
            'ilRet = mAddIidf("M", 0, 0, "N")
            tmIidfSrchKey1.lCode = llIihfCode
            ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        ElseIf tmIidf.sSpotMatchType = "M" Then
            ilRet = btrDelete(hmIidf)
            tmIidfSrchKey1.lCode = llIihfCode
            ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Else
            ilRet = btrGetNext(hmIidf, tmIidf, imIidfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
    Loop
    'Remove from grid and add to grid
    llNoMatchRow = grdNoMatchedResult.FixedRows
    For llRow = grdNoMatchedResult.FixedRows To grdNoMatchedResult.Rows - 1 Step 1
        If grdNoMatchedResult.TextMatrix(llRow, NMRSTATIONINDEX) <> "" Then
            llNoMatchRow = llRow + 1
        End If
    Next llRow
    
    If llNoMatchRow >= grdNoMatchedResult.Rows Then
        grdNoMatchedResult.AddItem ""
    End If
    grdNoMatchedResult.RowHeight(llNoMatchRow) = fgBoxGridH + 15
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRSTATIONINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNSTATIONINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRADVERTISERINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNADVERTISERINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRESTIMATEINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNESTIMATEINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCONTRACTINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCONTRACTINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRINVOICEINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNINVOICEINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRFILENAMEINDEX) = grdMatchedResult.TextMatrix(lmMRRowSelected, MRFILENAMEINDEX)
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRMATCHCOUNTINDEX)) + Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCOUNTINDEX))
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRSTATUSINDEX) = "User pressed Undo Match"
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRIIHFCODEINDEX) = tmIihf.lCode
    
    llCntrNo = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETCONTRACTINDEX))
    mClearApfAirCount llCntrNo, llIihfCode
    grdMatchedResult.RemoveItem lmMRRowSelected
    lmMRRowSelected = -1
        
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbDefault
    mSetCommands
End Sub

Private Sub cmcUndoReconcile_Click()
    mUndoReconcile
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    ImportStationSpots.Refresh
    Me.KeyPreview = True
    tmcStart.Enabled = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    mResizeForm
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Terminate()
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmStnMatchInfo
    Erase tmImportSpotInfo
    Erase tmNetSpotInfo
    Erase tmCffInfo
    Erase tmMatchCntr
    Erase tmMatchCntrLen
    Erase tmIidfDetail
    Erase tmUnpostedCntrInfo
    
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    
    ilRet = btrClose(hmApf)
    btrDestroy hmApf
    ilRet = btrClose(hmIidf)
    btrDestroy hmIidf
    ilRet = btrClose(hmIihf)
    btrDestroy hmIihf
    ilRet = btrClose(hmAmf)
    btrDestroy hmAmf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Set ImportStationSpots = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
End Sub


Private Sub grdCntrNetwork_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdCntrNetwork.ToolTipText = ""
    'If grdCntrNetwork.MouseRow < grdCntrNetwork.FixedRows Then
    '    Exit Sub
    'End If
    If grdCntrNetwork.MouseCol <= CORDERCOUNTINDEX Then
        grdCntrNetwork.ToolTipText = Trim$(grdCntrNetwork.TextMatrix(grdCntrNetwork.MouseRow, grdCntrNetwork.MouseCol))
    End If
End Sub

Private Sub grdCntrNetwork_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdCntrNetwork.RowHeight(0) Then
        llCol = grdCntrNetwork.MouseCol
        If lmCRowSelected >= grdCntrNetwork.FixedRows Then
            grdCntrNetwork.TextMatrix(lmCRowSelected, CSELECTEDINDEX) = "0"
            mCPaintRowColor lmCRowSelected
        End If
        If (llCol <> CORDERCOUNTINDEX) And (llCol <> CPREVIEWINDEX) Then
            grdCntrNetwork.Col = llCol
            mCSortCol grdCntrNetwork.MouseCol
        End If
        grdCntrNetwork.Row = 0
        grdCntrNetwork.Col = CCHFCODEINDEX
        grdCntrNetwork.Redraw = True
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdCntrNetwork, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdCntrNetwork.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdCntrNetwork.FixedRows Then
        If grdCntrNetwork.TextMatrix(llCurrentRow, CADVERTISERINDEX) <> "" Then
            llTopRow = grdCntrNetwork.TopRow
            If lmCRowSelected >= grdCntrNetwork.FixedRows Then
                grdCntrNetwork.TextMatrix(lmCRowSelected, CSELECTEDINDEX) = "0"
                mCPaintRowColor lmCRowSelected
            End If
            If lmCRowSelected <> llCurrentRow Then
                grdCntrNetwork.TextMatrix(llCurrentRow, CSELECTEDINDEX) = "1"
                mCPaintRowColor llCurrentRow
                lmCRowSelected = llCurrentRow
                grdCntrNetwork.TopRow = llTopRow
                grdCntrNetwork.Row = llCurrentRow
            Else
                If grdCntrNetwork.MouseCol <> CPREVIEWINDEX Then
                    lmCRowSelected = -1
                End If
                grdCntrNetwork.TopRow = llTopRow
                grdCntrNetwork.Row = llCurrentRow
            End If
            If grdCntrNetwork.MouseCol = CPREVIEWINDEX Then
                bmInPreview = True
                grdCntrNetwork.TextMatrix(llCurrentRow, CSELECTEDINDEX) = "1"
                mCPaintRowColor llCurrentRow
                lmCRowSelected = llCurrentRow
                mMousePointer grdMatchedResult, grdNoMatchedResult, vbHourglass
                mMousePointer grdCntrStation, grdCntrNetwork, vbHourglass
                ReDim tmIidfDetail(0 To 0) As IIDF
                mProcessUserMatch True
                imMSLastResultColSorted = -1
                imMSLastResultSort = -1
                mPopDetail True
                mMousePointer grdCntrStation, grdCntrNetwork, vbDefault
                mMousePointer grdMatchedResult, grdNoMatchedResult, vbDefault
                If Not bmFrcDetailAdjusted Then
                    frcDetail.Width = frcDetail.Width + 270
                    grdNetworkSpots.Left = 60
                    grdStationSpots.Left = grdNetworkSpots.Left + grdNetworkSpots.Width + 120
                    grdMatchedSpots.Left = grdNetworkSpots.Left
                    bmFrcDetailAdjusted = True
                End If
                mSetDetailButtons True
                frcDetail.Visible = True
                frcCntr.Visible = False
                frcDetail.ZOrder vbBringToFront
            End If
        End If
    End If
    grdCntrNetwork.Redraw = True
    mCntrSetCommands
End Sub

Private Sub grdCntrStation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdCntrStation.ToolTipText = ""
    'If grdCntrStation.MouseRow < grdCntrStation.FixedRows Then
    '    Exit Sub
    'End If
    If grdCntrStation.MouseCol = NMRFILENAMEINDEX Then
        grdCntrStation.ToolTipText = Trim$(grdCntrStation.TextMatrix(grdCntrStation.MouseRow, grdCntrStation.MouseCol))
    End If
End Sub

Private Sub grdMatchedResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdMatchedResult.ToolTipText = ""
    If (grdMatchedResult.MouseRow < grdMatchedResult.FixedRows) Or (grdMatchedResult.MouseCol > MRFILENAMEINDEX) Then
        Exit Sub
    End If
    grdMatchedResult.ToolTipText = Trim$(grdMatchedResult.TextMatrix(grdMatchedResult.MouseRow, grdMatchedResult.MouseCol))

End Sub

Private Sub grdMatchedResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdMatchedResult.RowHeight(0) Then
        llCol = grdMatchedResult.MouseCol
        If lmMRRowSelected >= grdMatchedResult.FixedRows Then
            grdMatchedResult.TextMatrix(lmMRRowSelected, MRSELECTEDINDEX) = "0"
            mPaintRowColor lmMRRowSelected
        End If
        grdMatchedResult.Col = llCol
        mMRSortCol grdMatchedResult.Col
        grdMatchedResult.Row = 0
        grdMatchedResult.Col = MRIIHFCODEINDEX
        grdMatchedResult.Redraw = True
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdMatchedResult, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdMatchedResult.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdMatchedResult.FixedRows Then
        If grdMatchedResult.TextMatrix(llCurrentRow, MRSTNSTATIONINDEX) <> "" Then
            llTopRow = grdMatchedResult.TopRow
            If lmMRRowSelected >= grdMatchedResult.FixedRows Then
                grdMatchedResult.TextMatrix(lmMRRowSelected, MRSELECTEDINDEX) = "0"
                mPaintRowColor lmMRRowSelected
            End If
            If lmMRRowSelected <> llCurrentRow Then
                grdMatchedResult.TextMatrix(llCurrentRow, MRSELECTEDINDEX) = "1"
                mPaintRowColor llCurrentRow
                lmMRRowSelected = llCurrentRow
                grdMatchedResult.TopRow = llTopRow
                grdMatchedResult.Row = llCurrentRow
            Else
                lmMRRowSelected = -1
                grdMatchedResult.TopRow = llTopRow
                grdMatchedResult.Row = llCurrentRow
            End If
        End If
    End If
    grdMatchedResult.Redraw = True
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False

    Screen.MousePointer = vbHourglass
    lmMRRowSelected = -1
    lmNMRRowSelected = -1
    lmCRowSelected = -1
    lmNRowSelected = -1
    lmMRowSelected = -1
    lmSRowSelected = -1
    imMRLastResultColSorted = -1
    imMRLastResultSort = -1
    imNMRLastResultColSorted = -1
    imNMRLastResultSort = -1
    imMSLastResultColSorted = -1
    imMSLastResultSort = -1
    imCLastResultColSorted = -1
    imCLastResultSort = -1
    
    smImportFiles = Split(sgBrowserFile, "|")

    imTaxChg = False
    bmInPreview = False
    bmFrcDetailAdjusted = False

    imFirstFocus = True
    imCtrlKey = False
    
    imCHFRecLen = Len(tmChf)
    hmCHF = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", ImportStationSpots
    On Error GoTo 0

    imClfRecLen = Len(tmClf)
    hmClf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", ImportStationSpots
    On Error GoTo 0

    imCffRecLen = Len(tmCff)
    hmCff = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", ImportStationSpots
    On Error GoTo 0

    imSdfRecLen = Len(tmSdf)
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", ImportStationSpots
    On Error GoTo 0

    imCifRecLen = Len(tmCif)
    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", ImportStationSpots
    On Error GoTo 0

    imCpfRecLen = Len(tmCpf)
    hmCpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", ImportStationSpots
    On Error GoTo 0

    imSsfRecLen = Len(tmSsf(0))
    hmSsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", ImportStationSpots
    On Error GoTo 0

    imAmfRecLen = Len(tmAmf)
    hmAmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmAmf, "", sgDBPath & "Amf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Amf.Btr)", ImportStationSpots
    On Error GoTo 0

    imIidfRecLen = Len(tmIidf)
    hmIidf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmIidf, "", sgDBPath & "Iidf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Iidf.Btr)", ImportStationSpots
    On Error GoTo 0

    imIihfRecLen = Len(tmIihf)
    hmIihf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Iihf.Btr)", ImportStationSpots
    On Error GoTo 0

    imApfRecLen = Len(tmApf)
    hmApf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmApf, "", sgDBPath & "Apf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Apf.Btr)", ImportStationSpots
    On Error GoTo 0

    imSmfRecLen = Len(tmSmf)
    hmSmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", ImportStationSpots
    On Error GoTo 0

    hmCrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", ImportStationSpots
    On Error GoTo 0

    hmGsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", ImportStationSpots
    On Error GoTo 0

    hmGhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", ImportStationSpots
    On Error GoTo 0

    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", ImportStationSpots
    On Error GoTo 0

    mInitBox

    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), lmLastStdMnthBilled
    
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())

    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Unload ImportStationSpots
End Sub






Private Sub grdMatchedSpots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdMatchedSpots.ToolTipText = ""
    'If grdMatchedSpots.MouseRow < grdMatchedSpots.FixedRows Then
    '    Exit Sub
    'End If
    If grdMatchedSpots.MouseCol <= MSTNISCIINDEX Then
        grdMatchedSpots.ToolTipText = Trim$(grdMatchedSpots.TextMatrix(grdMatchedSpots.MouseRow, grdMatchedSpots.MouseCol))
    End If
End Sub

Private Sub grdMatchedSpots_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdMatchedSpots.RowHeight(0) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdMatchedSpots, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdMatchedSpots.FixedRows Then
        Exit Sub
    End If
    grdMatchedSpots.Redraw = False
    If llCurrentRow >= grdMatchedSpots.FixedRows Then
        If grdMatchedSpots.TextMatrix(llCurrentRow, MNETDPDAYSINDEX) <> "" Then
            llTopRow = grdMatchedSpots.TopRow
            If lmMRowSelected >= grdMatchedSpots.FixedRows Then
                grdMatchedSpots.TextMatrix(lmMRowSelected, MSELECTEDINDEX) = "0"
                mMPaintRowColor lmMRowSelected
            End If
            If lmMRowSelected <> llCurrentRow Then
                grdMatchedSpots.TextMatrix(llCurrentRow, MSELECTEDINDEX) = "1"
                mMPaintRowColor llCurrentRow
                lmMRowSelected = llCurrentRow
                grdMatchedSpots.TopRow = llTopRow
                grdMatchedSpots.Row = llCurrentRow
                If lmNRowSelected >= grdNetworkSpots.FixedRows Then
                    grdNetworkSpots.TextMatrix(lmNRowSelected, NSELECTEDINDEX) = "0"
                    mNPaintRowColor lmNRowSelected
                    lmNRowSelected = -1
                End If
                If lmSRowSelected >= grdStationSpots.FixedRows Then
                    grdStationSpots.TextMatrix(lmSRowSelected, SSELECTEDINDEX) = "0"
                    mSPaintRowColor lmSRowSelected
                    lmSRowSelected = -1
                End If
            Else
                lmMRowSelected = -1
                grdMatchedSpots.TopRow = llTopRow
                grdMatchedSpots.Row = llCurrentRow
            End If
        End If
    End If
    grdMatchedSpots.Redraw = True
    mDetailSetCommands
End Sub

Private Sub grdNetworkSpots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdNetworkSpots.ToolTipText = ""
    'If grdNetworkSpots.MouseRow < grdNetworkSpots.FixedRows Then
    '    Exit Sub
    'End If
    If grdNetworkSpots.MouseCol <= NDATESINDEX Then
        grdNetworkSpots.ToolTipText = Trim$(grdNetworkSpots.TextMatrix(grdNetworkSpots.MouseRow, grdNetworkSpots.MouseCol))
    End If
End Sub

Private Sub grdNetworkSpots_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdNetworkSpots.RowHeight(0) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdNetworkSpots, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdNetworkSpots.FixedRows Then
        Exit Sub
    End If
    grdNetworkSpots.Redraw = False
    If llCurrentRow >= grdNetworkSpots.FixedRows Then
        If grdNetworkSpots.TextMatrix(llCurrentRow, NDPDAYSINDEX) <> "" Then
            llTopRow = grdNetworkSpots.TopRow
            If lmNRowSelected >= grdNetworkSpots.FixedRows Then
                grdNetworkSpots.TextMatrix(lmNRowSelected, NSELECTEDINDEX) = "0"
                mNPaintRowColor lmNRowSelected
            End If
            If lmNRowSelected <> llCurrentRow Then
                grdNetworkSpots.TextMatrix(llCurrentRow, NSELECTEDINDEX) = "1"
                mNPaintRowColor llCurrentRow
                lmNRowSelected = llCurrentRow
                grdNetworkSpots.TopRow = llTopRow
                grdNetworkSpots.Row = llCurrentRow
                If lmMRowSelected >= grdMatchedSpots.FixedRows Then
                    grdMatchedSpots.TextMatrix(lmMRowSelected, MSELECTEDINDEX) = "0"
                    mMPaintRowColor lmMRowSelected
                    lmMRowSelected = -1
                End If
            Else
                lmNRowSelected = -1
                grdNetworkSpots.TopRow = llTopRow
                grdNetworkSpots.Row = llCurrentRow
            End If
        End If
    End If
    grdNetworkSpots.Redraw = True
    If (lmNRowSelected >= grdNetworkSpots.FixedRows) And (lmSRowSelected >= grdStationSpots.FixedRows) Then
        If Val(grdNetworkSpots.TextMatrix(lmNRowSelected, NLENGTHINDEX)) <> Val(grdStationSpots.TextMatrix(lmSRowSelected, SLENGTHINDEX)) Then
            grdStationSpots.TextMatrix(lmSRowSelected, SSELECTEDINDEX) = "0"
            mSPaintRowColor lmSRowSelected
            lmSRowSelected = -1
        End If
    End If
    mDetailSetCommands
End Sub

Private Sub grdNoMatchedResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdNoMatchedResult.ToolTipText = ""
    'If grdNoMatchedResult.MouseRow < grdNoMatchedResult.FixedRows Then
    '    Exit Sub
    'End If
    If grdNoMatchedResult.MouseCol <= NMRFILENAMEINDEX Then
        grdNoMatchedResult.ToolTipText = Trim$(grdNoMatchedResult.TextMatrix(grdNoMatchedResult.MouseRow, grdNoMatchedResult.MouseCol))
    End If
End Sub

Private Sub grdNoMatchedResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdNoMatchedResult.RowHeight(0) Then
        grdNoMatchedResult.Col = grdNoMatchedResult.MouseCol
        If grdNoMatchedResult.Col <> NMRCOUNTINDEX Then
            mNMRSortCol grdNoMatchedResult.Col
        End If
        grdNoMatchedResult.Row = 0
        grdNoMatchedResult.Col = NMRIIHFCODEINDEX
        grdNoMatchedResult.Redraw = True
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdNoMatchedResult, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdNoMatchedResult.FixedRows Then
        Exit Sub
    End If
    grdNoMatchedResult.Redraw = False
    If llCurrentRow >= grdNoMatchedResult.FixedRows Then
        If grdNoMatchedResult.TextMatrix(llCurrentRow, NMRSTATIONINDEX) <> "" Then
            llTopRow = grdNoMatchedResult.TopRow
            lmNMRRowSelected = llCurrentRow
            mPopCntr
        End If
    End If
    grdNoMatchedResult.Redraw = True
End Sub


Private Sub grdStationSpots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdStationSpots.ToolTipText = ""
    If grdStationSpots.MouseRow < grdStationSpots.FixedRows Then
        Exit Sub
    End If
    If grdStationSpots.MouseCol <= SISCIINDEX Then
        grdStationSpots.ToolTipText = Trim$(grdStationSpots.TextMatrix(grdStationSpots.MouseRow, grdStationSpots.MouseCol))
    End If
End Sub

Private Sub grdStationSpots_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdStationSpots.RowHeight(0) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdStationSpots, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdStationSpots.FixedRows Then
        Exit Sub
    End If
    grdStationSpots.Redraw = False
    If llCurrentRow >= grdStationSpots.FixedRows Then
        If grdStationSpots.TextMatrix(llCurrentRow, SDATEINDEX) <> "" Then
            llTopRow = grdStationSpots.TopRow
            If lmSRowSelected >= grdStationSpots.FixedRows Then
                grdStationSpots.TextMatrix(lmSRowSelected, SSELECTEDINDEX) = "0"
                mSPaintRowColor lmSRowSelected
            End If
            If lmSRowSelected <> llCurrentRow Then
                grdStationSpots.TextMatrix(llCurrentRow, SSELECTEDINDEX) = "1"
                mSPaintRowColor llCurrentRow
                lmSRowSelected = llCurrentRow
                grdStationSpots.TopRow = llTopRow
                grdStationSpots.Row = llCurrentRow
                If lmMRowSelected >= grdMatchedSpots.FixedRows Then
                    grdMatchedSpots.TextMatrix(lmMRowSelected, MSELECTEDINDEX) = "0"
                    mMPaintRowColor lmMRowSelected
                    lmMRowSelected = -1
                End If
            Else
                lmSRowSelected = llCurrentRow
                grdStationSpots.TopRow = llTopRow
                grdStationSpots.Row = -1
            End If
        End If
    End If
    grdStationSpots.Redraw = True
    If (lmNRowSelected >= grdNetworkSpots.FixedRows) And (lmSRowSelected >= grdStationSpots.FixedRows) Then
        If Val(grdNetworkSpots.TextMatrix(lmNRowSelected, NLENGTHINDEX)) <> Val(grdStationSpots.TextMatrix(lmSRowSelected, SLENGTHINDEX)) Then
            grdNetworkSpots.TextMatrix(lmNRowSelected, NSELECTEDINDEX) = "0"
            mNPaintRowColor lmNRowSelected
            lmNRowSelected = -1
        End If
    End If
    mDetailSetCommands
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdMatchedResult.Visible Then
        'lmMRRowSelected = -1
        'grdMatchedResult.Row = 0
        'grdMatchedResult.Col = MRIIHFCODEINDEX
        mSetCommands
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub mPopulate()

    Dim ilRet As Integer

    mMoveRecToCtrl
End Sub


Private Sub mSetCommands()
    Dim ilRet As Integer
    If lmMRRowSelected >= grdMatchedResult.FixedRows Then
        cmcDetail.Enabled = True
        cmcUndoMatch.Enabled = True
    Else
        cmcDetail.Enabled = False
        cmcUndoMatch.Enabled = False
    End If
    Exit Sub
End Sub

Private Sub mDetailSetCommands()
    Dim ilRet As Integer
    If frcCntr.Visible Then
        Exit Sub
    End If
    If lmMRowSelected >= grdMatchedSpots.FixedRows Then
        cmcUndoReconcile.Enabled = True
    Else
        cmcUndoReconcile.Enabled = False
    End If
    If (lmNRowSelected >= grdNetworkSpots.FixedRows) And (lmSRowSelected >= grdStationSpots.FixedRows) Then
        cmcReconcile.Enabled = True
    Else
        cmcReconcile.Enabled = False
    End If
    Exit Sub
End Sub
Private Sub mCntrSetCommands()
    Dim ilRet As Integer
    If lmCRowSelected >= grdCntrNetwork.FixedRows Then
        cmcCntrMatch.Enabled = True
    Else
        cmcCntrMatch.Enabled = False
    End If
    Exit Sub
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilLoop                        ilRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35

    'frcResult.Move 60, 300, Me.Width - 240, Me.Height - 300 - 120
    'grdMatchedResult.Move 0, 0, frcResult.Width, frcResult.Height - 2 * cmcDetail.Height - 90
    'cmcDetail.Move frcResult.Width / 2 - 2 * cmcDetail.Width, frcResult.Height - cmcDetail.Height
    'cmcCancel.Move frcResult.Width / 2 + cmcCancel.Width, cmcDetail.Top

    'mGridResultLayout
    'mGridResultColumnWidths
    'mGridResultColumns
End Sub


Private Sub mGridResultColumns()

    grdNoMatchedResult.Row = grdNoMatchedResult.FixedRows - 1
    grdNoMatchedResult.Col = NMRSTATIONINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Call Letters"
    grdNoMatchedResult.Col = NMRADVERTISERINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Advertiser"
    grdNoMatchedResult.Col = NMRESTIMATEINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Estimate #"
    grdNoMatchedResult.Col = NMRCONTRACTINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Contract #"
    grdNoMatchedResult.Col = NMRINVOICEINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Invoice #"
    grdNoMatchedResult.Col = NMRSTATUSINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Status"
    grdNoMatchedResult.Col = NMRCOUNTINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    'grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "Count"
    grdNoMatchedResult.Col = NMRFILENAMEINDEX
    grdNoMatchedResult.CellFontBold = False
    grdNoMatchedResult.CellFontName = "Arial"
    grdNoMatchedResult.CellFontSize = 6.75
    grdNoMatchedResult.CellForeColor = vbBlue
    grdNoMatchedResult.CellBackColor = LIGHTBLUE
    grdNoMatchedResult.TextMatrix(grdNoMatchedResult.Row, grdNoMatchedResult.Col) = "File Name"


    grdMatchedResult.Row = grdMatchedResult.FixedRows - 2
    grdMatchedResult.Col = MRNETADVERTISERINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Network"
    grdMatchedResult.Col = MRNETESTIMATEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Network"
    grdMatchedResult.Col = MRNETCONTRACTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Network"
    grdMatchedResult.Col = MRSTNSTATIONINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRSTNADVERTISERINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRSTNESTIMATEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRSTNCONTRACTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRSTNINVOICEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRSTATUSINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Status"
    grdMatchedResult.Col = MRMATCHCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Match"
    grdMatchedResult.Col = MRNETCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Unmatch"
    grdMatchedResult.Col = MRSTNCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Unmatch"
    grdMatchedResult.Col = MRCOMPLIANTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Compliant"
    grdMatchedResult.Col = MRFILENAMEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "File"

    grdMatchedResult.Row = grdMatchedResult.FixedRows - 1
    grdMatchedResult.Col = MRNETADVERTISERINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Advertiser"
    grdMatchedResult.Col = MRNETESTIMATEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Estimate #"
    grdMatchedResult.Col = MRNETCONTRACTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Contract #"
    grdMatchedResult.Col = MRSTNSTATIONINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Call Letters"
    grdMatchedResult.Col = MRSTNADVERTISERINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Advertiser"
    grdMatchedResult.Col = MRSTNESTIMATEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Estimate #"
    grdMatchedResult.Col = MRSTNCONTRACTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Contract #"
    grdMatchedResult.Col = MRSTNINVOICEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Invoice #"
    grdMatchedResult.Col = MRSTATUSINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = ""
    grdMatchedResult.Col = MRMATCHCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Count"
    grdMatchedResult.Col = MRNETCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Network"
    grdMatchedResult.Col = MRSTNCOUNTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Station"
    grdMatchedResult.Col = MRCOMPLIANTINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = ""
    grdMatchedResult.Col = MRFILENAMEINDEX
    grdMatchedResult.CellFontBold = False
    grdMatchedResult.CellFontName = "Arial"
    grdMatchedResult.CellFontSize = 6.75
    grdMatchedResult.CellForeColor = vbBlue
    grdMatchedResult.CellBackColor = LIGHTBLUE
    grdMatchedResult.TextMatrix(grdMatchedResult.Row, grdMatchedResult.Col) = "Name"

End Sub

Private Sub mGridResultColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    
    
    grdNoMatchedResult.ColWidth(NMRIIHFCODEINDEX) = 0
    grdNoMatchedResult.ColWidth(NMRSORTINDEX) = 0
    grdNoMatchedResult.ColWidth(NMRSTATIONINDEX) = 0.05 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRADVERTISERINDEX) = 0.07 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRESTIMATEINDEX) = 0.05 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRCONTRACTINDEX) = 0.05 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRINVOICEINDEX) = 0.05 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRSTATUSINDEX) = 0.08 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRCOUNTINDEX) = 0.05 * grdNoMatchedResult.Width
    grdNoMatchedResult.ColWidth(NMRFILENAMEINDEX) = 0.14 * grdNoMatchedResult.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdNoMatchedResult.Width
    For ilCol = 0 To grdNoMatchedResult.Cols - 1 Step 1
        llWidth = llWidth + grdNoMatchedResult.ColWidth(ilCol)
        If (grdNoMatchedResult.ColWidth(ilCol) > 15) And (grdNoMatchedResult.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdNoMatchedResult.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdNoMatchedResult.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdNoMatchedResult.Width
            For ilCol = 0 To grdNoMatchedResult.Cols - 1 Step 1
                If (grdNoMatchedResult.ColWidth(ilCol) > 15) And (grdNoMatchedResult.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdNoMatchedResult.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdNoMatchedResult.FixedCols To grdNoMatchedResult.Cols - 1 Step 1
                If grdNoMatchedResult.ColWidth(ilCol) > 15 Then
                    ilColInc = grdNoMatchedResult.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdNoMatchedResult.ColWidth(ilCol) = grdNoMatchedResult.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If



    grdMatchedResult.ColWidth(MRIIHFCODEINDEX) = 0
    grdMatchedResult.ColWidth(MRCHFCODEINDEX) = 0
    grdMatchedResult.ColWidth(MRSORTINDEX) = 0
    grdMatchedResult.ColWidth(MRSELECTEDINDEX) = 0
    grdMatchedResult.ColWidth(MRNETADVERTISERINDEX) = 0.08 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRNETESTIMATEINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRNETCONTRACTINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNSTATIONINDEX) = 0.07 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNADVERTISERINDEX) = 0.08 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNESTIMATEINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNCONTRACTINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNINVOICEINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTATUSINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRMATCHCOUNTINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRNETCOUNTINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRSTNCOUNTINDEX) = 0.05 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRCOMPLIANTINDEX) = 0.04 * grdMatchedResult.Width
    grdMatchedResult.ColWidth(MRFILENAMEINDEX) = 0.14 * grdMatchedResult.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdMatchedResult.Width
    For ilCol = 0 To grdMatchedResult.Cols - 1 Step 1
        llWidth = llWidth + grdMatchedResult.ColWidth(ilCol)
        If (grdMatchedResult.ColWidth(ilCol) > 15) And (grdMatchedResult.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdMatchedResult.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdMatchedResult.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdMatchedResult.Width
            For ilCol = 0 To grdMatchedResult.Cols - 1 Step 1
                If (grdMatchedResult.ColWidth(ilCol) > 15) And (grdMatchedResult.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdMatchedResult.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdMatchedResult.FixedCols To grdMatchedResult.Cols - 1 Step 1
                If grdMatchedResult.ColWidth(ilCol) > 15 Then
                    ilColInc = grdMatchedResult.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdMatchedResult.ColWidth(ilCol) = grdMatchedResult.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub mGridCntrColumns()

    grdCntrStation.Row = grdCntrStation.FixedRows - 1
    grdCntrStation.Col = NMRSTATIONINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Call Letters"
    grdCntrStation.Col = NMRADVERTISERINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Advertiser"
    grdCntrStation.Col = NMRESTIMATEINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Estimate #"
    grdCntrStation.Col = NMRCONTRACTINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Contract #"
    grdCntrStation.Col = NMRINVOICEINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Invoice #"
    grdCntrStation.Col = NMRSTATUSINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Status"
    grdCntrStation.Col = NMRCOUNTINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "Count"
    grdCntrStation.Col = NMRFILENAMEINDEX
    grdCntrStation.CellFontBold = False
    grdCntrStation.CellFontName = "Arial"
    grdCntrStation.CellFontSize = 6.75
    grdCntrStation.CellForeColor = vbBlue
    'grdCntrStation.CellBackColor = LIGHTBLUE
    grdCntrStation.TextMatrix(grdCntrStation.Row, grdCntrStation.Col) = "File Name"


    grdCntrNetwork.Row = grdCntrNetwork.FixedRows - 1
    grdCntrNetwork.Col = CADVERTISERINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Advertiser"
    grdCntrNetwork.Col = CESTIMATEINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Estimate #"
    grdCntrNetwork.Col = CCONTRACTINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Contract #"
    grdCntrNetwork.Col = CPRODUCTINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Product"
    grdCntrNetwork.Col = CORDERCOUNTINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    'grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "# Ordered"
    grdCntrNetwork.Col = CSTATUSINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Status"
    grdCntrNetwork.Col = CPREVIEWINDEX
    grdCntrNetwork.CellFontBold = False
    grdCntrNetwork.CellFontName = "Arial"
    grdCntrNetwork.CellFontSize = 6.75
    grdCntrNetwork.CellForeColor = vbBlue
    'grdCntrNetwork.CellBackColor = LIGHTBLUE
    grdCntrNetwork.TextMatrix(grdCntrNetwork.Row, grdCntrNetwork.Col) = "Preview"

End Sub

Private Sub mGridCntrColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    
    'Copy of grdNoMatchResult
    grdCntrStation.ColWidth(NMRSORTINDEX) = 0
    grdCntrStation.ColWidth(NMRIIHFCODEINDEX) = 0
    grdCntrStation.ColWidth(NMRSTATIONINDEX) = 0.05 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRADVERTISERINDEX) = 0.07 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRESTIMATEINDEX) = 0.05 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRCONTRACTINDEX) = 0.05 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRINVOICEINDEX) = 0.05 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRSTATUSINDEX) = 0.08 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRCOUNTINDEX) = 0.05 * grdCntrStation.Width
    grdCntrStation.ColWidth(NMRFILENAMEINDEX) = 0.14 * grdCntrStation.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = 0  'grdCntrStation.Width
    For ilCol = 0 To grdCntrStation.Cols - 1 Step 1
        llWidth = llWidth + grdCntrStation.ColWidth(ilCol)
        If (grdCntrStation.ColWidth(ilCol) > 15) And (grdCntrStation.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCntrStation.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCntrStation.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCntrStation.Width
            For ilCol = 0 To grdCntrStation.Cols - 1 Step 1
                If (grdCntrStation.ColWidth(ilCol) > 15) And (grdCntrStation.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCntrStation.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCntrStation.FixedCols To grdCntrStation.Cols - 1 Step 1
                If grdCntrStation.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCntrStation.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdCntrStation.ColWidth(ilCol) = grdCntrStation.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If



    grdCntrNetwork.ColWidth(CCHFCODEINDEX) = 0
    grdCntrNetwork.ColWidth(CSORTINDEX) = 0
    grdCntrNetwork.ColWidth(CSELECTEDINDEX) = 0
    grdCntrNetwork.ColWidth(CADVERTISERINDEX) = 0.15 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CESTIMATEINDEX) = 0.05 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CCONTRACTINDEX) = 0.05 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CPRODUCTINDEX) = 0.15 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CORDERCOUNTINDEX) = 0.05 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CSTATUSINDEX) = 0.15 * grdCntrNetwork.Width
    grdCntrNetwork.ColWidth(CPREVIEWINDEX) = 0.05 * grdCntrNetwork.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdCntrNetwork.Width
    For ilCol = 0 To grdCntrNetwork.Cols - 1 Step 1
        llWidth = llWidth + grdCntrNetwork.ColWidth(ilCol)
        If (grdCntrNetwork.ColWidth(ilCol) > 15) And (grdCntrNetwork.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCntrNetwork.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCntrNetwork.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCntrNetwork.Width
            For ilCol = 0 To grdCntrNetwork.Cols - 1 Step 1
                If (grdCntrNetwork.ColWidth(ilCol) > 15) And (grdCntrNetwork.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCntrNetwork.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCntrNetwork.FixedCols To grdCntrNetwork.Cols - 1 Step 1
                If grdCntrNetwork.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCntrNetwork.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdCntrNetwork.ColWidth(ilCol) = grdCntrNetwork.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub
Private Sub mMRSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdMatchedResult.FixedRows To grdMatchedResult.Rows - 1 Step 1
        slStr = Trim$(grdMatchedResult.TextMatrix(llRow, MRSTNSTATIONINDEX))
        If slStr <> "" Then
            If (ilCol = MRNETCONTRACTINDEX) Then
                slSort = Val(grdMatchedResult.TextMatrix(llRow, MRNETCONTRACTINDEX))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MRMATCHCOUNTINDEX) Then
                slSort = Val(grdMatchedResult.TextMatrix(llRow, MRMATCHCOUNTINDEX))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MRNETCOUNTINDEX) Then
                'sort in reverse order
                slSort = 99999 - Val(grdMatchedResult.TextMatrix(llRow, MRNETCOUNTINDEX))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MRSTNCOUNTINDEX) Then
                slSort = Val(grdMatchedResult.TextMatrix(llRow, MRSTNCOUNTINDEX))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MRCOMPLIANTINDEX) Then
                If grdMatchedResult.TextMatrix(llRow, MRCOMPLIANTINDEX) = "No" Then
                    slSort = "A"
                Else
                    slSort = "B"
                End If
            Else
                slSort = grdMatchedResult.TextMatrix(llRow, ilCol)
            End If
            slStr = grdMatchedResult.TextMatrix(llRow, MRSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imMRLastResultColSorted) Or ((ilCol = imMRLastResultColSorted) And (imMRLastResultSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMatchedResult.TextMatrix(llRow, MRSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMatchedResult.TextMatrix(llRow, MRSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imMRLastResultColSorted Then
        imMRLastResultColSorted = MRSORTINDEX
    Else
        imMRLastResultColSorted = -1
        imMRLastResultSort = -1
    End If
    gGrid_SortByCol grdMatchedResult, MRSTNSTATIONINDEX, MRSORTINDEX, imMRLastResultColSorted, imMRLastResultSort
    imMRLastResultColSorted = ilCol
End Sub
Private Sub mNMRSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdNoMatchedResult.FixedRows To grdNoMatchedResult.Rows - 1 Step 1
        slStr = Trim$(grdNoMatchedResult.TextMatrix(llRow, NMRFILENAMEINDEX))
        If slStr <> "" Then
            If (ilCol = NMRCOUNTINDEX) Then
                'sort in reverse order
                slSort = 99999 - Val(grdNoMatchedResult.TextMatrix(llRow, NMRCOUNTINDEX))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            Else
                slSort = grdNoMatchedResult.TextMatrix(llRow, ilCol)
            End If
            slStr = grdNoMatchedResult.TextMatrix(llRow, NMRSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imNMRLastResultColSorted) Or ((ilCol = imNMRLastResultColSorted) And (imNMRLastResultSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdNoMatchedResult.TextMatrix(llRow, NMRSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdNoMatchedResult.TextMatrix(llRow, NMRSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imNMRLastResultColSorted Then
        imNMRLastResultColSorted = NMRSORTINDEX
    Else
        imNMRLastResultColSorted = -1
        imNMRLastResultSort = -1
    End If
    gGrid_SortByCol grdNoMatchedResult, NMRFILENAMEINDEX, NMRSORTINDEX, imNMRLastResultColSorted, imNMRLastResultSort
    imNMRLastResultColSorted = ilCol
End Sub

Private Sub mMSSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdMatchedSpots.FixedRows To grdMatchedSpots.Rows - 1 Step 1
        slStr = Trim$(grdMatchedSpots.TextMatrix(llRow, MNETDPDAYSINDEX))
        If slStr <> "" Then
            If (ilCol = MNETDPTIMEINDEX) Then
                ilPos = InStr(1, grdMatchedSpots.TextMatrix(llRow, MNETDPTIMEINDEX), "-", vbTextCompare)
                If ilPos > 0 Then
                    slSort = Left$(grdMatchedSpots.TextMatrix(llRow, MNETDPTIMEINDEX), ilPos - 1)
                    slSort = gTimeToLong(slSort, False)
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = grdMatchedSpots.TextMatrix(llRow, MNETDPTIMEINDEX)
                End If
            ElseIf (ilCol = MSTNDATEINDEX) Then
                slSort = gDateValue(grdMatchedSpots.TextMatrix(llRow, MSTNDATEINDEX))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MLENGTHINDEX) Then
                slSort = gDateValue(grdMatchedSpots.TextMatrix(llRow, MLENGTHINDEX))
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = MSTNTIMEINDEX) Then
                'sort in reverse order
                slSort = gTimeToLong(grdMatchedSpots.TextMatrix(llRow, MSTNTIMEINDEX), False)
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = grdMatchedSpots.TextMatrix(llRow, ilCol)
            End If
            slStr = grdMatchedSpots.TextMatrix(llRow, MSSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imMSLastResultColSorted) Or ((ilCol = imMSLastResultColSorted) And (imMSLastResultSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMatchedSpots.TextMatrix(llRow, MSSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdMatchedSpots.TextMatrix(llRow, MSSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imMSLastResultColSorted Then
        imMSLastResultColSorted = MSSORTINDEX
    Else
        imMSLastResultColSorted = -1
        imMSLastResultSort = -1
    End If
    gGrid_SortByCol grdMatchedSpots, MNETDPDAYSINDEX, MSSORTINDEX, imMSLastResultColSorted, imMSLastResultSort
    imMSLastResultColSorted = ilCol
End Sub
Private Sub mCSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdCntrNetwork.FixedRows To grdCntrNetwork.Rows - 1 Step 1
        slStr = Trim$(grdCntrNetwork.TextMatrix(llRow, CADVERTISERINDEX))
        If slStr <> "" Then
            If (ilCol = CCONTRACTINDEX) Then
                slSort = grdCntrNetwork.TextMatrix(llRow, CCONTRACTINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = CORDERCOUNTINDEX) Then
                slSort = grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX)
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = CESTIMATEINDEX Then
                slSort = grdCntrNetwork.TextMatrix(llRow, CESTIMATEINDEX)
                Do While Len(slSort) < 20
                    slSort = " " & slSort
                Loop
            Else
                slSort = grdCntrNetwork.TextMatrix(llRow, ilCol)
            End If
            slStr = grdCntrNetwork.TextMatrix(llRow, CSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imCLastResultColSorted) Or ((ilCol = imCLastResultColSorted) And (imCLastResultSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCntrNetwork.TextMatrix(llRow, CSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCntrNetwork.TextMatrix(llRow, CSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imCLastResultColSorted Then
        imCLastResultColSorted = CSORTINDEX
    Else
        imCLastResultColSorted = -1
        imCLastResultSort = -1
    End If
    gGrid_SortByCol grdCntrNetwork, CADVERTISERINDEX, CSORTINDEX, imCLastResultColSorted, imCLastResultSort
    imCLastResultColSorted = ilCol
End Sub

Private Sub mGridNetworkColumns()

    grdNetworkSpots.Row = grdNetworkSpots.FixedRows - 1
    grdNetworkSpots.Col = NDPDAYSINDEX
    grdNetworkSpots.CellFontBold = False
    grdNetworkSpots.CellFontName = "Arial"
    grdNetworkSpots.CellFontSize = 6.75
    grdNetworkSpots.CellForeColor = vbBlue
    'grdNetworkSpots.CellBackColor = LIGHTBLUE
    grdNetworkSpots.TextMatrix(grdNetworkSpots.Row, grdNetworkSpots.Col) = "Days"
    grdNetworkSpots.Col = NDPTIMEINDEX
    grdNetworkSpots.CellFontBold = False
    grdNetworkSpots.CellFontName = "Arial"
    grdNetworkSpots.CellFontSize = 6.75
    grdNetworkSpots.CellForeColor = vbBlue
    'grdNetworkSpots.CellBackColor = LIGHTBLUE
    grdNetworkSpots.TextMatrix(grdNetworkSpots.Row, grdNetworkSpots.Col) = "Times"
    grdNetworkSpots.Col = NLENGTHINDEX
    grdNetworkSpots.CellFontBold = False
    grdNetworkSpots.CellFontName = "Arial"
    grdNetworkSpots.CellFontSize = 6.75
    grdNetworkSpots.CellForeColor = vbBlue
    'grdNetworkSpots.CellBackColor = LIGHTBLUE
    grdNetworkSpots.TextMatrix(grdNetworkSpots.Row, grdNetworkSpots.Col) = "Len"
    grdNetworkSpots.Col = NACQRATEINDEX
    grdNetworkSpots.CellFontBold = False
    grdNetworkSpots.CellFontName = "Arial"
    grdNetworkSpots.CellFontSize = 6.75
    grdNetworkSpots.CellForeColor = vbBlue
    'grdNetworkSpots.CellBackColor = LIGHTBLUE
    grdNetworkSpots.TextMatrix(grdNetworkSpots.Row, grdNetworkSpots.Col) = "Acq $"
    grdNetworkSpots.Col = NDATESINDEX
    grdNetworkSpots.CellFontBold = False
    grdNetworkSpots.CellFontName = "Arial"
    grdNetworkSpots.CellFontSize = 6.75
    grdNetworkSpots.CellForeColor = vbBlue
    'grdNetworkSpots.CellBackColor = LIGHTBLUE
    grdNetworkSpots.TextMatrix(grdNetworkSpots.Row, grdNetworkSpots.Col) = "Dates"
End Sub

Private Sub mGridNetworkColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdNetworkSpots.ColWidth(NSDFCODEINDEX) = 0
    grdNetworkSpots.ColWidth(NSELECTEDINDEX) = 0
    grdNetworkSpots.ColWidth(NDPDAYSINDEX) = 0.2 * grdNetworkSpots.Width
    grdNetworkSpots.ColWidth(NDPTIMEINDEX) = 0.2 * grdNetworkSpots.Width
    grdNetworkSpots.ColWidth(NLENGTHINDEX) = 0.1 * grdNetworkSpots.Width
    grdNetworkSpots.ColWidth(NACQRATEINDEX) = 0.2 * grdNetworkSpots.Width
    grdNetworkSpots.ColWidth(NDATESINDEX) = 0.25 * grdNetworkSpots.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdNetworkSpots.Width
    For ilCol = 0 To grdNetworkSpots.Cols - 1 Step 1
        llWidth = llWidth + grdNetworkSpots.ColWidth(ilCol)
        If (grdNetworkSpots.ColWidth(ilCol) > 15) And (grdNetworkSpots.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdNetworkSpots.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdNetworkSpots.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdNetworkSpots.Width
            For ilCol = 0 To grdNetworkSpots.Cols - 1 Step 1
                If (grdNetworkSpots.ColWidth(ilCol) > 15) And (grdNetworkSpots.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdNetworkSpots.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdNetworkSpots.FixedCols To grdNetworkSpots.Cols - 1 Step 1
                If grdNetworkSpots.ColWidth(ilCol) > 15 Then
                    ilColInc = grdNetworkSpots.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdNetworkSpots.ColWidth(ilCol) = grdNetworkSpots.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub
Private Sub mGridStationColumns()

    grdStationSpots.Row = grdStationSpots.FixedRows - 1
    grdStationSpots.Col = SDATEINDEX
    grdStationSpots.CellFontBold = False
    grdStationSpots.CellFontName = "Arial"
    grdStationSpots.CellFontSize = 6.75
    grdStationSpots.CellForeColor = vbBlue
    'grdStationSpots.CellBackColor = LIGHTBLUE
    grdStationSpots.TextMatrix(grdStationSpots.Row, grdStationSpots.Col) = "Date"
    grdStationSpots.Col = STIMEINDEX
    grdStationSpots.CellFontBold = False
    grdStationSpots.CellFontName = "Arial"
    grdStationSpots.CellFontSize = 6.75
    grdStationSpots.CellForeColor = vbBlue
    'grdStationSpots.CellBackColor = LIGHTBLUE
    grdStationSpots.TextMatrix(grdStationSpots.Row, grdStationSpots.Col) = "Time"
    grdStationSpots.Col = SLENGTHINDEX
    grdStationSpots.CellFontBold = False
    grdStationSpots.CellFontName = "Arial"
    grdStationSpots.CellFontSize = 6.75
    grdStationSpots.CellForeColor = vbBlue
    'grdStationSpots.CellBackColor = LIGHTBLUE
    grdStationSpots.TextMatrix(grdStationSpots.Row, grdStationSpots.Col) = "Len"
    grdStationSpots.Col = SACQRATEINDEX
    grdStationSpots.CellFontBold = False
    grdStationSpots.CellFontName = "Arial"
    grdStationSpots.CellFontSize = 6.75
    grdStationSpots.CellForeColor = vbBlue
    'grdStationSpots.CellBackColor = LIGHTBLUE
    grdStationSpots.TextMatrix(grdStationSpots.Row, grdStationSpots.Col) = "Acq $"
    grdStationSpots.Col = SISCIINDEX
    grdStationSpots.CellFontBold = False
    grdStationSpots.CellFontName = "Arial"
    grdStationSpots.CellFontSize = 6.75
    grdStationSpots.CellForeColor = vbBlue
    'grdStationSpots.CellBackColor = LIGHTBLUE
    grdStationSpots.TextMatrix(grdStationSpots.Row, grdStationSpots.Col) = "ISCI"
End Sub

Private Sub mGridStationColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdStationSpots.ColWidth(SIIDFCODEINDEX) = 0
    grdStationSpots.ColWidth(SSELECTEDINDEX) = 0
    grdStationSpots.ColWidth(SDATEINDEX) = 0.15 * grdStationSpots.Width
    grdStationSpots.ColWidth(STIMEINDEX) = 0.15 * grdStationSpots.Width
    grdStationSpots.ColWidth(SLENGTHINDEX) = 0.1 * grdStationSpots.Width
    grdStationSpots.ColWidth(SACQRATEINDEX) = 0.1 * grdStationSpots.Width
    grdStationSpots.ColWidth(SISCIINDEX) = 0.4 * grdStationSpots.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdStationSpots.Width
    For ilCol = 0 To grdStationSpots.Cols - 1 Step 1
        llWidth = llWidth + grdStationSpots.ColWidth(ilCol)
        If (grdStationSpots.ColWidth(ilCol) > 15) And (grdStationSpots.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdStationSpots.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdStationSpots.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdStationSpots.Width
            For ilCol = 0 To grdStationSpots.Cols - 1 Step 1
                If (grdStationSpots.ColWidth(ilCol) > 15) And (grdStationSpots.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdStationSpots.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdStationSpots.FixedCols To grdStationSpots.Cols - 1 Step 1
                If grdStationSpots.ColWidth(ilCol) > 15 Then
                    ilColInc = grdStationSpots.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdStationSpots.ColWidth(ilCol) = grdStationSpots.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub mGridMatchedColumns()

    grdMatchedSpots.Row = grdMatchedSpots.FixedRows - 2
    grdMatchedSpots.Col = MNETDPDAYSINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Network"
    grdMatchedSpots.Col = MNETDPTIMEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Network"
    grdMatchedSpots.Col = MNETDATESINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Network"
    grdMatchedSpots.Col = MSTNDATEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Station"
    grdMatchedSpots.Col = MSTNTIMEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Station"
    grdMatchedSpots.Col = MSTNCOMPLIANTINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Station"
    grdMatchedSpots.Col = MLENGTHINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Spot"
    grdMatchedSpots.Col = MACQRATEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Acquisition"
    grdMatchedSpots.Col = MSTNISCIINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Station"

    grdMatchedSpots.Row = grdMatchedSpots.FixedRows - 1
    grdMatchedSpots.Col = MNETDPDAYSINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Days"
    grdMatchedSpots.Col = MNETDPTIMEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Times"
    grdMatchedSpots.Col = MNETDATESINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Dates"
    grdMatchedSpots.Col = MSTNDATEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Air Date"
    grdMatchedSpots.Col = MSTNTIMEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Air Time"
    grdMatchedSpots.Col = MSTNCOMPLIANTINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Compliant"
    grdMatchedSpots.Col = MLENGTHINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Len"
    grdMatchedSpots.Col = MACQRATEINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "Dollars"
    grdMatchedSpots.Col = MSTNISCIINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "ISCI"
    grdMatchedSpots.Col = MSTNISCIINDEX
    grdMatchedSpots.CellFontBold = False
    grdMatchedSpots.CellFontName = "Arial"
    grdMatchedSpots.CellFontSize = 6.75
    grdMatchedSpots.CellForeColor = vbBlue
    'grdMatchedSpots.CellBackColor = LIGHTBLUE
    grdMatchedSpots.TextMatrix(grdMatchedSpots.Row, grdMatchedSpots.Col) = "ISCI"

End Sub

Private Sub mGridMatchedColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdMatchedSpots.ColWidth(MSDFCODEINDEX) = 0
    grdMatchedSpots.ColWidth(MIIDFCODEINDEX) = 0
    grdMatchedSpots.ColWidth(MSSORTINDEX) = 0
    grdMatchedSpots.ColWidth(MSELECTEDINDEX) = 0
    grdMatchedSpots.ColWidth(MNETDPDAYSINDEX) = 0.1 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MNETDPTIMEINDEX) = 0.1 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MNETDATESINDEX) = 0.1 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MSTNDATEINDEX) = 0.1 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MSTNTIMEINDEX) = 0.1 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MSTNCOMPLIANTINDEX) = 0.06 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MLENGTHINDEX) = 0.04 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MACQRATEINDEX) = 0.06 * grdMatchedSpots.Width
    grdMatchedSpots.ColWidth(MSTNISCIINDEX) = 0.25 * grdMatchedSpots.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdMatchedSpots.Width
    For ilCol = 0 To grdMatchedSpots.Cols - 1 Step 1
        llWidth = llWidth + grdMatchedSpots.ColWidth(ilCol)
        If (grdMatchedSpots.ColWidth(ilCol) > 15) And (grdMatchedSpots.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdMatchedSpots.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdMatchedSpots.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdMatchedSpots.Width
            For ilCol = 0 To grdMatchedSpots.Cols - 1 Step 1
                If (grdMatchedSpots.ColWidth(ilCol) > 15) And (grdMatchedSpots.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdMatchedSpots.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdMatchedSpots.FixedCols To grdMatchedSpots.Cols - 1 Step 1
                If grdMatchedSpots.ColWidth(ilCol) > 15 Then
                    ilColInc = grdMatchedSpots.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdMatchedSpots.ColWidth(ilCol) = grdMatchedSpots.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Function mSaveRec() As Integer

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer control values to     *
'*                      records                        *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()

    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(ilRowNo As Integer) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    slStr = Trim$(grdMatchedResult.TextMatrix(ilRowNo, MRSTNSTATIONINDEX))
    If slStr <> "" Then
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function
Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdMatchedResult.Row = llRow
    For llCol = MRNETADVERTISERINDEX To MRFILENAMEINDEX Step 1
        grdMatchedResult.Col = llCol
        If grdMatchedResult.TextMatrix(llRow, MRSELECTEDINDEX) <> "1" Then
            grdMatchedResult.CellBackColor = vbWhite    'LIGHTYELLOW
        Else
            grdMatchedResult.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub
Private Sub mNPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdNetworkSpots.Row = llRow
    For llCol = NDPDAYSINDEX To NDATESINDEX Step 1
        grdNetworkSpots.Col = llCol
        If grdNetworkSpots.TextMatrix(llRow, NSELECTEDINDEX) <> "1" Then
            grdNetworkSpots.CellBackColor = vbWhite 'LIGHTYELLOW
        Else
            grdNetworkSpots.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub
Private Sub mSPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdStationSpots.Row = llRow
    For llCol = SDATEINDEX To SISCIINDEX Step 1
        grdStationSpots.Col = llCol
        If grdStationSpots.TextMatrix(llRow, SSELECTEDINDEX) <> "1" Then
            grdStationSpots.CellBackColor = vbWhite 'IGHTYELLOW
        Else
            grdStationSpots.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub
Private Sub mMPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdMatchedSpots.Row = llRow
    For llCol = MNETDPDAYSINDEX To MSTNISCIINDEX Step 1
        grdMatchedSpots.Col = llCol
        If grdMatchedSpots.TextMatrix(llRow, MSELECTEDINDEX) <> "1" Then
            grdMatchedSpots.CellBackColor = vbWhite 'LIGHTYELLOW
        Else
            grdMatchedSpots.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub
Private Sub mCPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdCntrNetwork.Row = llRow
    For llCol = CADVERTISERINDEX To CPREVIEWINDEX Step 1
        grdCntrNetwork.Col = llCol
        If llCol <> CPREVIEWINDEX Then
            If grdCntrNetwork.TextMatrix(llRow, CSELECTEDINDEX) <> "1" Then
                grdCntrNetwork.CellBackColor = vbWhite    'LIGHTYELLOW
            Else
                grdCntrNetwork.CellBackColor = GRAY    'vbBlue
            End If
        Else
            grdCntrNetwork.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub
Private Sub mResizeForm()
    Me.Width = (CLng(90) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = PostLog.Height '(lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    Me.Top = PostLog.Top
    
    frcResult.Move 60, 360, Me.Width - 240, Me.Height - 300 - 240
    lacUnresolvedImports.Move 0, 0, frcResult.Width
    grdNoMatchedResult.Move 0, lacUnresolvedImports.Height + 30, frcResult.Width, frcResult.Height / 3 - 2 * cmcDetail.Height - lacUnresolvedImports.Height
    lacResolvedImports.Move 0, grdNoMatchedResult.Height + lacUnresolvedImports.Height, frcResult.Width
    grdMatchedResult.Move 0, grdNoMatchedResult.Height + 2 * lacUnresolvedImports.Height + 30, frcResult.Width, frcResult.Height - grdNoMatchedResult.Height - 2 * cmcDetail.Height - 2 * lacUnresolvedImports.Height
    cmcDetail.Move frcResult.Width / 2 - cmcDetail.Width / 2, frcResult.Height - cmcDetail.Height
    cmcCancel.Move cmcDetail.Left - (3 * cmcDetail.Width / 2), cmcDetail.Top
    cmcUndoMatch.Move cmcDetail.Left + (3 * cmcDetail.Width / 2), cmcDetail.Top
    
    mGridResultColumnWidths
    mGridResultColumns
    
    gGrid_IntegralHeight grdNoMatchedResult, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdNoMatchedResult, fgBoxGridH + 15
    mClearGrid grdNoMatchedResult
    grdNoMatchedResult.Height = grdNoMatchedResult.Height + 30
    gGrid_AlignAllColsLeft grdNoMatchedResult
    
    gGrid_IntegralHeight grdMatchedResult, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdMatchedResult, fgBoxGridH + 15
    mClearGrid grdMatchedResult
    grdMatchedResult.Height = grdMatchedResult.Height + 30
    gGrid_AlignAllColsLeft grdMatchedResult
    
    ''frcDetail.Move 60, 300, Me.Width - 240, Me.Height - 300 - 120
    'frcDetail.Move Me.Width / 6, 120, (2 * Me.Width) / 3, Me.Height - 300 - 120
    frcDetail.Move (Me.Width - (0.7 * Me.Width)) / 2, 120, (0.7 * Me.Width), Me.Height - 300 - 120
    grdNetworkSpots.Move 0, 2 * lacDetail.Height, (0.4 * frcDetail.Width) - 120, frcDetail.Height / 3 - cmcDetail.Height - 90
    grdStationSpots.Move grdNetworkSpots.Width + 90, grdNetworkSpots.Top, frcDetail.Width - grdNetworkSpots.Width, grdNetworkSpots.Height
    
    lacDetail.Move 0, 0, frcDetail.Width
    lacNetSpots.Move grdNetworkSpots.Left, lacDetail.Height, grdNetworkSpots.Width
    lacStnSpots.Move grdStationSpots.Left, lacDetail.Height, grdStationSpots.Width
    
    mGridNetworkColumnWidths
    mGridNetworkColumns
    
    mGridStationColumnWidths
    mGridStationColumns
    
    gGrid_IntegralHeight grdNetworkSpots, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdNetworkSpots, fgBoxGridH + 15
    mClearGrid grdNetworkSpots
    grdNetworkSpots.Height = grdNetworkSpots.Height + 15
    gGrid_AlignAllColsLeft grdNetworkSpots
    
    gGrid_IntegralHeight grdStationSpots, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdStationSpots, fgBoxGridH + 15
    mClearGrid grdStationSpots
    grdStationSpots.Height = grdStationSpots.Height + 15
    gGrid_AlignAllColsLeft grdStationSpots
    
    'grdMatchedSpots.Move 0, grdNetworkSpots.Top + grdNetworkSpots.Height + 2 * cmcSave.Height, frcDetail.Width
    'grdMatchedSpots.Height = frcDetail.Height - grdMatchedSpots.Top - 2 * cmcSave.Height - 120
    grdMatchedSpots.Move 0, grdNetworkSpots.Top + grdNetworkSpots.Height + (3 * cmcReturn.Height) / 2, frcDetail.Width
    grdMatchedSpots.Height = frcDetail.Height - grdMatchedSpots.Top - (3 * cmcReturn.Height) / 2
    
    lacReconciledSpots.Move grdMatchedSpots.Left, grdMatchedSpots.Top - lacReconciledSpots.Height, grdMatchedSpots.Width
    'cmcReturn.Move frcDetail.Width / 2 - (3 * cmcReturn.Width) / 2, frcDetail.Height - cmcReturn.Height - 60
    'cmcUndoReconcile.Move cmcReturn.Left + cmcReturn.Width + cmcUndoReconcile.Width / 2, cmcReturn.Top
    'cmcReconcile.Move frcDetail.Width / 2 - cmcReconcile.Width / 2, (grdNetworkSpots.Top + grdNetworkSpots.Height) + (grdMatchedSpots.Top - (grdNetworkSpots.Top + grdNetworkSpots.Height)) / 2 - cmcReconcile.Height / 2
    mSetDetailButtons
    
    mGridMatchedColumnWidths
    mGridMatchedColumns
    
    gGrid_IntegralHeight grdMatchedSpots, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdMatchedSpots, fgBoxGridH + 15
    mClearGrid grdMatchedSpots
    grdMatchedSpots.Height = grdMatchedSpots.Height + 15
    gGrid_AlignAllColsLeft grdMatchedSpots
    
    
    frcCntr.Move 60, 360, Me.Width - 240, Me.Height - 300 - 240
    lacSelectedUnresolvedImport.Move 0, 0, frcCntr.Width
    grdCntrStation.Move 0, lacSelectedUnresolvedImport.Height + 30, frcCntr.Width
    lacPossibleCntr.Move 0, grdCntrStation.Height + lacSelectedUnresolvedImport.Height, frcCntr.Width
    grdCntrNetwork.Width = (4 * frcCntr.Width / 5)
    grdCntrNetwork.Move (frcCntr.Width - grdCntrNetwork.Width) / 2, grdCntrStation.Height + 2 * lacSelectedUnresolvedImport.Height + 30, grdCntrNetwork.Width, frcCntr.Height - grdCntrStation.Height - 2 * lacSelectedUnresolvedImport.Height - cmcCntrMatch.Height - 90
    cmcCntrMatch.Move frcCntr.Width / 2 - 2 * cmcCntrMatch.Width, frcCntr.Height - cmcCntrMatch.Height
    cmcCntrCancel.Move frcCntr.Width / 2 + cmcCntrCancel.Width, cmcCntrMatch.Top
    
    mGridCntrColumnWidths
    mGridCntrColumns
    
    
    gGrid_FillWithRows grdCntrStation, fgBoxGridH + 15
    mClearGrid grdCntrStation
    grdCntrStation.Height = 2 * grdCntrStation.RowHeight(0) + 30
    gGrid_AlignAllColsLeft grdCntrStation
    
    gGrid_IntegralHeight grdCntrNetwork, CInt(fgBoxGridH + 45) ' + 15
    gGrid_FillWithRows grdCntrNetwork, fgBoxGridH + 15
    mClearGrid grdCntrNetwork
    grdCntrNetwork.Height = grdCntrNetwork.Height + 15
    gGrid_AlignAllColsLeft grdCntrNetwork
    
    pbcClickFocus.Left = -pbcClickFocus.Width
    pbcSetFocus.Left = -pbcSetFocus.Width
    
    gCenterStdAlone ImportStationSpots
    
    Me.Top = Me.Top - 210
End Sub

Private Sub tmcStart_Timer()
    Dim ilFile As Integer
    Dim llNoMatchResultRow As Long
    Dim llMatchResultRow As Long
    Dim ilRet As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slInvoiceType As String
    Dim slRecord As String
    Dim blCreateResultFile As Boolean
    Dim slSubfolderPath As String
    Dim fs As New FileSystemObject
    
    tmcStart.Enabled = False
    
    If Not IsArray(smImportFiles) Then
        Exit Sub
    End If
    edcProcessing.Move Me.Width / 2 - edcProcessing.Width / 2, Me.Height / 2 - edcProcessing.Height / 2
    edcProcessing.Visible = True
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbHourglass
    llNoMatchResultRow = grdNoMatchedResult.FixedRows
    llMatchResultRow = grdMatchedResult.FixedRows
    blCreateResultFile = True
    For ilFile = LBound(smImportFiles) To UBound(smImportFiles) - 1 Step 1
        If InStr(1, sgBrowserDrivePath, "StationInvoices-", vbTextCompare) <= 0 Then
            If InStr(1, UCase$(smImportFiles(ilFile)), ".PDF", vbTextCompare) > 0 Then
                edcProcessing.Text = "Converting PDF to TEXT: " & smImportFiles(ilFile) & sgCR & sgLF & ilFile + 1 & " of " & UBound(smImportFiles)
                gShellAndWait ImportStationSpots, sgExePath & "PDFToText.exe" & " -table -clip -eol DOS " & """" & sgBrowserDrivePath & smImportFiles(ilFile) & """", vbMinimizedFocus, True    'vbTrue
            End If
            edcProcessing.Text = "Processing: " & smImportFiles(ilFile) & sgCR & sgLF & ilFile + 1 & " of " & UBound(smImportFiles)
            mProcessImport sgBrowserDrivePath & smImportFiles(ilFile), smImportFiles(ilFile), llMatchResultRow, llNoMatchResultRow
        Else
            blCreateResultFile = False
            edcProcessing.Text = "Processing: " & smImportFiles(ilFile) & sgCR & sgLF & ilFile + 1 & " of " & UBound(smImportFiles)
            mShowPreviousResults smImportFiles(ilFile), llMatchResultRow, llNoMatchResultRow
        End If
    Next ilFile
    mMRSortCol MRCOMPLIANTINDEX
    mMRSortCol MRNETCOUNTINDEX

    ilRet = mUpdateApf(-1)

    If blCreateResultFile Then
        slSubfolderPath = sgBrowserDrivePath & "Import_Results"
        If Not fs.FolderExists(slSubfolderPath) Then
            fs.CreateFolder (slSubfolderPath)
        End If
    
        gLogMsgWODT "O", hmUnMatch, slSubfolderPath & "\" & "InvoiceImport_Unmatched_Result_" & Format(Now, "mmddyyyy") & ".csv"
        gLogMsgWODT "W", hmUnMatch, "Invoice Import Unmatched Results " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
        
        gLogMsgWODT "W", hmUnMatch, ""
        gLogMsgWODT "W", hmUnMatch, "Station,Advertiser,Estimate #,Contract #,Invoice #,Status,Count,File Name,Invoice Type"
        For llRow = grdNoMatchedResult.FixedRows To grdNoMatchedResult.Rows - 1 Step 1
            If grdNoMatchedResult.TextMatrix(llRow, NMRFILENAMEINDEX) <> "" Then
                slInvoiceType = ""
                tmIihfSrchKey0.lCode = Val(grdNoMatchedResult.TextMatrix(llRow, NMRIIHFCODEINDEX))
                ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    Select Case Trim$(tmIihf.sSourceForm)
                        Case "M"
                            slInvoiceType = "Marketron"
                        Case "W"
                            slInvoiceType = "Wide Orbit"
                        Case "R"
                            slInvoiceType = "Radio Traffic"
                        Case "V"
                            slInvoiceType = "Visual Traffic"
                        Case "N"
                            slInvoiceType = "Natural Log"
                        Case "WE"
                            slInvoiceType = "Wide Orbit-EDI"
                        Case "ME"
                            slInvoiceType = "Marketron-EDI"
                    End Select
                End If
                slRecord = ""
                For ilCol = 0 To NMRFILENAMEINDEX Step 1
                    slRecord = slRecord & """" & Trim$(grdNoMatchedResult.TextMatrix(llRow, ilCol)) & """" & ","
                Next ilCol
                slRecord = slRecord & """" & slInvoiceType & """"
                gLogMsgWODT "W", hmUnMatch, slRecord
            End If
        Next llRow
        gLogMsgWODT "C", hmUnMatch, ""
        
        gLogMsgWODT "O", hmMatch, slSubfolderPath & "\" & "InvoiceImport_Matched_Result_" & Format(Now, "mmddyyyy") & ".csv"
        gLogMsgWODT "W", hmMatch, "Invoice Import Matched Results " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
        
        gLogMsgWODT "W", hmMatch, ""
        gLogMsgWODT "W", hmMatch, "Net Advertiser,Net Estimate #,Net Contract #,Station,Advertiser,Estimate #,Contract #,Invoice #,Status, Match Count, Net Unmatched, Station Unmatched,Compliant,File Name,Invoice Type"
        For llRow = grdMatchedResult.FixedRows To grdMatchedResult.Rows - 1 Step 1
            If grdMatchedResult.TextMatrix(llRow, MRFILENAMEINDEX) <> "" Then
                slInvoiceType = ""
                tmIihfSrchKey0.lCode = Val(grdMatchedResult.TextMatrix(llRow, MRIIHFCODEINDEX))
                ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    Select Case Trim$(tmIihf.sSourceForm)
                        Case "M"
                            slInvoiceType = "Marketron"
                        Case "W"
                            slInvoiceType = "Wide Orbit"
                        Case "R"
                            slInvoiceType = "Radio Traffic"
                        Case "V"
                            slInvoiceType = "Visual Traffic"
                        Case "N"
                            slInvoiceType = "Natural Log"
                        Case "WE"
                            slInvoiceType = "Wide Orbit-EDI"
                        Case "ME"
                            slInvoiceType = "Marketron-EDI"
                    End Select
                End If
                slRecord = ""
                For ilCol = 0 To MRFILENAMEINDEX Step 1
                    slRecord = slRecord & """" & Trim$(grdMatchedResult.TextMatrix(llRow, ilCol)) & """" & ","
                Next ilCol
                slRecord = slRecord & """" & slInvoiceType & """"
                gLogMsgWODT "W", hmMatch, slRecord
            End If
        Next llRow
        gLogMsgWODT "C", hmMatch, ""
    End If
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbDefault
    edcProcessing.Visible = False
End Sub

Private Sub mMousePointer(grdCtrl1 As MSHFlexGrid, grdCtrl2 As MSHFlexGrid, ilSet As Integer)
    Screen.MousePointer = ilSet
    gSetMousePointer grdCtrl1, grdCtrl2, ilSet
    Exit Sub
End Sub

Private Sub mProcessImport(slINTextFile As String, slPDFFileName As String, llMatchResultRow As Long, llNoMatchResultRow As Long)
    Dim slTextFile As String
    Dim slLine As String
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim ilPos3 As Integer
    Dim ilPos4 As Integer
    Dim ilPos5 As Integer
    Dim ilPos6 As Integer
    Dim ilPos7 As Integer
    Dim ilRatePos As Integer
    Dim ilDecPointPos As Integer
    Dim ilLnPos1 As Integer
    Dim ilLnPos2 As Integer
    Dim ilLnPos3 As Integer
    Dim ilLnPos4 As Integer
    Dim ilLnPos5 As Integer
    Dim ilLnPos6 As Integer
    Dim ilSpotPos As Integer
    Dim ilAMPMPos As Integer
    Dim slChar As String
    Dim ilChar As Integer
    Dim slSvLine As String
    ReDim slPrevLines(0 To 5) As String
    Dim ilUpper As Integer
    Dim ilUpperStart As Integer
    Dim blEndIfInvoice As Boolean
    Dim ilLen As Integer
    Dim slDay As String
    Dim ilLine As Integer
    Dim slStr As String
    Dim blLine As Boolean
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilRet As Integer
    Dim blAnyFound As Boolean
    Dim blAnyFoundSetToFalse As Boolean
    Dim slMarketronForm As String
    Dim slWideOrbitForm As String
    Dim blMarketronSkipRead As Boolean
    Dim slISCI As String
    Dim blFirstSpot As Boolean
    Dim blFound As Boolean
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim blInvoiceProcessed As Boolean
    Dim ilLoop As Integer
    Dim slCallLetters As String
    Dim slPoss1CallLetters As String
    Dim slPoss2CallLetters As String
    Dim ilVef As Integer
    Dim slPDForEDI As String * 1
    Dim slFields(0 To 40) As String
    Dim llDPStartTime As Long
    Dim llDPEndTime As Long
    Dim slDPDays As String
    Dim bl41First As Boolean
    Dim blLineHeaderFound As Boolean
    Dim ilCol As Integer
    Dim slAltOrderNumber As String

    smIihfFileName = Trim$(slINTextFile)
    ilPos1 = InStrRev(smIihfFileName, "\", -1, vbBinaryCompare)
    If ilPos1 > 0 Then
        smIihfFileName = Trim$(Mid$(smIihfFileName, ilPos1 + 1))
    End If
    If llMatchResultRow >= grdMatchedResult.Rows Then
        grdMatchedResult.AddItem ""
    End If
    grdMatchedResult.RowHeight(llMatchResultRow) = fgBoxGridH + 15
    'grdMatchedResult.TextMatrix(llMatchResultRow, MRFILENAMEINDEX) = smIihfFileName  'Trim$(slINTextFile)
    grdMatchedResult.TextMatrix(llMatchResultRow, MRSELECTEDINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchResultRow, MRIIHFCODEINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchResultRow, MRCHFCODEINDEX) = "0"
    smSourceForm = ""
    smInvoiceNumber = ""
    smContractNumber = ""
    smAdvertiserName = ""
    smEstimateNumber = ""
    smCallLetters = ""
    slAltOrderNumber = ""
    llDPStartTime = -1
    llDPEndTime = -1
    slDPDays = ""
    bl41First = True
    smNetContractNumber = ""
    For ilLoop = 0 To UBound(slPrevLines) Step 1
        slPrevLines(ilLoop) = ""
    Next ilLoop
    blInvoiceProcessed = False
    llStartDate = 99999999
    llEndDate = 0
    lmInvStartDate = llStartDate
    ReDim tmImportSpotInfo(0 To 0) As IMPORTSPOTINFO
    slTextFile = slINTextFile
    ilPos1 = InStr(1, slTextFile, "..", vbBinaryCompare)
    If ilPos1 <= 0 Then
        ilPos1 = InStr(1, slTextFile, ".", vbBinaryCompare)
    Else
        ilPos1 = ilPos1 + 1
    End If
    If ilPos1 <= 0 Then
        'Output error to grdNoMatchedResult
        If llNoMatchResultRow >= grdNoMatchedResult.Rows Then
            grdNoMatchedResult.AddItem ""
        End If
        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRFILENAMEINDEX) = slPDFFileName
        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "No Extension"
        llNoMatchResultRow = llNoMatchResultRow + 1
        Exit Sub
    End If
    If InStr(1, slTextFile, ".PDF", vbTextCompare) > 0 Then
        slTextFile = Left(slTextFile, ilPos1) & "Txt"
        slPDForEDI = "P"
    Else
        slPDForEDI = "E"
    End If
    
    'smIihfFileName = slINTextFile
    'ilPos1 = InStrRev(smIihfFileName, "\", -1, vbBinaryCompare)
    'If ilPos1 > 0 Then
    '    smIihfFileName = Mid$(smIihfFileName, ilPos1 + 1)
    'End If
    blAnyFound = False
    blAnyFoundSetToFalse = False
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slTextFile) Then
        Set MyFile = oMyFileObj.OpenTextFile(slTextFile, ForReading, False)
        slLine = MyFile.ReadLine
        Do While Not MyFile.AtEndOfStream
            slLine = UCase(slLine)
            'Process lines
            If smSourceForm = "" Then
                If ((InStr(1, slLine, "INVOICE SUMMARY", vbBinaryCompare) > 0) And (slPDForEDI = "P")) Or ((InStr(1, slLine, "34:", vbBinaryCompare) > 0) And (slPDForEDI = "E")) Then
                    MyFile.Close
                    Set MyFile = Nothing
                    'Set result to File does not exist in grid
                    If llNoMatchResultRow >= grdNoMatchedResult.Rows Then
                        grdNoMatchedResult.AddItem ""
                    End If
                    grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRFILENAMEINDEX) = smIihfFileName
                    grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Not Invoice File"
                    llNoMatchResultRow = llNoMatchResultRow + 1
                    Set oMyFileObj = Nothing
                    If Not MyFile Is Nothing Then
                        On Error Resume Next
                        MyFile.Close
                        On Error GoTo 0
                        Set MyFile = Nothing
                    End If
                    mMoveFile "U", slPDFFileName, lmInvStartDate
                    Exit Sub
                End If
                'grdMatchedResult.RowHeight(llMatchResultRow) = fgBoxGridH + 15
                'grdMatchedResult.TextMatrix(llMatchResultRow, MRFILENAMEINDEX) = smIihfFileName  'Trim$(slINTextFile)
                'grdMatchedResult.TextMatrix(llMatchResultRow, MRSELECTEDINDEX) = "0"
                'grdMatchedResult.TextMatrix(llMatchResultRow, MRIIHFCODEINDEX) = "0"
                ilPos1 = InStr(1, slLine, "INVOICE #:", vbBinaryCompare)
                If (ilPos1 > 0) And (slPDForEDI = "P") Then  'Marketron
                    smSourceForm = "M"
                    smInvoiceNumber = Trim$(Mid$(slLine, ilPos1 + 10))
                    smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                    Do
                        slLine = MyFile.ReadLine
                        slLine = UCase$(slLine)
                        ilPos1 = InStr(1, slLine, "CONTRACT #:", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smContractNumber = Trim$(Mid$(slLine, ilPos1 + 11))
                            smContractNumber = mRemoveBlanks(smContractNumber)
                            Exit Do
                        End If
                    Loop While Not MyFile.AtEndOfStream
                    Do
                        slLine = MyFile.ReadLine
                        slLine = UCase$(slLine)
                        ilPos1 = InStr(1, slLine, "STATION(S):", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smCallLetters = Trim$(Mid$(slLine, ilPos1 + 11))
                            smCallLetters = mRemoveBlanks(smCallLetters)
                            Exit Do
                        End If
                    Loop While Not MyFile.AtEndOfStream
                ElseIf (InStr(1, slLine, "22;", vbTextCompare) > 0) And (slPDForEDI = "E") Then
                    '22;call letters; media type; band; station name; addr 1; addr 2; addr 3; addr 4;station computer system (mkt=marketron; wos=wide orbit
                    gParseItemFields slLine, ";", slFields()
                    For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
                        slFields(ilCol + 1) = slFields(ilCol)
                    Next ilCol
                    slFields(0) = ""
                    smCallLetters = mRemoveBlanks(slFields(5))
                    smSourceForm = Left$(slFields(10), 1) & "E"
                ElseIf (slPDForEDI = "P") Then
                    blFound = False
                    ilPos1 = InStr(1, slLine, "INVOICE", vbBinaryCompare)
                    ilPos2 = InStr(1, slLine, "#", vbBinaryCompare)
                    If (ilPos1 > 0) And (ilPos1 + 8 <= ilPos2) And (InStr(1, slLine, ":", vbBinaryCompare) <= 0) And (InStr(1, slLine, "ORDER", vbBinaryCompare) <= 0) Then
                        ilPos3 = ilPos2 + 1
                        Do
                            If ilPos3 > Len(slLine) Then
                                Exit Do
                            End If
                            slChar = Mid(slLine, ilPos3, 1)
                            If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Then
                                blFound = True
                                Exit Do
                            End If
                            ilPos3 = ilPos3 + 1
                        Loop
                    End If
                    If (ilPos1 > 0) And (ilPos1 + 8 <= ilPos2) And (InStr(1, slLine, ":", vbBinaryCompare) <= 0) And (InStr(1, slLine, "ORDER", vbBinaryCompare) <= 0) And (blFound) Then
                        'Radio Traffic
                        blFound = False
                        ilPos1 = ilPos2 + 1
                        Do
                            If ilPos1 > Len(slLine) Then
                                Exit Do
                            End If
                            slChar = Mid(slLine, ilPos1, 1)
                            If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Then
                                blFound = True
                                Exit Do
                            End If
                            ilPos1 = ilPos1 + 1
                        Loop
                        'If blFound Then
                            smSourceForm = "R"
                            smInvoiceNumber = Mid$(slLine, ilPos1)
                            smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                If Len(Trim$(slLine)) > 0 Then
                                    ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        smCallLetters = Trim$(Left$(slLine, ilPos1 - 1))
                                        If Len(smCallLetters) > 0 Then
                                            Exit Do
                                        End If
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                            smCallLetters = mRemoveBlanks(smCallLetters)
                            If Not MyFile.AtEndOfStream Then
                                Do
                                    slLine = MyFile.ReadLine
                                    slLine = UCase$(slLine)
                                    If Len(Trim$(slLine)) > 0 Then
                                        ilPos1 = InStr(1, slLine, "ESTIMATE #", vbBinaryCompare)
                                        If ilPos1 > 0 Then
                                            smEstimateNumber = Trim$(Mid$(slLine, ilPos1 + 10))
                                            smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                            Exit Do
                                        End If
                                    End If
                                Loop While Not MyFile.AtEndOfStream
                                If smEstimateNumber = "" Then
                                    smEstimateNumber = "Not Specified"
                                End If
                            End If
                        'End If
                   ElseIf (slPDForEDI = "P") Then
                        ilPos1 = InStr(1, slLine, "INVOICE #", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            'Wide Orbit
                            ilPos2 = InStr(ilPos1, slLine, "ORDER #", vbBinaryCompare)
                            If ilPos2 > 0 Then
                                smSourceForm = "W"
                                smCallLetters = Trim$(Left(slLine, ilPos1 - 1))
                                smCallLetters = mRemoveBlanks(smCallLetters)
                                smInvoiceNumber = Trim$(Mid$(slLine, ilPos1 + Len("Inovice #"), ilPos2 - (ilPos1 + 9)))
                                smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                                smContractNumber = Trim$(Mid(slLine, ilPos2 + Len("Order #")))
                                If Len(smContractNumber) > 0 Then
                                    smContractNumber = mRemoveBlanks(smContractNumber)
                                End If
                                slWideOrbitForm = "3"
                            Else
                                slLine = MyFile.ReadLine    'Blank line
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                ilPos2 = InStr(ilPos1, slLine, " ", vbBinaryCompare)
                                If ilPos2 > 0 Then
                                    smSourceForm = "W"
                                    smInvoiceNumber = Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1))
                                    smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                                    Do
                                        slLine = MyFile.ReadLine
                                        slLine = UCase$(slLine)
                                        slWideOrbitForm = "1"
                                        ilPos1 = InStr(1, slLine, "STATION", vbBinaryCompare)
                                        If ilPos1 <= 0 Then
                                            ilPos1 = InStr(1, slLine, "PROPERTY", vbBinaryCompare)
                                            slWideOrbitForm = "2"
                                        End If
                                        If ilPos1 > 0 Then
                                            ilPos2 = InStr(1, slLine, "ACCOUNT", vbBinaryCompare)
                                            If ilPos2 > 0 Then
                                                Do
                                                    slLine = MyFile.ReadLine
                                                    slLine = UCase$(slLine)
                                                    If Len(Trim$(slLine)) > 0 Then
                                                        smCallLetters = Trim$(Mid(slLine, ilPos1, ilPos2 - ilPos1 - 1))
                                                        If Len(smCallLetters) > 0 Then
                                                            Exit Do
                                                        End If
                                                    End If
                                                Loop While Not MyFile.AtEndOfStream
                                                smCallLetters = mRemoveBlanks(smCallLetters)
                                                If Len(smCallLetters) > 0 Then
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    Loop While Not MyFile.AtEndOfStream
                                End If
                            End If
                        Else
                            blFound = False
                            ilPos1 = InStr(1, slLine, "INVOICE", vbBinaryCompare)
                            If (ilPos1 > 2) And (InStr(1, slLine, "OFFICIAL", vbBinaryCompare) <= 0) Then
                                If Trim$(Mid(slLine, ilPos1 - 2, 1)) <> "" Then
                                    blFound = True
                                End If
                            End If
                            If blFound Then
                                slPoss1CallLetters = ""
                                slPoss2CallLetters = ""
                                If Trim$(Left(slLine, 1)) <> "" Then
                                    smCallLetters = Trim$(Left(slLine, 10))
                                    smCallLetters = mRemoveBlanks(smCallLetters)
                                    slLine = MyFile.ReadLine    'Blank line
                                    If MyFile.AtEndOfStream Then
                                        Exit Do
                                    End If
                                    slLine = MyFile.ReadLine
                                    If MyFile.AtEndOfStream Then
                                        Exit Do
                                    End If
                                    slLine = UCase$(slLine)
                                Else
                                    slPoss1CallLetters = Left(Trim$(slLine), 4)
                                    If InStr(1, slLine, "/", vbBinaryCompare) > 0 Then
                                        slStr = Trim$(Mid(slLine, InStr(1, slLine, "/", vbBinaryCompare) + 1))
                                        slPoss2CallLetters = Left(slStr, 4)
                                    Else
                                        slPoss2CallLetters = ""
                                    End If
                                    slLine = MyFile.ReadLine    'Blank line
                                    If MyFile.AtEndOfStream Then
                                        Exit Do
                                    End If
                                    slLine = MyFile.ReadLine
                                    If MyFile.AtEndOfStream Then
                                        Exit Do
                                    End If
                                    slLine = UCase$(slLine)
                                    smCallLetters = Trim$(Left(Trim$(slLine), 10))
                                    smCallLetters = mRemoveBlanks(smCallLetters)
                                    'See if station, if not try slSvLine
                                    blFound = False
                                    slCallLetters = UCase$(smCallLetters)   'UCase$(Trim$(grdMatchedResult.TextMatrix(llMatchRow, MRSTNSTATIONINDEX)))
                                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                        If UCase$(Trim$(tgMVef(ilVef).sName)) = slCallLetters Then
                                            blFound = True
                                            Exit For
                                        End If
                                    Next ilVef
                                    If Not blFound Then
                                        slCallLetters = UCase$(slPoss1CallLetters)   'UCase$(Trim$(grdMatchedResult.TextMatrix(llMatchRow, MRSTNSTATIONINDEX)))
                                        If slPoss1CallLetters <> "" Then
                                            slPoss1CallLetters = ""
                                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                If Left(UCase$(Trim$(tgMVef(ilVef).sName)), 4) = slCallLetters Then
                                                    If Not blFound Then
                                                        blFound = True
                                                        slPoss1CallLetters = Trim$(tgMVef(ilVef).sName)
                                                    Else
                                                        slPoss1CallLetters = ""
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilVef
                                        End If
                                        If blFound And (slPoss1CallLetters <> "") Then
                                            smCallLetters = slPoss1CallLetters
                                        ElseIf (Not blFound) And (slPoss2CallLetters <> "") Then
                                            slCallLetters = UCase$(slPoss2CallLetters)   'UCase$(Trim$(grdMatchedResult.TextMatrix(llMatchRow, MRSTNSTATIONINDEX)))
                                            slPoss2CallLetters = ""
                                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                If Left(UCase$(Trim$(tgMVef(ilVef).sName)), 4) = slCallLetters Then
                                                    If Not blFound Then
                                                        blFound = True
                                                        slPoss2CallLetters = Trim$(tgMVef(ilVef).sName)
                                                    Else
                                                        slPoss2CallLetters = ""
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilVef
                                            If blFound And (slPoss2CallLetters <> "") Then
                                                smCallLetters = slPoss2CallLetters
                                            End If
                                        End If
                                    End If
                                End If
                                If smCallLetters <> "" Then
                                    'Remove blanks
                                    'Do
                                    '    ilPos2 = InStr(1, smCallLetters, " ", vbTextCompare)
                                    '    If ilPos2 <= 0 Then
                                    '        Exit Do
                                    '    End If
                                    '    smCallLetters = Left(smCallLetters, ilPos2 - 1) & Mid$(smCallLetters, ilPos2 + 1)
                                    'Loop
                                    smCallLetters = mRemoveBlanks(smCallLetters)
                                    If smCallLetters <> "" Then
                                        ilPos1 = InStr(1, slLine, "INVOICE ID:", vbBinaryCompare)
                                        If ilPos1 > 0 Then
                                            smSourceForm = "N"
                                            smInvoiceNumber = Trim$(Mid$(slLine, ilPos1 + 11))
                                            smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                                            Do
                                                slLine = MyFile.ReadLine
                                                slLine = UCase$(slLine)
                                                ilPos1 = InStr(1, slLine, "ORDER ID:", vbBinaryCompare)
                                                If ilPos1 > 0 Then
                                                    smContractNumber = Trim$(Mid$(slLine, ilPos1 + 10))
                                                    smContractNumber = mRemoveBlanks(smContractNumber)
                                                    Exit Do
                                                End If
                                            Loop While Not MyFile.AtEndOfStream
                                        Else
                                            smCallLetters = ""
                                        End If
                                    End If
                                End If
                            Else
                                ilPos1 = InStr(1, slLine, "OFFICIAL INVOICE", vbBinaryCompare)
                                If (ilPos1 > 0) Or ((InStr(1, slLine, "DETACH", vbBinaryCompare) > 0) And (InStr(1, slLine, "RETURN", vbBinaryCompare) > 0) And (InStr(1, slLine, "PAYMENT", vbBinaryCompare) > 0)) Then
                                    smSourceForm = "V"
                                    blFound = False
                                    For ilLoop = 0 To UBound(slPrevLines) Step 1
                                        slStr = slPrevLines(ilLoop)
                                        For ilLine = 1 To Len(slStr) Step 1
                                            slChar = Mid$(slStr, ilLine, 1)
                                            If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Then
                                                ilPos1 = ilLine
                                                blFound = True
                                                Exit For
                                            End If
                                        Next ilLine
                                        If blFound Then
                                            'smContractNumber = Trim$(Mid$(slStr, ilPos1, 15))
                                            'smContractNumber = mRemoveBlanks(smContractNumber)
                                            smInvoiceNumber = Trim$(Mid$(slStr, ilPos1, 15))
                                            smInvoiceNumber = mRemoveBlanks(smInvoiceNumber)
                                            Exit For
                                        End If
                                    Next ilLoop
                                    smCallLetters = ""
                                Else
                                    If Trim$(slLine) <> "" Then
                                        For ilLoop = 0 To UBound(slPrevLines) - 1 Step 1
                                            slPrevLines(ilLoop + 1) = slPrevLines(ilLoop)
                                        Next ilLoop
                                        slPrevLines(0) = slLine
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If smAdvertiserName = "" Then
                    If smSourceForm = "M" Then
                        ilPos1 = InStr(1, slLine, "ADVERTISER:", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smAdvertiserName = Trim$(Mid$(slLine, ilPos1 + 11))
                            smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                        End If
                    ElseIf smSourceForm = "W" Then
                        If slWideOrbitForm = "3" Then
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                If Len(Trim$(slLine)) > 0 Then
                                    ilPos1 = InStr(1, slLine, "ALT ORDER #", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        slAltOrderNumber = Trim$(Mid(slLine, ilPos1 + Len("ALT ORDER #")))
                                        slAltOrderNumber = mRemoveExtraBlanks(slAltOrderNumber)
                                        Exit Do
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                If Len(Trim$(slLine)) > 0 Then
                                    ilPos1 = InStr(1, slLine, "ADVERTISER", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        smAdvertiserName = Trim$(Mid(slLine, ilPos1 + Len("Advertiser")))
                                        smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                                        Exit Do
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                If Len(Trim$(slLine)) > 0 Then
                                    ilPos1 = InStr(1, slLine, "ESTIMATE #", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        smEstimateNumber = Trim$(Mid(slLine, ilPos1 + Len("ESTIMATE #")))
                                        smEstimateNumber = mRemoveExtraBlanks(smEstimateNumber)
                                        Exit Do
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                            If smEstimateNumber = "" Then
                                If slAltOrderNumber = "" Then
                                    smEstimateNumber = "Not Specified"
                                Else
                                    smEstimateNumber = slAltOrderNumber
                                End If
                            End If

                        Else
                            ilPos1 = InStr(1, slLine, "ADVERTISER", vbBinaryCompare)
                            If ilPos1 > 0 Then
                                slSvLine = slLine
                                ilPos2 = InStr(1, slSvLine, "PRODUCT", vbBinaryCompare)
                                Do
                                    slLine = MyFile.ReadLine
                                    slLine = UCase$(slLine)
                                    If Len(Trim$(slLine)) > 0 Then
                                        smAdvertiserName = Trim$(Mid(slLine, ilPos1, ilPos2 - ilPos1 - 1))
                                        smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                                        ilPos3 = InStr(1, slSvLine, "ESTIMATE NUMBER", vbBinaryCompare)
                                        If ilPos3 > 0 Then
                                            smEstimateNumber = Trim$(Mid$(slLine, ilPos3))
                                            smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                        End If
                                        If smAdvertiserName = "" Then
                                            slLine = MyFile.ReadLine
                                            slLine = UCase$(slLine)
                                            smAdvertiserName = Trim$(Mid(slLine, ilPos1, ilPos2 - ilPos1 - 1))
                                            smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                                        End If
                                        Exit Do
                                    End If
                                Loop While Not MyFile.AtEndOfStream
                                If smEstimateNumber = "" Then
                                    smEstimateNumber = "Not Specified"
                                End If
                                Do
                                    slLine = MyFile.ReadLine
                                    slLine = UCase$(slLine)
                                    If Len(Trim(slLine)) > 0 Then
                                        ilPos1 = InStr(1, slLine, "ORDER #", vbBinaryCompare)
                                        If ilPos1 > 0 Then
                                            ilPos2 = InStr(1, slLine, "ALT ORDER #", vbBinaryCompare)
                                            Do
                                                slLine = MyFile.ReadLine
                                                slLine = UCase$(slLine)
                                                If Len(Trim$(slLine)) > 0 Then
                                                    smContractNumber = Trim$(Mid(slLine, ilPos1, ilPos2 - ilPos1 - 1))
                                                    If Len(smContractNumber) > 0 Then
                                                        smContractNumber = mRemoveBlanks(smContractNumber)
                                                        slAltOrderNumber = Trim$(Mid(slLine, ilPos2))
                                                        If Len(slAltOrderNumber) > 0 Then
                                                            slAltOrderNumber = mRemoveBlanks(slAltOrderNumber)
                                                            If smEstimateNumber = "Not Specified" Then
                                                                smEstimateNumber = slAltOrderNumber
                                                            End If
                                                        End If
                                                        Exit Do
                                                    End If
                                                End If
                                            Loop While Not MyFile.AtEndOfStream
                                        End If
                                        If Len(smContractNumber) > 0 Then
                                            Exit Do
                                        End If
                                    End If
                                Loop While Not MyFile.AtEndOfStream
                                
                            End If
                        End If
                    ElseIf smSourceForm = "R" Then
                        ilPos1 = InStr(1, slLine, "C/O", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smAdvertiserName = Trim$(Left$(slLine, ilPos1 - 1))
                            smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                        End If
                    ElseIf smSourceForm = "N" Then
                        ilPos1 = InStr(1, slLine, "SPONSOR:", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            ilPos3 = 0
                            ilPos2 = InStr(1, slLine, "ESTIMATE", vbBinaryCompare)
                            If ilPos2 > 0 Then
                                ilPos3 = InStr(ilPos2 + 1, slLine, "EST", vbBinaryCompare)
                            End If
                            If (ilPos2 > 0) And (ilPos3 > 0) Then
                                smAdvertiserName = Trim$(Mid(slLine, ilPos1 + 8, ilPos2 - ilPos1 - 18))
                                smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                                ilPos3 = InStr(1, slLine, "ORDER#", vbBinaryCompare)
                                If ilPos3 > 0 Then
                                    smEstimateNumber = Trim$(Mid$(slLine, ilPos2 + 16, ilPos3 - ilPos2 - 17))
                                    smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                    smNetContractNumber = Trim$(Mid(slLine, ilPos3 + 6))
                                    smNetContractNumber = mRemoveBlanks(smNetContractNumber)
                                Else
                                    smEstimateNumber = Trim$(Mid$(slLine, ilPos2 + 16))
                                    smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                End If
                                If smEstimateNumber = "" Then
                                    smEstimateNumber = "Not Specified"
                                End If
                            Else
                                ilPos2 = InStr(1, slLine, "ESTIMATE", vbBinaryCompare)
                                ilPos4 = InStr(1, slLine, "#", vbBinaryCompare)
                                If (ilPos2 > 0) And (ilPos4 > 0) And (ilPos4 > ilPos2 + 7) Then
                                    smAdvertiserName = Trim$(Mid(slLine, ilPos1 + 8, ilPos2 - ilPos1 - 18))
                                    smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                                    ilPos3 = InStr(1, slLine, "ORDER#", vbBinaryCompare)
                                    If ilPos3 > 0 Then
                                        smEstimateNumber = Trim$(Mid$(slLine, ilPos4 + 1, ilPos3 - ilPos4 - 1))
                                        smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                        smNetContractNumber = Trim$(Mid(slLine, ilPos3 + 6))
                                        smNetContractNumber = mRemoveBlanks(smNetContractNumber)
                                    Else
                                        smEstimateNumber = Trim$(Mid$(slLine, ilPos4 + 1))
                                        smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                    End If
                                    If smEstimateNumber = "" Then
                                        smEstimateNumber = "Not Specified"
                                    End If
                                End If
                            End If
                        End If
                    ElseIf smSourceForm = "V" Then
                        ilPos1 = InStr(1, slLine, "FOR:", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smAdvertiserName = Trim$(Mid$(slLine, ilPos1 + 4))
                            smAdvertiserName = mRemoveExtraBlanks(smAdvertiserName)
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                If Len(Trim(slLine)) > 0 Then
                                    ilPos1 = InStr(1, slLine, "EST. NUMBER:", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        smEstimateNumber = Trim$(Mid$(slLine, ilPos1 + 12))
                                        smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                                        If smEstimateNumber = "" Then
                                            smEstimateNumber = "Not Specified"
                                        End If
                                    End If
                                    ilPos1 = InStr(1, slLine, "DESCRIPTION:", vbBinaryCompare)
                                    If ilPos1 > 0 Then
                                        ilPos2 = InStr(1, slLine, "ORDER #", vbBinaryCompare)
                                        If ilPos2 > 0 Then
                                            smNetContractNumber = Trim$(Mid(slLine, ilPos2 + 7))
                                            smNetContractNumber = mRemoveBlanks(smNetContractNumber)
                                        Else
                                            ilPos2 = InStr(1, slLine, "ORD#", vbBinaryCompare)
                                            If ilPos2 > 0 Then
                                                smNetContractNumber = Trim$(Mid(slLine, ilPos2 + 4))
                                                smNetContractNumber = mRemoveBlanks(smNetContractNumber)
                                            End If
                                        End If
                                        If (smEstimateNumber = "Not Specified") Or (Trim$(smEstimateNumber) = "") Then
                                            ilPos1 = InStr(1, slLine, "EST#", vbBinaryCompare)
                                            If ilPos1 > 0 Then
                                                smEstimateNumber = Trim$(Mid$(slLine, ilPos1 + 4, ilPos2 - ilPos1 - 5))
                                                If smEstimateNumber = "" Then
                                                    smEstimateNumber = "Not Specified"
                                                End If
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                        End If
                    ElseIf (smSourceForm = "WE") Or (smSourceForm = "ME") Then
                        If (InStr(1, slLine, "31;", vbTextCompare) > 0) And (slPDForEDI = "E") Then
                            gParseItemFields slLine, ";", slFields()
                            For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
                                slFields(ilCol + 1) = slFields(ilCol)
                            Next ilCol
                            slFields(0) = ""
                            smAdvertiserName = Trim$(slFields(4))
                            smEstimateNumber = mRemoveBlanks(slFields(8))
                            If smEstimateNumber = "" Then
                                smEstimateNumber = "Not Specified"
                            End If
                            smInvoiceNumber = mRemoveBlanks(slFields(9))
                            smNetContractNumber = mRemoveBlanks(slFields(22))
                            smContractNumber = mRemoveBlanks(slFields(23))
                        End If
                    End If
                ElseIf smEstimateNumber = "" Then
                    If smSourceForm = "M" Then
                        ilPos1 = InStr(1, slLine, "ESTIMATE #:", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            smEstimateNumber = Trim$(Mid$(slLine, ilPos1 + 11))
                            smEstimateNumber = mRemoveBlanks(smEstimateNumber)
                            If smEstimateNumber = "" Then
                                smEstimateNumber = "Not Specified"
                            End If
                            'ilPos1 = InStrRev(smEstimateNumber, " ", -1, vbBinaryCompare)
                            'If ilPos1 > 0 Then
                            '    smEstimateNumber = Mid$(smEstimateNumber, ilPos1 + 1)
                            'End If
                        End If
                    End If
                Else
                    'Get spots Date, Time and ISCI
                    If smSourceForm = "M" Then
                        ilPos1 = InStr(1, slLine, "DAY", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                            If ilPos1 > 0 Then
                                ilPos2 = InStr(1, slLine, "TIME", vbBinaryCompare)
                                ilPos3 = InStr(1, slLine, "PRODUCT", vbBinaryCompare)
                                ilPos4 = InStr(1, slLine, "ISCI", vbBinaryCompare)
                                ilPos5 = InStr(1, slLine, "RATE", vbBinaryCompare)
                                slMarketronForm = "1"
                                ilPos6 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                If (ilPos6 > 0) And (ilPos6 > ilPos2) And (ilPos6 < ilPos3) Then
                                    slMarketronForm = "2"
                                End If
                                blEndIfInvoice = False
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                Do
                                    If Len(Trim$(slLine)) > 0 Then
                                        If (Asc(slLine) >= 32) Then
                                            blMarketronSkipRead = False
                                            If slMarketronForm <> "2" Then
                                                ilPos6 = InStr(1, slLine, "LENGTH:", vbBinaryCompare)
                                            Else
                                                blMarketronSkipRead = True
                                            End If
                                            If ilPos6 > 0 Then
                                                ilLen = Val(Trim$(Mid$(slLine, ilPos6 + 7)))
                                                'Build daypart info
                                                ilPos7 = InStr(1, slLine, ".", vbBinaryCompare)
                                                ilUpperStart = UBound(tmImportSpotInfo)
                                                If (ilPos7 > 0) And (slMarketronForm <> "2") Then
                                                    ilPos7 = ilPos7 + 3
                                                    tmImportSpotInfo(ilUpperStart).sDPDays = Trim$(Mid$(slLine, ilPos7, ilPos6 - ilPos7 - 1))

                                                    If (InStr(1, slLine, "AM", vbBinaryCompare) > 0) And (InStr(1, slLine, "PM", vbBinaryCompare) > 0) Then
                                                        If InStr(1, slLine, "AM", vbBinaryCompare) < InStr(1, slLine, "PM", vbBinaryCompare) Then
                                                            ilPos7 = InStr(1, slLine, "AM", vbBinaryCompare)
                                                        Else
                                                            ilPos7 = InStr(1, slLine, "PM", vbBinaryCompare)
                                                        End If
                                                    ElseIf (InStr(ilPos2, slLine, "AM", vbBinaryCompare) > 0) Then
                                                        ilPos7 = InStr(1, slLine, "AM", vbBinaryCompare)
                                                    Else
                                                        ilPos7 = InStr(1, slLine, "PM", vbBinaryCompare)
                                                    End If

                                                    ilPos6 = InStrRev(slLine, " ", ilPos7, vbBinaryCompare)
                                                    tmImportSpotInfo(ilUpperStart).lDPStartTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos6, ilPos7 - ilPos6 + 1)), "h:mm:ssAM/PM"), False)
                                                    ilPos6 = ilPos7 + 3
                                                    ilPos7 = InStr(ilPos6, slLine, " ", vbBinaryCompare)
                                                    tmImportSpotInfo(ilUpperStart).lDPEndTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos6, ilPos7 - ilPos6 - 1)), "h:mm:ssAM/PM"), True)
                                                Else
                                                    tmImportSpotInfo(ilUpperStart).lDPStartTime = -1
                                                    tmImportSpotInfo(ilUpperStart).lDPEndTime = -1
                                                    tmImportSpotInfo(ilUpperStart).sDPDays = ""
                                                End If
                                                'Get air info
                                                Do
                                                    If Not blMarketronSkipRead Then
                                                        slLine = MyFile.ReadLine
                                                        slLine = UCase$(slLine)
                                                    Else
                                                        blMarketronSkipRead = False
                                                    End If
                                                    If Len(Trim$(slLine)) > 0 Then
                                                        If InStr(1, slLine, "LENGTH:", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "INVOICE TOTAL", vbBinaryCompare) > 0 Then
                                                            blEndIfInvoice = True
                                                            Exit Do
                                                        End If
                                                        blLine = False
                                                        Select Case Left$(slLine, 3)
                                                            Case "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"
                                                                blLine = True
                                                        End Select
                                                        If blLine = True Then
                                                            ilUpper = UBound(tmImportSpotInfo)
                                                            tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1)))
                                                            If (InStr(ilPos2, slLine, "A", vbBinaryCompare) > 0) And (InStr(ilPos2, slLine, "P", vbBinaryCompare) > 0) Then
                                                                If InStr(ilPos2, slLine, "A", vbBinaryCompare) < InStr(ilPos2, slLine, "P", vbBinaryCompare) Then
                                                                    tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, InStr(ilPos2, slLine, "A", vbBinaryCompare) - ilPos2 + 1)), "h:mm:ssAM/PM"), False)
                                                                Else
                                                                    tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, InStr(ilPos2, slLine, "P", vbBinaryCompare) - ilPos2 + 1)), "h:mm:ssAM/PM"), False)
                                                                End If
                                                            ElseIf (InStr(ilPos2, slLine, "A", vbBinaryCompare) > 0) Then
                                                                tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, InStr(ilPos2, slLine, "A", vbBinaryCompare) - ilPos2 + 1)), "h:mm:ssAM/PM"), False)
                                                            Else
                                                                tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, InStr(ilPos2, slLine, "P", vbBinaryCompare) - ilPos2 + 1)), "h:mm:ssAM/PM"), False)
                                                            End If
                                                            If slMarketronForm <> "2" Then
                                                                tmImportSpotInfo(ilUpper).iLen = ilLen
                                                            Else
                                                                tmImportSpotInfo(ilUpper).iLen = Trim$(Mid$(slLine, ilPos6, ilPos3 - ilPos6))
                                                            End If
                                                            tmImportSpotInfo(ilUpper).lRate = -1
                                                            ilRatePos = InStr(ilPos2, slLine, "$", vbBinaryCompare)
                                                            If ilRatePos > 0 Then
                                                                ilDecPointPos = InStr(ilRatePos, slLine, ".", vbBinaryCompare)
                                                                If ilDecPointPos > 0 Then
                                                                    tmImportSpotInfo(ilUpper).lRate = gStrDecToLong(Mid$(slLine, ilRatePos + 1, ilDecPointPos - ilRatePos + 2), 2)
                                                                End If
                                                            End If
                                                            If tmImportSpotInfo(ilUpper).iLen > 0 Then
                                                                tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(Trim$(Mid$(slLine, ilPos4, ilPos5 - ilPos4 - 1)))
                                                                tmImportSpotInfo(ilUpper).lDPStartTime = tmImportSpotInfo(ilUpperStart).lDPStartTime
                                                                tmImportSpotInfo(ilUpper).lDPEndTime = tmImportSpotInfo(ilUpperStart).lDPEndTime
                                                                tmImportSpotInfo(ilUpper).sDPDays = tmImportSpotInfo(ilUpperStart).sDPDays
                                                                tmImportSpotInfo(ilUpper).bMatched = False
                                                                ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                                                If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                                                    llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                                                End If
                                                                lmInvStartDate = llStartDate
                                                                If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                                                    llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                                                End If
                                                            End If
                                                        Else
                                                            If InStr(1, slLine, "DAY", vbBinaryCompare) > 0 Then
                                                                If InStr(1, slLine, "DATE", vbBinaryCompare) > 0 Then
                                                                    ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                                                                    ilPos2 = InStr(1, slLine, "TIME", vbBinaryCompare)
                                                                    ilPos3 = InStr(1, slLine, "PRODUCT", vbBinaryCompare)
                                                                    ilPos4 = InStr(1, slLine, "ISCI", vbBinaryCompare)
                                                                    ilPos5 = InStr(1, slLine, "RATE", vbBinaryCompare)
                                                                    If slMarketronForm = "2" Then
                                                                        ilPos6 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Loop While Not MyFile.AtEndOfStream
                                                If blEndIfInvoice Then
                                                    Exit Do
                                                End If
                                            Else
                                                slLine = MyFile.ReadLine
                                                slLine = UCase$(slLine)
                                            End If
                                        Else
                                            slLine = MyFile.ReadLine
                                            slLine = UCase$(slLine)
                                        End If
                                    Else
                                        slLine = MyFile.ReadLine
                                        slLine = UCase$(slLine)
                                    End If
                                Loop While Not MyFile.AtEndOfStream
                                'Match up info
                                If llStartDate <> 99999999 Then
                                    ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                                    blInvoiceProcessed = True
                                    If (ilRet) Then
                                        If (Not blAnyFoundSetToFalse) Then
                                            blAnyFound = True
                                        End If
                                    Else
                                        blAnyFound = False
                                        blAnyFoundSetToFalse = True
                                    End If
                                End If
                                'Look for next Invoice within file
                                mClearValues llStartDate, llEndDate
                            End If
                        End If
                    ElseIf smSourceForm = "W" Then
                        blLineHeaderFound = False
                        Do
                            If Len(Trim$(slLine)) > 0 Then
                                If (Asc(slLine) >= 32) Then
                                    If InStr(1, slLine, "SPOT TOTALS", vbBinaryCompare) > 0 Then
                                        blLineHeaderFound = False
                                        Exit Do
                                    End If
                                    If InStr(1, slLine, "TOTAL SPOTS", vbBinaryCompare) > 0 Then
                                        blLineHeaderFound = False
                                        Exit Do
                                    End If
                                    If InStr(1, slLine, "GROSS TOTAL", vbBinaryCompare) > 0 Then
                                        blLineHeaderFound = False
                                        Exit Do
                                    End If
                                    If InStr(1, slLine, "NET TOTAL", vbBinaryCompare) > 0 Then
                                        blLineHeaderFound = False
                                        Exit Do
                                    End If
                                    If InStr(1, slLine, "LINE", vbBinaryCompare) > 0 Then
                                        If InStr(1, slLine, "MTWTFSS", vbBinaryCompare) > 0 Then
                                            ilLnPos1 = InStr(1, slLine, "START/END", vbBinaryCompare)
                                            If ilLnPos1 > 0 Then
                                                blLineHeaderFound = True
                                                ilLnPos2 = InStr(1, slLine, "MTWTFSS", vbBinaryCompare)
                                                ilLnPos3 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                                ilLnPos4 = InStr(1, slLine, "WEEK", vbBinaryCompare)
                                                ilLnPos5 = InStr(ilLnPos1, slLine, "-", vbBinaryCompare)
                                                ilLnPos6 = InStr(ilLnPos1, slLine, "RATE", vbBinaryCompare)
                                                ilUpperStart = UBound(tmImportSpotInfo)
                                                tmImportSpotInfo(ilUpperStart).lDPStartTime = -1
                                                tmImportSpotInfo(ilUpperStart).lDPEndTime = -1
                                                tmImportSpotInfo(ilUpperStart).sDPDays = ""
                                            End If
                                        End If
                                    End If
                                    If blLineHeaderFound Then
                                        'Line Information
                                        slStr = Trim$(Left(slLine, 4))
                                        blLine = False
                                        If Len(slStr) > 0 Then
                                            blLine = True
                                            For ilLine = 1 To Len(slStr) Step 1
                                                slChar = Mid$(slStr, ilLine, 1)
                                                If (Asc(slChar) < Asc("0")) Or (Asc(slChar) > Asc("9")) Then
                                                    blLine = False
                                                    Exit For
                                                End If
                                            Next ilLine
                                            If blLine Then
                                                If (InStr(1, slLine, "VARIOUS", vbBinaryCompare) <= 0) And (InStr(1, slLine, ":", vbBinaryCompare) <= 0) Then
                                                    blLine = False
                                                End If
                                            End If
                                        End If
                                        If blLine And (ilLnPos1 > 0) Then
                                            ilUpperStart = UBound(tmImportSpotInfo)
                                            If InStr(ilLnPos1, slLine, "VARIOUS", vbBinaryCompare) <= 0 Then
                                                ilLnPos5 = InStr(ilLnPos1, slLine, "-", vbBinaryCompare)
                                                If InStr(ilLnPos1, slLine, "XM", vbBinaryCompare) > 0 Then
                                                    slLine = Replace(slLine, " XM ", " AM ", 1, 2, vbBinaryCompare)
                                                End If
                                                If (ilLnPos5 > ilLnPos1) And (ilLnPos2 > (ilLnPos5 + 1)) Then
                                                    tmImportSpotInfo(ilUpperStart).lDPStartTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilLnPos1, ilLnPos5 - ilLnPos1)), "h:mm:ssAM/PM"), False)
                                                    tmImportSpotInfo(ilUpperStart).lDPEndTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilLnPos5 + 1, ilLnPos2 - ilLnPos5 - 1)), "h:mm:ssAM/PM"), True)
                                                    tmImportSpotInfo(ilUpperStart).sDPDays = Trim$(Mid$(slLine, ilLnPos2, ilLnPos3 - ilLnPos2 - 1))
                                                End If
                                            End If
                                            ilPos1 = InStr(ilLnPos3, slLine, ":", vbBinaryCompare)
                                            slStr = Trim$(Mid$(slLine, ilPos1 - 1, 4))
                                            If InStr(1, slStr, ":", vbBinaryCompare) = 1 Then
                                                tmImportSpotInfo(ilUpperStart).iLen = gLengthToLong("00:00" & slStr)
                                            ElseIf InStr(1, slStr, ":", vbBinaryCompare) = 2 Then
                                                tmImportSpotInfo(ilUpperStart).iLen = gLengthToLong("00:0" & slStr)
                                            Else
                                                tmImportSpotInfo(ilUpperStart).iLen = gLengthToLong("00:00:" & slStr)
                                            End If
                                            ilLen = tmImportSpotInfo(ilUpperStart).iLen
                                        End If
                                        ilPos1 = InStr(1, slLine, "SPOTS:", vbBinaryCompare)
                                        If ilPos1 > 0 Then
                                            ilSpotPos = InStr(1, slLine, "CH", vbBinaryCompare) - 1
                                            ilPos1 = InStr(1, slLine, "AIR DATE", vbBinaryCompare)
                                            If ilPos1 > 0 Then
                                                'ilPos2 = InStr(1, slLine, "AIR TIME", vbBinaryCompare)
                                                ilPos2 = ilPos1 + 9
                                                ilPos3 = InStr(1, slLine, "DESCRIPTION", vbBinaryCompare)
                                                ilPos4 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                                'ilPos5 = InStr(1, slLine, "AD-ID", vbBinaryCompare)
                                                ilPos6 = InStr(1, slLine, "RATE", vbBinaryCompare)
                                                'ilPos7 = ilPos1 - 8
                                                Do
                                                    slLine = MyFile.ReadLine
                                                    slLine = UCase$(slLine)
                                                    If (Len(Trim$(slLine)) > 0) And (Trim$(Mid$(slLine, 1, ilSpotPos)) <> "") Then
                                                        If InStr(1, slLine, "SPOT TOTALS", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "TOTAL SPOTS", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "GROSS TOTAL", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "NET TOTAL", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "WEEKS:", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        If InStr(1, slLine, "LINE", vbBinaryCompare) > 0 Then
                                                            Exit Do
                                                        End If
                                                        slStr = Trim$(Left(slLine, 4))
                                                        blLine = False
                                                        If Len(slStr) > 0 Then
                                                            blLine = True
                                                            For ilLine = 1 To Len(slStr) Step 1
                                                                slChar = Mid$(slStr, ilLine, 1)
                                                                If (Asc(slChar) < Asc("0")) Or (Asc(slChar) > Asc("9")) Then
                                                                    blLine = False
                                                                End If
                                                            Next ilLine
                                                        End If
                                                        If blLine Then
                                                            Exit Do
                                                        End If
                                                        'Check if actual spot
                                                        ilPos7 = ilPos1 - 1
                                                        Do
                                                            slChar = Mid$(slLine, ilPos7, 1)
                                                            If Trim$(slChar) <> "" Then
                                                                Exit Do
                                                            End If
                                                            ilPos7 = ilPos7 - 1
                                                        Loop
                                                        If ilPos7 > 1 Then
                                                            ilPos7 = ilPos7 - 1
                                                            slDay = Trim$(Mid$(slLine, ilPos7, 3))
                                                            If (slDay = "M") Or (slDay = "TU") Or (slDay = "W") Or (slDay = "TH") Or (slDay = "F") Or (slDay = "SA") Or (slDay = "SU") Then
                                                                ilUpper = UBound(tmImportSpotInfo)
                                                                tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1)))
                                                                ilAMPMPos = InStr(ilPos2, slLine, "M", vbTextCompare)
                                                                If ilAMPMPos > 0 Then
                                                                    tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, ilAMPMPos - ilPos2 + 1)), "h:mm:ssAM/PM"), False)
                                                                Else
                                                                    tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, ilPos3 - ilPos2 - 2)), "h:mm:ssAM/PM"), False)
                                                                End If
                                                                ilPos5 = ilPos4
                                                                Do
                                                                    slChar = Mid$(slLine, ilPos5, 1)
                                                                    If Trim$(slChar) = ":" Then
                                                                        Exit Do
                                                                    End If
                                                                    ilPos5 = ilPos5 + 1
                                                                Loop
                                                                Do
                                                                    slChar = Mid$(slLine, ilPos5, 1)
                                                                    If Trim$(slChar) = "" Then
                                                                        Exit Do
                                                                    End If
                                                                    ilPos5 = ilPos5 + 1
                                                                Loop
                                                                'ilPos4 = ilPos4 + 1
                                                                'ilPos5 = ilPos4 + 1
                                                                'Do
                                                                '    slChar = Mid$(slLine, ilPos5, 1)
                                                                '    If Trim$(slChar) = "" Then
                                                                '        Exit Do
                                                                '    End If
                                                                '    ilPos5 = ilPos5 + 1
                                                                'Loop
                                                                'If Mid(slLine, ilPos4, 1) = ":" Then
                                                                '    tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:00" & Trim$(Mid$(slLine, ilPos4, ilPos5 - ilPos4)))
                                                                'Else
                                                                '    tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:" & Trim$(Mid$(slLine, ilPos4, ilPos5 - ilPos4 - 1)))
                                                                'End If
                                                                tmImportSpotInfo(ilUpper).iLen = ilLen
                                                                If ilLen > 0 Then
                                                                    'Remove extra blanks
                                                                    'tmImportSpotInfo(ilUpper).sISCI = Trim$(Mid$(slLine, ilPos5, ilPos6 - ilPos5 - 1))
                                                                    If InStr(ilPos5, slLine, "$", vbBinaryCompare) > 0 Then
                                                                        slStr = Trim$(Mid$(slLine, ilPos5, InStr(ilPos5, slLine, "$", vbBinaryCompare) - ilPos5))
                                                                    Else
                                                                        slStr = Trim$(Mid$(slLine, ilPos5, ilPos6 - ilPos5))
                                                                    End If
                                                                    tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(slStr)
                                                                    tmImportSpotInfo(ilUpper).lDPStartTime = tmImportSpotInfo(ilUpperStart).lDPStartTime
                                                                    tmImportSpotInfo(ilUpper).lDPEndTime = tmImportSpotInfo(ilUpperStart).lDPEndTime
                                                                    tmImportSpotInfo(ilUpper).sDPDays = tmImportSpotInfo(ilUpperStart).sDPDays
                                                                    tmImportSpotInfo(ilUpper).bMatched = False
                                                                    tmImportSpotInfo(ilUpper).lRate = -1
                                                                    ilRatePos = InStr(ilPos5, slLine, "$", vbBinaryCompare)
                                                                    If ilRatePos > 0 Then
                                                                        ilDecPointPos = InStr(ilRatePos, slLine, ".", vbBinaryCompare)
                                                                        If ilDecPointPos > 0 Then
                                                                            tmImportSpotInfo(ilUpper).lRate = gStrDecToLong(Mid$(slLine, ilRatePos + 1, ilDecPointPos - ilRatePos + 2), 2)
                                                                        End If
                                                                    End If
                                                                    ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                                                    If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                                                        llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                                                    End If
                                                                    lmInvStartDate = llStartDate
                                                                    If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                                                        llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Loop While Not MyFile.AtEndOfStream
                                            Else
                                                slLine = MyFile.ReadLine
                                                slLine = UCase$(slLine)
                                            End If
                                        Else
                                            slLine = MyFile.ReadLine
                                            slLine = UCase$(slLine)
                                        End If
                                    Else
                                        slLine = MyFile.ReadLine
                                        slLine = UCase$(slLine)
                                    End If
                                Else
                                    slLine = MyFile.ReadLine
                                    slLine = UCase$(slLine)
                                End If
                            Else
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                            End If
                        Loop While Not MyFile.AtEndOfStream
                        'Match up info
                        If llStartDate <> 99999999 Then
                            ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                            blInvoiceProcessed = True
                            If (ilRet) Then
                                If (Not blAnyFoundSetToFalse) Then
                                    blAnyFound = True
                                End If
                            Else
                                blAnyFound = False
                                blAnyFoundSetToFalse = True
                            End If
                        End If
                        'Look for next Invoice within file
                        mClearValues llStartDate, llEndDate
                    ElseIf smSourceForm = "R" Then
                        ilPos1 = InStr(1, slLine, "CART", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                            If ilPos1 > 0 Then
                                ilPos2 = InStr(1, slLine, "TIME", vbBinaryCompare)
                                ilPos3 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                ilPos4 = InStr(1, slLine, "DESCRIPTION", vbBinaryCompare)
                                ilPos6 = InStr(1, slLine, "RATE", vbBinaryCompare)
                                Do
                                    slLine = MyFile.ReadLine
                                    slLine = UCase$(slLine)
                                    ilPos5 = InStr(1, slLine, "SUB", vbBinaryCompare)
                                    If ilPos5 > 0 Then
                                        ilPos5 = InStr(1, slLine, "TOTAL", vbBinaryCompare)
                                        If ilPos5 > 0 Then
                                            Exit Do
                                        End If
                                    End If
                                    If (InStr(1, slLine, "DATE", vbBinaryCompare) > 0) And (InStr(1, slLine, "TIME", vbBinaryCompare) > 0) And (InStr(1, slLine, "LENGTH", vbBinaryCompare) > 0) And (InStr(1, slLine, "DESCRIPTION", vbBinaryCompare) > 0) Then
                                        ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                                        ilPos2 = InStr(1, slLine, "TIME", vbBinaryCompare)
                                        ilPos3 = InStr(1, slLine, "LENGTH", vbBinaryCompare)
                                        ilPos4 = InStr(1, slLine, "DESCRIPTION", vbBinaryCompare)
                                        ilPos6 = InStr(1, slLine, "RATE", vbBinaryCompare)
                                        slLine = MyFile.ReadLine
                                        slLine = UCase$(slLine)
                                    End If
                                    If (Len(Trim$(slLine)) > 0) And (InStr(1, slLine, "PRINTED", vbBinaryCompare) <= 0) Then
                                        If (Trim$(Mid$(slLine, ilPos1, 1)) <> "") And (Asc(slLine) >= 32) Then
                                            If (Len(Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1))) >= 9) And (Len(Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1))) > 5) Then
                                                ilUpper = UBound(tmImportSpotInfo)
                                                slStr = Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1))
                                                If Len(slStr) = 9 Then
                                                    ilMonth = InStr(ilPos1, slLine, "/", vbBinaryCompare)
                                                    ilYear = InStr(ilMonth + 1, slLine, "/", vbBinaryCompare)
                                                    If (Val(Mid$(slLine, ilPos1, ilMonth - ilPos1)) = 12) And (Month(Now) = 1) Then
                                                        slStr = Mid$(slLine, ilPos1, ilYear - ilPos1 + 1) & (Val(Year(Now)) - 1)
                                                    Else
                                                        slStr = Mid$(slLine, ilPos1, ilYear - ilPos1 + 1) & Year(Now)
                                                    End If
                                                    tmImportSpotInfo(ilUpper).lAirDate = gDateValue(slStr)
                                                Else
                                                    tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Trim$(Mid$(slLine, ilPos1, ilPos2 - ilPos1 - 1)))
                                                End If
                                                tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, ilPos3 - ilPos2 - 1)), "h:mm:ssAM/PM"), False)
                                                tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:" & Trim$(Mid$(slLine, ilPos3, ilPos4 - ilPos3 - 1)))
                                                If tmImportSpotInfo(ilUpper).iLen > 0 Then
                                                    tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(Trim$(Mid$(slLine, ilPos4, ilPos6 - ilPos4)))
                                                    tmImportSpotInfo(ilUpper).lDPStartTime = -1
                                                    tmImportSpotInfo(ilUpper).lDPEndTime = -1
                                                    tmImportSpotInfo(ilUpper).sDPDays = ""
                                                    tmImportSpotInfo(ilUpper).bMatched = False
                                                    tmImportSpotInfo(ilUpper).lRate = -1
                                                    ilRatePos = ilPos6 - 5
                                                    If ilRatePos > 0 Then
                                                        ilDecPointPos = InStr(ilRatePos, slLine, ".", vbBinaryCompare)
                                                        If ilDecPointPos > 0 Then
                                                            tmImportSpotInfo(ilUpper).lRate = gStrDecToLong(Mid$(slLine, ilRatePos + 1, ilDecPointPos - ilRatePos + 2), 2)
                                                        End If
                                                    End If
                                                    ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                                    If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                                        llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                                    End If
                                                    If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                                        llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Loop While Not MyFile.AtEndOfStream
                                'Match up info
                                If llStartDate <> 99999999 Then
                                    ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                                    blInvoiceProcessed = True
                                    If (ilRet) Then
                                        If (Not blAnyFoundSetToFalse) Then
                                            blAnyFound = True
                                        End If
                                    Else
                                        blAnyFound = False
                                        blAnyFoundSetToFalse = True
                                    End If
                                End If
                                'Look for next Invoice within file
                                mClearValues llStartDate, llEndDate
                            End If
                        End If
                    ElseIf smSourceForm = "N" Then
                        ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                        If ilPos1 > 0 Then
                            ilPos7 = InStr(1, slLine, "CODEID", vbBinaryCompare)
                            Do
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                                ilPos5 = InStr(1, slLine, "AMOUNT DUE:", vbBinaryCompare)
                                If ilPos5 > 0 Then
                                    Exit Do
                                End If
                                If Len(Trim$(slLine)) > 0 Then
                                    If (Asc(slLine) >= 32) Then
                                        ilPos6 = InStr(1, slLine, "SPOT", vbBinaryCompare)
                                        If ilPos6 <= 0 Then
                                            ilPos6 = InStr(1, slLine, "ADDED", vbBinaryCompare)
                                            If ilPos6 > 0 Then
                                                If InStr(1, slLine, "VALUE", vbBinaryCompare) <= 0 Then
                                                    ilPos6 = 0
                                                End If
                                            Else
                                                ilPos6 = InStr(1, slLine, "BONUS", vbBinaryCompare)
                                            End If
                                        End If
                                        If ilPos6 > 0 Then
                                            ilUpper = UBound(tmImportSpotInfo)
                                            ilPos2 = ilPos1 + 1
                                            Do
                                                slChar = Mid$(slLine, ilPos2, 1)
                                                If Trim$(slChar) = "" Then
                                                    Exit Do
                                                End If
                                                ilPos2 = ilPos2 + 1
                                            Loop
                                            tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Trim$(Mid$(slLine, 1, ilPos2)))
                                            ilPos2 = ilPos2 + 1
                                            ilPos3 = ilPos2 + 1
                                            Do
                                                slChar = Mid$(slLine, ilPos3, 1)
                                                If Trim$(slChar) = "M" Then
                                                    Exit Do
                                                End If
                                                ilPos3 = ilPos3 + 1
                                            Loop
                                            ilPos3 = ilPos3 + 1
                                            tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilPos2, ilPos3 - ilPos2)), "h:mm:ssAM/PM"), False)
                                            slStr = Trim$(Mid$(slLine, ilPos3, ilPos6 - ilPos3 - 1))
                                            If Len(slStr) = 3 Then
                                                tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:00" & Trim$(Mid$(slLine, ilPos3, ilPos6 - ilPos3 - 1)))
                                            ElseIf Len(slStr) = 4 Then
                                                tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:0" & Trim$(Mid$(slLine, ilPos3, ilPos6 - ilPos3 - 1)))
                                            Else
                                                tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:" & Trim$(Mid$(slLine, ilPos3, ilPos6 - ilPos3 - 1)))
                                            End If
                                            ilPos6 = InStr(1, slLine, "[", vbBinaryCompare)
                                            If (ilPos6 > 0) And (ilPos7 > 0) Then
                                                tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(Trim$(Mid$(slLine, ilPos7, ilPos6 - ilPos7 - 1)))
                                            Else
                                                tmImportSpotInfo(ilUpper).sISCI = ""
                                            End If
                                            tmImportSpotInfo(ilUpper).lDPStartTime = -1
                                            tmImportSpotInfo(ilUpper).lDPEndTime = -1
                                            tmImportSpotInfo(ilUpper).sDPDays = ""
                                            tmImportSpotInfo(ilUpper).bMatched = False
                                            tmImportSpotInfo(ilUpper).lRate = -1
                                            If tmImportSpotInfo(ilUpper).iLen > 0 Then
                                                ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                                If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                                    llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                                End If
                                                If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                                    llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                                End If
                                            End If
                                        Else
                                            If (InStr(1, slLine, "DATE", vbBinaryCompare) > 0) And (InStr(1, slLine, "ISCI CODE", vbBinaryCompare) > 0) Then
                                                ilPos1 = InStr(1, slLine, "DATE", vbBinaryCompare)
                                                ilPos7 = InStr(1, slLine, "ISCI CODE", vbBinaryCompare)
                                            End If
                                        End If
                                    End If
                                End If
                            Loop While Not MyFile.AtEndOfStream
                            'Match up info
                            If llStartDate <> 99999999 Then
                                ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                                blInvoiceProcessed = True
                                If (ilRet) Then
                                    If (Not blAnyFoundSetToFalse) Then
                                        blAnyFound = True
                                    End If
                                Else
                                    blAnyFound = False
                                    blAnyFoundSetToFalse = True
                                End If
                            End If
                            'Look for next Invoice within file
                            mClearValues llStartDate, llEndDate
                        End If
                    ElseIf smSourceForm = "V" Then
                        slISCI = ""
                        Do
                            If Len(Trim$(slLine)) > 0 Then
                                If (Asc(slLine) >= 32) Then
                                If InStr(1, slLine, "TOTAL DUE", vbBinaryCompare) > 0 Then
                                        Exit Do
                                    End If
                                    If InStr(1, slLine, "QUANTITY", vbBinaryCompare) > 0 Then
                                        Exit Do
                                    End If
                                    ilLnPos1 = InStr(1, slLine, "ISCI CODE:", vbBinaryCompare)
                                    If ilLnPos1 > 0 Then
                                        slISCI = Trim$(Mid$(slLine, ilLnPos1 + 10, 20))
                                    End If
                                    ilLnPos1 = InStr(1, slLine, " AM", vbBinaryCompare)
                                    ilLnPos2 = InStr(1, slLine, " PM", vbBinaryCompare)
                                    ilLnPos3 = InStr(1, slLine, "$", vbBinaryCompare)
                                    ilLnPos4 = InStr(1, slLine, ":", vbBinaryCompare)
                                    If ((ilLnPos1 > 0) Or (ilLnPos2 > 0)) And (ilLnPos3 > 0) And (ilLnPos4 > 0) Then
                                        ilUpper = UBound(tmImportSpotInfo)
                                        tmImportSpotInfo(ilUpper).lDPStartTime = -1
                                        tmImportSpotInfo(ilUpper).lDPEndTime = -1
                                        tmImportSpotInfo(ilUpper).sDPDays = ""
                                        tmImportSpotInfo(ilUpper).bMatched = False
                                        tmImportSpotInfo(ilUpper).lRate = -1
                                        ilRatePos = InStr(1, slLine, "$", vbBinaryCompare)
                                        If ilRatePos > 0 Then
                                            ilDecPointPos = InStr(ilRatePos, slLine, ".", vbBinaryCompare)
                                            If ilDecPointPos > 0 Then
                                                tmImportSpotInfo(ilUpper).lRate = gStrDecToLong(Mid$(slLine, ilRatePos + 1, ilDecPointPos - ilRatePos + 2), 2)
                                            End If
                                        End If
                                        tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(slISCI)
                                        slLine = Trim$(slLine)
                                        ilLnPos5 = InStr(1, slLine, " ", vbBinaryCompare)
                                        tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Left$(slLine, ilLnPos5 - 1))
                                        If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                            llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                        End If
                                        If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                            llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                        End If
                                        ilLnPos5 = InStr(1, slLine, ":", vbBinaryCompare)
                                        If Trim$(Mid$(slLine, ilLnPos5 - 1, 1)) = "" Then
                                            tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:00" & Trim$(Mid$(slLine, ilLnPos5, 3)))
                                        Else
                                            tmImportSpotInfo(ilUpper).iLen = gLengthToLong("00:0" & Trim$(Mid$(slLine, ilLnPos5 - 1, 4)))
                                        End If
                                        ilLnPos5 = InStr(ilLnPos5, slLine, " ", vbBinaryCompare)
                                        ilAMPMPos = InStr(ilLnPos5, slLine, ":", vbBinaryCompare) - 3
                                        smCallLetters = Trim$(Mid$(slLine, ilLnPos5, ilAMPMPos - ilLnPos5))
                                        smCallLetters = mRemoveBlanks(smCallLetters)
                                        'Parse Times
                                        If tmImportSpotInfo(ilUpper).iLen > 0 Then
                                            blFirstSpot = True
                                            Do
                                                ilAMPMPos = InStr(ilAMPMPos, slLine, "M", vbBinaryCompare)
                                                If ilAMPMPos <= 0 Then
                                                    Exit Do
                                                End If
                                                If Not blFirstSpot Then
                                                    tmImportSpotInfo(ilUpper).lDPStartTime = -1
                                                    tmImportSpotInfo(ilUpper).lDPEndTime = -1
                                                    tmImportSpotInfo(ilUpper).sDPDays = ""
                                                    tmImportSpotInfo(ilUpper).bMatched = False
                                                    tmImportSpotInfo(ilUpper).lRate = tmImportSpotInfo(ilUpper - 1).lRate
                                                    tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(slISCI)
                                                    tmImportSpotInfo(ilUpper).lAirDate = tmImportSpotInfo(ilUpper - 1).lAirDate
                                                    tmImportSpotInfo(ilUpper).iLen = tmImportSpotInfo(ilUpper - 1).iLen
                                                End If
                                                blFirstSpot = False
                                                tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Format(Trim$(Mid$(slLine, ilAMPMPos - 10, 11)), "h:mm:ssAM/PM"), False)
                                                ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                                ilUpper = ilUpper + 1
                                                ilAMPMPos = ilAMPMPos + 11
                                            Loop
                                        End If
                                    End If
                                End If
                            End If
                            slLine = MyFile.ReadLine
                            slLine = UCase$(slLine)
                        Loop While Not MyFile.AtEndOfStream
                        'Match up info
                        If llStartDate <> 99999999 Then
                            ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                            blInvoiceProcessed = True
                            If (ilRet) Then
                                If (Not blAnyFoundSetToFalse) Then
                                    blAnyFound = True
                                End If
                            Else
                                blAnyFound = False
                                blAnyFoundSetToFalse = True
                            End If
                        End If
                        'Look for next Invoice within file
                        For ilLoop = 0 To UBound(slPrevLines) Step 1
                            slPrevLines(ilLoop) = ""
                        Next ilLoop
                        mClearValues llStartDate, llEndDate
                    ElseIf (smSourceForm = "WE") Or (smSourceForm = "ME") Then
                        blFirstSpot = True
                        Do
                            If Not blFirstSpot Then
                                slLine = MyFile.ReadLine
                                slLine = UCase$(slLine)
                            End If
                            blFirstSpot = False
                            If (InStr(1, slLine, "34;", vbTextCompare) > 0) And (slPDForEDI = "E") Then
                                Exit Do
                            End If
                            If (InStr(1, slLine, "41;", vbTextCompare) > 0) And (slPDForEDI = "E") Then
                                '41;Line Number; Days of Week;Start Time; End Time;Rate Detail;Rate per Spot; # Spots; Line Start Date;Line End Date
                                If bl41First Then
                                    gParseItemFields slLine, ";", slFields()
                                    For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
                                        slFields(ilCol + 1) = slFields(ilCol)
                                    Next ilCol
                                    slFields(0) = ""
                                    If slFields(4) = "" Then
                                        llDPStartTime = -1
                                        llDPEndTime = -1
                                    Else
                                        If Len(Trim$(slFields(4))) = 4 Then
                                            llDPStartTime = gTimeToLong(Left(slFields(4), 2) & ":" & Mid$(slFields(4), 3, 2) & ":" & "00", False)
                                            llDPEndTime = gTimeToLong(Left(slFields(5), 2) & ":" & Mid$(slFields(5), 3, 2) & ":" & "00", False)
                                        Else
                                            llDPStartTime = gTimeToLong(Left(slFields(4), 2) & ":" & Mid$(slFields(4), 3, 2) & ":" & Mid$(slFields(4), 5, 2), False)
                                            llDPEndTime = gTimeToLong(Left(slFields(5), 2) & ":" & Mid$(slFields(5), 3, 2) & ":" & Mid$(slFields(5), 5, 2), False)
                                        End If
                                    End If
                                    slDPDays = slFields(3)
                                    slDPDays = Replace(slDPDays, " ", "-")
                                    bl41First = False
                                End If
                            ElseIf (InStr(1, slLine, "51;", vbTextCompare) > 0) And (slPDForEDI = "E") Then
                                '51;Run Code;Run Date;Day of Week;Time of Day;Sot Length;Copy ID;Rate
                                bl41First = True
                                ilUpper = UBound(tmImportSpotInfo)
                                gParseItemFields slLine, ";", slFields()
                                For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
                                    slFields(ilCol + 1) = slFields(ilCol)
                                Next ilCol
                                slFields(0) = ""
                                If slFields(2) = "Y" Then
                                    tmImportSpotInfo(ilUpper).lAirDate = gDateValue(Mid(slFields(3), 3, 2) & "/" & Mid$(slFields(3), 5, 2) & "/" & Mid(slFields(3), 1, 2))
                                    If Len(Trim(slFields(5))) = 4 Then
                                        tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Left(slFields(5), 2) & ":" & Mid$(slFields(5), 3, 2) & ":" & "00", False)
                                    Else
                                        tmImportSpotInfo(ilUpper).lAirTime = gTimeToLong(Left(slFields(5), 2) & ":" & Mid$(slFields(5), 3, 2) & ":" & Mid$(slFields(5), 5, 2), False)
                                    End If
                                    tmImportSpotInfo(ilUpper).iLen = slFields(6)
                                    tmImportSpotInfo(ilUpper).sISCI = mRemoveExtraBlanks(slFields(7))
                                    tmImportSpotInfo(ilUpper).lDPStartTime = llDPStartTime
                                    tmImportSpotInfo(ilUpper).lDPEndTime = llDPEndTime
                                    tmImportSpotInfo(ilUpper).sDPDays = slDPDays
                                    tmImportSpotInfo(ilUpper).bMatched = False
                                    If Trim$(slFields(8)) <> "" Then
                                        tmImportSpotInfo(ilUpper).lRate = gStrDecToLong(slFields(8), 2)
                                    Else
                                        tmImportSpotInfo(ilUpper).lRate = -1
                                    End If
                                    If tmImportSpotInfo(ilUpper).lAirDate < llStartDate Then
                                        llStartDate = tmImportSpotInfo(ilUpper).lAirDate
                                    End If
                                    lmInvStartDate = llStartDate
                                    If tmImportSpotInfo(ilUpper).lAirDate > llEndDate Then
                                        llEndDate = tmImportSpotInfo(ilUpper).lAirDate
                                    End If
                                    ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                                End If
                            End If
                        Loop While Not MyFile.AtEndOfStream
                        'Match up info
                        If llStartDate <> 99999999 Then
                            ilRet = mMatchSpots(llMatchResultRow, llNoMatchResultRow, llStartDate, llEndDate)
                            blInvoiceProcessed = True
                            If (ilRet) Then
                                If (Not blAnyFoundSetToFalse) Then
                                    blAnyFound = True
                                End If
                            Else
                                blAnyFound = False
                                blAnyFoundSetToFalse = True
                            End If
                        End If
                        'Look for next Invoice within file
                        llDPStartTime = -1
                        llDPEndTime = -1
                        slDPDays = ""
                        bl41First = True
                        mClearValues llStartDate, llEndDate
                    End If
                End If
            End If
            'Get next line
            If MyFile.AtEndOfStream Then
                Exit Do
            End If
            slLine = MyFile.ReadLine
        Loop
        'mMRSortCol MRCOMPLIANTINDEX
        'mMRSortCol MRNETCOUNTINDEX
        MyFile.Close
        Set MyFile = Nothing
        If (Not blAnyFound) And (Not blInvoiceProcessed) Then
            If smSourceForm = "" Then
                'Set result to File does not exist in grid
                If llNoMatchResultRow >= grdNoMatchedResult.Rows Then
                    grdNoMatchedResult.AddItem ""
                End If
                grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRFILENAMEINDEX) = smIihfFileName
                grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Source Missing or PDF to Text Failed"
                llNoMatchResultRow = llNoMatchResultRow + 1
            Else
                If (smCallLetters = "") Or (smAdvertiserName = "") Or (smEstimateNumber = "") Or (smContractNumber = "") Or (smInvoiceNumber = "") Then
                    'Set result to File does not exist in grid
                    If llNoMatchResultRow >= grdNoMatchedResult.Rows Then
                        grdNoMatchedResult.AddItem ""
                    End If
                    grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRFILENAMEINDEX) = Trim(smIihfFileName)
                    If (smCallLetters = "") Then
                        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Call Letters Missing"
                    ElseIf (smAdvertiserName = "") Then
                        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Advertiser Missing"
                    ElseIf (smEstimateNumber = "") Then
                        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Estimate # Missing"
                    ElseIf (smContractNumber = "") Then
                        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Contract # Missing or Unsupported Format"
                    ElseIf (smInvoiceNumber = "") Then
                        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "Invoice # Missing"
                    End If
                    llNoMatchResultRow = llNoMatchResultRow + 1
                End If
            End If
            mMoveFile "U", slPDFFileName, llStartDate
        Else
            mMoveFile "P", slPDFFileName, llStartDate
        End If
    Else
        'Set result to File does not exist in grid
        If llNoMatchResultRow >= grdNoMatchedResult.Rows Then
            grdNoMatchedResult.AddItem ""
        End If
        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRFILENAMEINDEX) = smIihfFileName
        grdNoMatchedResult.TextMatrix(llNoMatchResultRow, NMRSTATUSINDEX) = "File Missing"
        llNoMatchResultRow = llNoMatchResultRow + 1
        mMoveFile "U", slPDFFileName, llStartDate
        Exit Sub
    End If
    
    Set oMyFileObj = Nothing
    If Not MyFile Is Nothing Then
        On Error Resume Next
        MyFile.Close
        On Error GoTo 0
        Set MyFile = Nothing
    End If
End Sub

Private Function mMatchSpots(llMatchRow As Long, llNoMatchRow As Long, llInStartDate As Long, llInEndDate As Long, Optional llInChfCode As Long = -1, Optional ilInAdfCode As Integer = -1, Optional blPreviewMode As Boolean = False) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim slAgyEstNo As String
    Dim slAdvertiser As String
    Dim llChfCode As Long
    Dim slCallLetters As String
    Dim llDate As Long
    Dim llMoDate As Long
    Dim slLine As String
    Dim slDate As String
    Dim ilSdf As Integer
    Dim ilRdf As Integer
    Dim llLnStartDate As Long
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilTime As Integer
    Dim ilChf As Integer
    Dim ilCff As Integer
    Dim ilCountDiff As Integer
    Dim ilDPTime As Integer
    Dim blMatch As Boolean
    Dim blFound As Boolean
    Dim ilWeek As Integer
    Dim ilDay As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slSchDate As String
    Dim ilGameNo As Integer
    Dim ilUnresolvedNet As Integer
    Dim ilUnresovledStn As Integer
    Dim ilUnresolvedLines As Integer
    Dim ilImport As Integer
    Dim ilStn As Integer
    Dim ilDP As Integer
    Dim slEstNoPart1 As String
    Dim slEstNoPart2 As String
    Dim ilPos As Integer
    Dim ilMatchCount As Integer
    Dim slAcqCost As String
    Dim blAllowCntr As Boolean
    Dim llAmfCode As Long
    Dim llDPLength As Long
    Dim slDPLength As String
    Dim ilDayRange As Integer
    Dim slDayRange As String
    Dim blDayMatch As Boolean
    Dim blManuallyPosted As Boolean
    Dim ilLen As Integer
    Dim llAcqRate As Long
    Dim ilMPCount1 As Integer
    Dim ilMPCount2 As Integer
    Dim ilMPChf1 As Integer
    Dim ilMPChf2 As Integer
    Dim ilPass As Integer   '0=include acquisition cost; 1= exclude acquisition cost
    ReDim tmNetSpotInfo(0 To 0) As NETSPOTINFO
    ReDim tmCffInfo(0 To 0) As CFFINFO
    
    mMatchSpots = False
    If llInStartDate <> 99999999 Then
        llStartDate = gDateValue(gObtainStartStd(Format(llInStartDate, "m/d/yy")))
    Else
        Exit Function
    End If
    ReDim tmStnMatchInfo(0 To 0) As STNMATCHINFO
    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
        blFound = False
        For ilLen = 0 To UBound(tmStnMatchInfo) - 1 Step 1
            If tmStnMatchInfo(ilLen).iLen = tmImportSpotInfo(ilImport).iLen Then
                blFound = True
                ilWeek = (tmImportSpotInfo(ilImport).lAirDate - llStartDate) \ 7
                tmStnMatchInfo(ilLen).iAirWeek(ilWeek) = tmStnMatchInfo(ilLen).iAirWeek(ilWeek) + 1
                Exit For
            End If
        Next ilLen
        If Not blFound Then
            For ilWeek = 0 To UBound(tmStnMatchInfo(0).iAirWeek) Step 1
                tmStnMatchInfo(UBound(tmStnMatchInfo)).iAirWeek(ilWeek) = 0
            Next ilWeek
            ilWeek = (tmImportSpotInfo(ilImport).lAirDate - llStartDate) \ 7
            tmStnMatchInfo(UBound(tmStnMatchInfo)).iAirWeek(ilWeek) = 1
            tmStnMatchInfo(UBound(tmStnMatchInfo)).iLen = tmImportSpotInfo(ilImport).iLen
            ReDim Preserve tmStnMatchInfo(0 To UBound(tmStnMatchInfo) + 1) As STNMATCHINFO
        End If
    Next ilImport
    'llEndDate = gDateValue(gObtainEndStd(Format(llInEndDate, "m/d/yy")))
    llEndDate = gDateValue(gObtainEndStd(Format(llStartDate, "m/d/yy")))
    ilAdfCode = ilInAdfCode
    llChfCode = llInChfCode
    llAmfCode = 0
    ilVefCode = -1
    slCallLetters = UCase$(smCallLetters)   'UCase$(Trim$(grdMatchedResult.TextMatrix(llMatchRow, MRSTNSTATIONINDEX)))
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If UCase$(Trim$(tgMVef(ilVef).sName)) = slCallLetters Then
            ilVefCode = tgMVef(ilVef).iCode
            Exit For
        End If
    Next ilVef
    If ilVefCode = -1 Then
        If (InStr(1, slCallLetters, "AM", vbTextCompare) > 0) And (InStr(1, slCallLetters, "FM", vbTextCompare) > 0) Then
            ilPos = InStr(1, slCallLetters, "-", vbTextCompare)
            If ilPos > 0 Then
                slCallLetters = Left(slCallLetters, ilPos)
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If UCase$(Trim$(tgMVef(ilVef).sName)) = slCallLetters & "AM" Then
                        ilVefCode = tgMVef(ilVef).iCode
                        slCallLetters = slCallLetters & "AM"
                        Exit For
                    End If
                Next ilVef
                If ilVefCode = -1 Then
                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If UCase$(Trim$(tgMVef(ilVef).sName)) = slCallLetters & "FM" Then
                            ilVefCode = tgMVef(ilVef).iCode
                            slCallLetters = slCallLetters & "FM"
                            Exit For
                        End If
                    Next ilVef
                End If
            End If
        End If
    End If
    If ilVefCode = -1 Then
        ilRet = mAddiihfAndIidf(0, 0, llAmfCode, llStartDate, llNoMatchRow, "Vehicle Name Not Found")
        Exit Function
    End If
    
    If llChfCode = -1 Then
        slAdvertiser = UCase$(smAdvertiserName) 'UCase$(Trim$(grdMatchedResult.TextMatrix(llMatchRow, MRSTNADVERTISERINDEX)))
        tmAmfSrchKey2.iVefCode = ilVefCode
        tmAmfSrchKey2.sStationAdvtName = slAdvertiser
        ilRet = btrGetEqual(hmAmf, tmAmf, imAmfRecLen, tmAmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilAdfCode = tmAmf.iAdfCode
            llAmfCode = tmAmf.lCode
        Else
            For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                If UCase$(Trim$(tgCommAdf(ilAdf).sName)) = slAdvertiser Then
                    ilAdfCode = tgCommAdf(ilAdf).iCode
                    Exit For
                End If
            Next ilAdf
        End If
    Else
        tmChfSrchKey0.lCode = llChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilAdfCode = -1 Then
            ilAdfCode = tmChf.iAdfCode
        End If
    End If
    If smNetContractNumber <> "" Then
        tmChfSrchKey1.lCntrNo = Val(smNetContractNumber)
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(smNetContractNumber)) And ((tmChf.sSchStatus <> "F") And (tmChf.sSchStatus <> "M"))
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(smNetContractNumber)) And (tmChf.sDelete <> "Y") Then
            llChfCode = tmChf.lCode
            If ilAdfCode = -1 Then
                ilAdfCode = tmChf.iAdfCode
            End If
        End If
    End If

    If ilAdfCode = -1 Then
        If llChfCode = -1 Then
            ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "Advertiser Name Not Found")
        Else
            ilRet = mAddiihfAndIidf(ilVefCode, tmChf.lCode, llAmfCode, llStartDate, llNoMatchRow, "Advertiser Name Not Found")
        End If
        Exit Function
    End If
    If llChfCode = -1 Then
        ReDim llEstChfCode(0 To 0) As Long
        slAgyEstNo = smEstimateNumber
        If (slAgyEstNo <> "") And (slAgyEstNo <> "-") Then
            slEstNoPart1 = Trim$(Left$(slAgyEstNo, 10))
            If Len(slAgyEstNo) > 10 Then
                slEstNoPart2 = Trim$(Mid$(slAgyEstNo, 11))
            Else
                slEstNoPart2 = ""
            End If
            tmChfSrchKey4.sAgyEstNo = slEstNoPart1
            tmChfSrchKey4.sTitle = slEstNoPart2
            tmChfSrchKey4.iCntRevNo = 32000
            tmChfSrchKey4.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (Trim$(tmChf.sAgyEstNo) = slEstNoPart1) And (Trim$(tmChf.sTitle) = slEstNoPart2)
                If (tmChf.iAdfCode = ilAdfCode) And ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                    llEstChfCode(UBound(llEstChfCode)) = tmChf.lCode
                    ReDim Preserve llEstChfCode(0 To UBound(llEstChfCode) + 1) As Long
                End If
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If UBound(llEstChfCode) = 1 Then
                llChfCode = llEstChfCode(0)
                tmChfSrchKey0.lCode = llChfCode 'tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
        End If
        If llChfCode = -1 Then
            blManuallyPosted = False
            'Count number of contracts
            ReDim tmMatchCntr(0 To 0) As MATCHCNTR
            ReDim tmMatchCntrLen(0 To 0) As MATCHCNTRLEN
            tmSdfSrchKey7.iAdfCode = ilAdfCode
            gPackDateLong llStartDate, tmSdfSrchKey7.iDate(0), tmSdfSrchKey7.iDate(1)
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey7, INDEXKEY7, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iAdfCode = ilAdfCode)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                If llDate > llEndDate Then
                    Exit Do
                End If
                'If tmSdf.iVefCode = ilVefCode Then
                'Only include lengths defined in the invoice
                blFound = False
                For ilLen = 0 To UBound(tmStnMatchInfo) - 1 Step 1
                    If tmStnMatchInfo(ilLen).iLen = tmSdf.iLen Then
                        blFound = True
                        Exit For
                    End If
                Next ilLen
                If (tmSdf.iVefCode = ilVefCode) And (blFound) Then
                    blAllowCntr = True
                    If UBound(llEstChfCode) > 0 Then
                        blAllowCntr = False
                        For ilChf = 0 To UBound(llEstChfCode) - 1 Step 1
                            If tmSdf.lChfCode = llEstChfCode(ilChf) Then
                                blAllowCntr = True
                                Exit For
                            End If
                        Next ilChf
                    End If
                    If blAllowCntr Then
                        'tmIihfSrchKey2.lChfCode = tmSdf.lChfCode
                        'tmIihfSrchKey2.iVefCode = tmSdf.iVefCode
                        'gPackDateLong llStartDate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
                        'ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        'If (ilRet = BTRV_ERR_NONE) And ((Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C")) Then
                        '    blAllowCntr = False
                        '    blManuallyPosted = True
                        'End If
                    End If
                    If blAllowCntr Then
                        blFound = False
                        For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
                            If tmMatchCntr(ilChf).lChfCode = tmSdf.lChfCode Then
                                tmMatchCntr(ilChf).lSpotCount = tmMatchCntr(ilChf).lSpotCount + 1
                                ilWeek = (llDate - llStartDate) \ 7
                                For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
                                    If (tmMatchCntrLen(ilLen).iLen = tmSdf.iLen) And (tmMatchCntrLen(ilLen).lChfCode = tmSdf.lChfCode) Then
                                        tmMatchCntrLen(ilLen).iAirWeek(ilWeek) = tmMatchCntrLen(ilLen).iAirWeek(ilWeek) + 1
                                        'If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                        If mIncludeSpot(llStartDate, llEndDate) Then
                                            tmMatchCntrLen(ilLen).iSchWeek(ilWeek) = tmMatchCntrLen(ilLen).iSchWeek(ilWeek) + 1
                                        End If
                                        blFound = True
                                        Exit For
                                    End If
                                Next ilLen
                                If Not blFound Then
                                    tmMatchCntrLen(UBound(tmMatchCntrLen)).lChfCode = tmSdf.lChfCode
                                    tmMatchCntrLen(UBound(tmMatchCntrLen)).iLen = tmSdf.iLen
                                    For ilWeek = 0 To UBound(tmMatchCntrLen(0).iAirWeek) Step 1
                                        tmMatchCntrLen(UBound(tmMatchCntrLen)).iAirWeek(ilWeek) = 0
                                        tmMatchCntrLen(UBound(tmMatchCntrLen)).iSchWeek(ilWeek) = 0
                                    Next ilWeek
                                    ilWeek = (llDate - llStartDate) \ 7
                                    tmMatchCntrLen(UBound(tmMatchCntrLen)).iAirWeek(ilWeek) = 1
                                    'If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                    If mIncludeSpot(llStartDate, llEndDate) Then
                                        tmMatchCntrLen(UBound(tmMatchCntrLen)).iSchWeek(ilWeek) = 1
                                    End If
                                    ReDim Preserve tmMatchCntrLen(0 To UBound(tmMatchCntrLen) + 1) As MATCHCNTRLEN
                                End If
                                blFound = True
                                Exit For
                            End If
                        Next ilChf
                        If Not blFound Then
                            tmMatchCntr(UBound(tmMatchCntr)).lChfCode = tmSdf.lChfCode
                            tmMatchCntr(UBound(tmMatchCntr)).lSpotCount = 1
                            ReDim Preserve tmMatchCntr(0 To UBound(tmMatchCntr) + 1) As MATCHCNTR
                            tmMatchCntrLen(UBound(tmMatchCntrLen)).lChfCode = tmSdf.lChfCode
                            tmMatchCntrLen(UBound(tmMatchCntrLen)).iLen = tmSdf.iLen
                            For ilWeek = 0 To UBound(tmMatchCntrLen(0).iAirWeek) Step 1
                                tmMatchCntrLen(UBound(tmMatchCntrLen)).iAirWeek(ilWeek) = 0
                                tmMatchCntrLen(UBound(tmMatchCntrLen)).iSchWeek(ilWeek) = 0
                            Next ilWeek
                            ilWeek = (llDate - llStartDate) \ 7
                            tmMatchCntrLen(UBound(tmMatchCntrLen)).iAirWeek(ilWeek) = 1
                            'If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                            If mIncludeSpot(llStartDate, llEndDate) Then
                                tmMatchCntrLen(UBound(tmMatchCntrLen)).iSchWeek(ilWeek) = 1
                            End If
                            ReDim Preserve tmMatchCntrLen(0 To UBound(tmMatchCntrLen) + 1) As MATCHCNTRLEN
                        End If
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If UBound(tmMatchCntr) > 1 Then
                'If air in non-ordered week ignore contract
                ReDim tlMatchCntr(0 To UBound(tmMatchCntr)) As MATCHCNTR
                For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
                    tlMatchCntr(ilChf) = tmMatchCntr(ilChf)
                Next ilChf
                ReDim tmMatchCntr(0 To 0) As MATCHCNTR
                For ilChf = 0 To UBound(tlMatchCntr) - 1 Step 1
                    blAllowCntr = True
                    For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
                        If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
                            blFound = False
                            For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
                                If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
                                    blFound = True
                                    Exit For
                                End If
                            Next ilImport
                            If Not blFound Then
                                blAllowCntr = False
                                Exit For
                            End If
                        End If
                    Next ilLen
                    If blAllowCntr Then
                        For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
                            blFound = False
                            For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
                                If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
                                    If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
                                        blFound = True
                                        Exit For
                                    End If
                                End If
                            Next ilLen
                            If Not blFound Then
                                blAllowCntr = False
                                Exit For
                            End If
                        Next ilImport
                    End If
'                    If blAllowCntr Then
'                        For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
'                            If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
'                                For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
'                                    If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
'                                        For ilWeek = 0 To UBound(tmMatchCntrLen(ilLen).iAirWeek) Step 1
'                                            If (tmMatchCntrLen(ilLen).iAirWeek(ilWeek) <= 0) And (tmStnMatchInfo(ilImport).iAirWeek(ilWeek) > 0) Then
'                                                blAllowCntr = False
'                                                Exit For
'                                            End If
'                                        Next ilWeek
'                                    End If
'                                    If Not blAllowCntr Then
'                                        Exit For
'                                    End If
'                                Next ilImport
'                            End If
'                            If Not blAllowCntr Then
'                                Exit For
'                            End If
'                        Next ilLen
'                    End If
                    If blAllowCntr Then
                        tmMatchCntr(UBound(tmMatchCntr)) = tlMatchCntr(ilChf)
                        ReDim Preserve tmMatchCntr(0 To UBound(tmMatchCntr) + 1) As MATCHCNTR
                    End If
                Next ilChf
            End If
            If UBound(tmMatchCntr) <= LBound(tmMatchCntr) Then
                If blManuallyPosted Then
                    ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "Network Contract Manually Posted")
                Else
                    ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "No Network Contract Found for the Advertiser")
                End If
                Exit Function
            Else
                If UBound(tmMatchCntr) > 1 Then
                    'Check if manually posted
                    ilMPCount1 = 0
                    ilMPCount2 = 0
                    ilMPChf1 = 0
                    ilMPChf2 = 0
                    For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
                        tmIihfSrchKey2.lChfCode = tmMatchCntr(ilChf).lChfCode
                        tmIihfSrchKey2.iVefCode = ilVefCode
                        gPackDateLong llStartDate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
                        ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                        If (ilRet = BTRV_ERR_NONE) And ((Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C")) Then
                            ilMPCount1 = ilMPCount1 + 1
                            ilMPChf1 = ilChf
                            If (Trim$(tmIihf.sStnInvoiceNo) <> "") Or (Trim$(tmIihf.sStnContractNo) <> "") Then
                                If ((StrComp(Trim$(tmIihf.sStnInvoiceNo), smInvoiceNumber, vbBinaryCompare) = 0) And (Trim$(tmIihf.sStnInvoiceNo) <> "")) Or ((StrComp(Trim$(tmIihf.sStnContractNo), smContractNumber, vbBinaryCompare) = 0) And (Trim$(tmIihf.sStnContractNo) <> "")) Then
                                    ilMPCount2 = ilMPCount2 + 1
                                    ilMPChf2 = ilChf
                                End If
                            End If
                        End If
                    Next ilChf
                    If ilMPCount2 = 1 Then
                        ReDim tlMatchCntr(0 To 1) As MATCHCNTR
                        tlMatchCntr(0) = tmMatchCntr(ilMPChf2)
                        ReDim tmMatchCntr(0 To 1) As MATCHCNTR
                        tmMatchCntr(0) = tlMatchCntr(0)
'                    ElseIf ilMPCount1 = 1 Then
'                        blAllowCntr = True
'                        If blAllowCntr Then
'                            For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
'                                If tmMatchCntrLen(ilLen).lChfCode = tmMatchCntr(ilMPChf1).lChfCode Then
'                                    For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
'                                        If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
'                                            For ilWeek = 0 To UBound(tmMatchCntrLen(ilLen).iAirWeek) Step 1
'                                                If (tmMatchCntrLen(ilLen).iSchWeek(ilWeek) <= tmStnMatchInfo(ilImport).iAirWeek(ilWeek)) Then
'                                                Else
'                                                    blAllowCntr = False
'                                                    Exit For
'                                                End If
'                                            Next ilWeek
'                                        End If
'                                        If Not blAllowCntr Then
'                                            Exit For
'                                        End If
'                                    Next ilImport
'                                End If
'                                If Not blAllowCntr Then
'                                    Exit For
'                                End If
'                            Next ilLen
'                        End If
'                        If blAllowCntr Then
'                            ReDim tlMatchCntr(0 To 1) As MATCHCNTR
'                            tlMatchCntr(0) = tmMatchCntr(ilMPChf1)
'                            ReDim tmMatchCntr(0 To 1) As MATCHCNTR
'                            tmMatchCntr(0) = tlMatchCntr(0)
'                        End If
                    End If
'                    If UBound(tmMatchCntr) > 1 Then
'                        'Loop for matching weeks only
'                        ReDim tlMatchCntr(0 To UBound(tmMatchCntr)) As MATCHCNTR
'                        For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
'                            tlMatchCntr(ilChf) = tmMatchCntr(ilChf)
'                        Next ilChf
'                        ReDim tmMatchCntr(0 To 0) As MATCHCNTR
'                        For ilChf = 0 To UBound(tlMatchCntr) - 1 Step 1
'                            blAllowCntr = True
'        '                    For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
'        '                        If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
'        '                            blFound = False
'        '                            For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
'        '                                If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
'        '                                    blFound = True
'        '                                    Exit For
'        '                                End If
'        '                            Next ilImport
'        '                            If Not blFound Then
'        '                                blAllowCntr = False
'        '                                Exit For
'        '                            End If
'        '                        End If
'        '                    Next ilLen
'        '                    If blAllowCntr Then
'        '                        For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
'        '                            blFound = False
'        '                            For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
'        '                                If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
'        '                                    If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
'        '                                        blFound = True
'        '                                        Exit For
'        '                                    End If
'        '                                End If
'        '                            Next ilLen
'        '                            If Not blFound Then
'        '                                blAllowCntr = False
'        '                                Exit For
'        '                            End If
'        '                        Next ilImport
'        '                    End If
'                            If blAllowCntr Then
'                                For ilLen = 0 To UBound(tmMatchCntrLen) - 1 Step 1
'                                    If tmMatchCntrLen(ilLen).lChfCode = tlMatchCntr(ilChf).lChfCode Then
'                                        For ilImport = 0 To UBound(tmStnMatchInfo) - 1 Step 1
'                                            If tmStnMatchInfo(ilImport).iLen = tmMatchCntrLen(ilLen).iLen Then
'                                                For ilWeek = 0 To UBound(tmMatchCntrLen(ilLen).iAirWeek) Step 1
'                                                    If ((tmMatchCntrLen(ilLen).iAirWeek(ilWeek) > 0) And (tmStnMatchInfo(ilImport).iAirWeek(ilWeek) > 0)) Or ((tmMatchCntrLen(ilLen).iAirWeek(ilWeek) = 0) And (tmStnMatchInfo(ilImport).iAirWeek(ilWeek) = 0)) Then
'                                                    Else
'                                                        blAllowCntr = False
'                                                        Exit For
'                                                    End If
'                                                Next ilWeek
'                                            End If
'                                            If Not blAllowCntr Then
'                                                Exit For
'                                            End If
'                                        Next ilImport
'                                    End If
'                                    If Not blAllowCntr Then
'                                        Exit For
'                                    End If
'                                Next ilLen
'                            End If
'                            If blAllowCntr Then
'                                tmMatchCntr(UBound(tmMatchCntr)) = tlMatchCntr(ilChf)
'                                ReDim Preserve tmMatchCntr(0 To UBound(tmMatchCntr) + 1) As MATCHCNTR
'                            End If
'                        Next ilChf
'                    End If
                End If
                If UBound(tmMatchCntr) <> 1 Then
                    ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "Multi-Contracts found as possible matches")
                    Exit Function
'                    'Find exact match
'                    For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
'                        If tmMatchCntr(ilChf).iSpotCount = UBound(tmImportSpotInfo) Then
'                            If llChfCode = -1 Then
'                                llChfCode = tmMatchCntr(ilChf).lChfCode
'                            Else
'                                ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "Multi-Contracts Reference Advertiser")
'                                Exit Function
'                            End If
'                        End If
'                    Next ilChf
'                    If llChfCode = -1 Then
'                        'Find closest match
'                        ilMatchCount = 0
'                        For ilChf = 0 To UBound(tmMatchCntr) - 1 Step 1
'                            If tmMatchCntr(ilChf).iSpotCount > UBound(tmImportSpotInfo) Then
'                                If llChfCode = -1 Then
'                                    llChfCode = tmMatchCntr(ilChf).lChfCode
'                                    ilCountDiff = tmMatchCntr(ilChf).iSpotCount - UBound(tmImportSpotInfo)
'                                    ilMatchCount = 1
'                                Else
'                                    If tmMatchCntr(ilChf).iSpotCount - UBound(tmImportSpotInfo) < ilCountDiff Then
'                                        llChfCode = tmMatchCntr(ilChf).lChfCode
'                                        ilCountDiff = tmMatchCntr(ilChf).iSpotCount - UBound(tmImportSpotInfo)
'                                        ilMatchCount = 1
'                                    ElseIf tmMatchCntr(ilChf).iSpotCount - UBound(tmImportSpotInfo) = ilCountDiff Then
'                                        'ilRet = mAddiihfAndIidf(ilVefCode, 0, llStartDate, llNoMatchRow, "Multi-Contracts Reference Advertiser")
'                                        'Exit Sub
'                                        ilMatchCount = ilMatchCount + 1
'                                    End If
'                                End If
'                            End If
'                        Next ilChf
'                        If (llChfCode = -1) Or (ilMatchCount > 1) Then
'                            ilRet = mAddiihfAndIidf(ilVefCode, 0, llAmfCode, llStartDate, llNoMatchRow, "Multi-Contracts found as possible matches")
'                            Exit Function
'                        End If
'                    End If
                Else
                    llChfCode = tmMatchCntr(0).lChfCode
                End If
                tmChfSrchKey0.lCode = llChfCode 'tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
        End If
    End If
    
    ReDim llSdfCode(0 To 0) As Long
    tmSdfSrchKey7.iAdfCode = ilAdfCode
    gPackDateLong llStartDate, tmSdfSrchKey7.iDate(0), tmSdfSrchKey7.iDate(1)
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey7, INDEXKEY7, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iAdfCode = ilAdfCode)
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
        If llDate > llEndDate Then
            Exit Do
        End If
        'Only include lengths defined in the invoice
        blFound = False
        For ilLoop = 0 To UBound(tmStnMatchInfo) - 1 Step 1
            If tmStnMatchInfo(ilLoop).iLen = tmSdf.iLen Then
                blFound = True
                Exit For
            End If
        Next ilLoop
        If (tmSdf.iVefCode = ilVefCode) And (blFound) Then
            If llChfCode = tmSdf.lChfCode Then
                llSdfCode(UBound(llSdfCode)) = tmSdf.lCode
                ReDim Preserve llSdfCode(0 To UBound(llSdfCode) + 1) As Long
            End If
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If UBound(llSdfCode) <= LBound(llSdfCode) Then
        If Not blPreviewMode Then
            ilRet = mAddiihfAndIidf(ilVefCode, llChfCode, llAmfCode, llStartDate, llNoMatchRow, "No Network Spots found for Vehicle")
        End If
        Exit Function
    End If
    
    If Not blPreviewMode Then
        mAddMapAdf ilVefCode, llAmfCode
    End If
    
    If Not blPreviewMode Then
        mClearIihfAndIidf ilVefCode, llChfCode, llStartDate
        ilRet = mAddIihf(ilVefCode, llChfCode, llAmfCode, llStartDate)
        If ilRet <> BTRV_ERR_NONE Then
            mAddToNoMatchGrid llNoMatchRow, tmIihf.lCode, "Unable to Add Import Header File " & ilRet
            Exit Function
        End If
    End If
    
    If Not blPreviewMode Then
        mAddToMatchGrid llMatchRow, tmIihf.lCode
    End If
    
    tmChfSrchKey0.lCode = llChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmChf.iCntRevNo = 0
        tmChf.iPropVer = 0
    End If
    
    
    For ilSdf = 0 To UBound(llSdfCode) - 1 Step 1
        tmSdfSrchKey3.lCode = llSdfCode(ilSdf)
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
        If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
            imSelectedDay = gWeekDayStr(slSchDate)
            ilGameNo = 0
            If Not blPreviewMode Then
                ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
            Else
                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    'Obtain Original date
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        tmSdf.iVefCode = tmSmf.iOrigSchVef
                        tmSdf.iDate(0) = tmSmf.iMissedDate(0)
                        tmSdf.iDate(1) = tmSmf.iMissedDate(1)
                        tmSdf.iTime(0) = tmSmf.iMissedTime(0)
                        tmSdf.iTime(1) = tmSmf.iMissedTime(1)
                    End If
                End If
            End If
        End If
        slLine = tmSdf.iLineNo
        Do While Len(slLine) < 5
            slLine = "0" & slLine
        Loop
        slDate = llDate
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
        If tmClf.iLine <> tmSdf.iLineNo Then
            tmClfSrchKey0.lChfCode = tmSdf.lChfCode
            tmClfSrchKey0.iLine = tmSdf.iLineNo
            tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
            tmClfSrchKey0.iPropVer = tmChf.iPropVer
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                tmClf.lAcquisitionCost = 0
                llDPLength = 86400
                llLnStartDate = 0
            Else
                gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnStartDate
                ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
                If ilRdf <> -1 Then
                    tmRdf = tgMRdf(ilRdf)
                End If
                If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
                    gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llRdfStartTime
                    gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llRdfEndTime
                    llDPLength = llRdfEndTime - llRdfStartTime
                Else
                    If ilRdf <> -1 Then
                        llDPLength = 0
                        For ilDP = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                            If (tmRdf.iStartTime(0, ilDP) <> 1) Or (tmRdf.iStartTime(1, ilDP) <> 0) Then
                                gUnpackTimeLong tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), False, llRdfStartTime
                                gUnpackTimeLong tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), True, llRdfEndTime
                                If llRdfStartTime < llRdfEndTime Then
                                    llDPLength = llDPLength + llRdfEndTime - llRdfStartTime
                                Else
                                    llDPLength = llDPLength + 86400 - llRdfStartTime
                                    llDPLength = llDPLength + llRdfEndTime
                                End If
                            End If
                        Next ilDP
                    Else
                        llDPLength = 86400
                    End If
                End If
            End If
        End If
        slAcqCost = tmClf.lAcquisitionCost
        Do While Len(slAcqCost) < 10
            slAcqCost = "0" & slAcqCost
        Loop
        slDPLength = llDPLength
        Do While Len(slDPLength) < 6
            slDPLength = "0" & slDPLength
        Loop
        ilDayRange = 0
        tmCffSrchKey0.lChfCode = tmClf.lChfCode
        tmCffSrchKey0.iClfLine = tmClf.iLine
        tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
        tmCffSrchKey0.iPropVer = tmClf.iPropVer
        gPackDateLong llLnStartDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmCff.iClfLine = tmClf.iLine) And (tmCff.lChfCode = tmClf.lChfCode)
            gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), tmCffInfo(UBound(tmCffInfo)).lStartDate
            gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), tmCffInfo(UBound(tmCffInfo)).lEndDate
            If tmCffInfo(UBound(tmCffInfo)).lStartDate > llEndDate Then
                Exit Do
            End If
            If (llDate >= tmCffInfo(UBound(tmCffInfo)).lStartDate) And (llDate <= tmCffInfo(UBound(tmCffInfo)).lEndDate) Then
                For ilDay = 0 To 6 Step 1
                    If tmCff.iDay(ilDay) > 0 Then
                        ilDayRange = ilDayRange + 1
                    End If
                Next ilDay
                Exit Do
            End If
            ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        slDayRange = ilDayRange
        'tmNetSpotInfo(UBound(tmNetSpotInfo)).sKey = slAcqCost & slLine & slDate
        tmNetSpotInfo(UBound(tmNetSpotInfo)).sKey = slDayRange & slDPLength & slLine & slDate
        tmNetSpotInfo(UBound(tmNetSpotInfo)).tSdf = tmSdf
        tmNetSpotInfo(UBound(tmNetSpotInfo)).lAcqCost = tmClf.lAcquisitionCost
        tmNetSpotInfo(UBound(tmNetSpotInfo)).bMatched = False
        tmNetSpotInfo(UBound(tmNetSpotInfo)).lCffIndex = -1
        ReDim Preserve tmNetSpotInfo(0 To UBound(tmNetSpotInfo) + 1) As NETSPOTINFO
    Next ilSdf
    If UBound(tmNetSpotInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tmNetSpotInfo(), 0), UBound(tmNetSpotInfo), 0, LenB(tmNetSpotInfo(0)), 0, LenB(tmNetSpotInfo(0).sKey), 0
    End If
    If UBound(tmNetSpotInfo) > LBound(tmNetSpotInfo) Then
        tmChf.lCode = -1
        tmClf.iLine = -1    'Required to build the DP times
        llAcqRate = 0
        For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
            tmSdf = tmNetSpotInfo(ilSdf).tSdf
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
            'If (tmSdf.iLineNo <> tmClf.iLine) Or (tmSdf.lChfCode <> tmClf.lChfCode) Then
            If (tmSdf.lChfCode <> tmChf.lCode) Then
                tmChfSrchKey0.lCode = tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            If (tmSdf.iLineNo <> tmClf.iLine) Then
                tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                tmClfSrchKey0.iLine = tmSdf.iLineNo
                tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
                tmClfSrchKey0.iPropVer = tmChf.iPropVer
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnStartDate
                ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
                If ilRdf <> -1 Then
                    tmRdf = tgMRdf(ilRdf)
                End If
                If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
                    ReDim llStartTime(0 To 1) As Long
                    ReDim llEndTime(0 To 1) As Long
                    gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llStartTime(0)
                    gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llEndTime(0)
                Else
                    ReDim llStartTime(0 To 0) As Long
                    ReDim llEndTime(0 To 0) As Long
                    If ilRdf <> -1 Then
                        For ilDP = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                            If (tmRdf.iStartTime(0, ilDP) <> 1) Or (tmRdf.iStartTime(1, ilDP) <> 0) Then
                                gUnpackTimeLong tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), False, llRdfStartTime
                                gUnpackTimeLong tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), True, llRdfEndTime
                                If llRdfStartTime < llRdfEndTime Then
                                    llStartTime(UBound(llStartTime)) = llRdfStartTime
                                    llEndTime(UBound(llEndTime)) = llRdfEndTime
                                    ReDim Preserve llStartTime(0 To UBound(llStartTime) + 1) As Long
                                    ReDim Preserve llEndTime(0 To UBound(llEndTime) + 1) As Long
                                Else
                                    llStartTime(UBound(llStartTime)) = llRdfStartTime
                                    llEndTime(UBound(llEndTime)) = 86400
                                    ReDim Preserve llEndTime(0 To UBound(llEndTime) + 1) As Long
                                    llStartTime(UBound(llStartTime)) = 0
                                    llEndTime(UBound(llEndTime)) = llRdfEndTime
                                    ReDim Preserve llEndTime(0 To UBound(llEndTime) + 1) As Long
                                End If
                            End If
                        Next ilDP
                    End If
                End If
            End If
            llAcqRate = tmNetSpotInfo(ilSdf).lAcqCost
            tmCffSrchKey0.lChfCode = tmClf.lChfCode
            tmCffSrchKey0.iClfLine = tmClf.iLine
            tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
            tmCffSrchKey0.iPropVer = tmClf.iPropVer
            gPackDateLong llLnStartDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
            ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmCff.iClfLine = tmClf.iLine) And (tmCff.lChfCode = tmClf.lChfCode)
                gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), tmCffInfo(UBound(tmCffInfo)).lStartDate
                gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), tmCffInfo(UBound(tmCffInfo)).lEndDate
                If tmCffInfo(UBound(tmCffInfo)).lStartDate > llEndDate Then
                    Exit Do
                End If
                If (llDate >= tmCffInfo(UBound(tmCffInfo)).lStartDate) And (llDate <= tmCffInfo(UBound(tmCffInfo)).lEndDate) Then
                    blFound = False
                    llMoDate = llDate
                    Do While gWeekDayLong(llMoDate) <> 0
                        llMoDate = llMoDate - 1
                    Loop
                    For ilCff = 0 To UBound(tmCffInfo) - 1 Step 1
                        If (tmCffInfo(ilCff).lCode = tmCff.lCode) And (tmCffInfo(ilCff).lMoDate = llMoDate) Then
                            tmNetSpotInfo(ilSdf).lCffIndex = ilCff
                            blFound = True
                            Exit For
                        End If
                    Next ilCff
                    If Not blFound Then
                        tmNetSpotInfo(ilSdf).lCffIndex = UBound(tmCffInfo)
                        tmCffInfo(UBound(tmCffInfo)).lCode = tmCff.lCode
                        If tmCffInfo(UBound(tmCffInfo)).lStartDate <= llMoDate Then
                            tmCffInfo(UBound(tmCffInfo)).lStartDate = llMoDate
                        End If
                        If tmCffInfo(UBound(tmCffInfo)).lEndDate < llMoDate + 6 Then
                            tmCffInfo(UBound(tmCffInfo)).lEndDate = llMoDate + 6
                        End If
                        tmCffInfo(UBound(tmCffInfo)).iSpotsWk = tmCff.iSpotsWk
                        For ilDay = 0 To 6 Step 1
                            tmCffInfo(UBound(tmCffInfo)).iDay(ilDay) = tmCff.iDay(ilDay)
                        Next ilDay
                        If tmCff.iSpotsWk > 0 Then
                            tmCffInfo(UBound(tmCffInfo)).sDW = "W"
                        Else
                            tmCffInfo(UBound(tmCffInfo)).sDW = "D"
                        End If
                        ReDim Preserve tmCffInfo(0 To UBound(tmCffInfo) + 1) As CFFINFO
                    End If
                    Exit Do
                End If
                ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            blMatch = False
            If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmNetSpotInfo(ilSdf).lCffIndex <> -1) Then
                For ilPass = 0 To 1 Step 1
                    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                        If tmImportSpotInfo(ilImport).bMatched <> True Then
                            If (tmImportSpotInfo(ilImport).iLen = tmClf.iLen) And ((ilPass = 1) Or ((ilPass = 0) And (tmImportSpotInfo(ilImport).lRate = llAcqRate))) Then
                                'Daypart match
                                If tmImportSpotInfo(ilImport).lDPStartTime <> -1 Then
                                    For ilDPTime = 0 To UBound(llStartTime) - 1 Step 1
                                        If (llStartTime(ilDPTime) = tmImportSpotInfo(ilImport).lDPStartTime And (llEndTime(ilDPTime) = tmImportSpotInfo(ilImport).lDPEndTime)) Then
                                            'Test day
                                            If (tmImportSpotInfo(ilImport).lAirDate >= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lStartDate) And (tmImportSpotInfo(ilImport).lAirDate <= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lEndDate) Then
                                                blDayMatch = True
                                                If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW <> "W" Then
                                                    If llDate <> tmImportSpotInfo(ilImport).lAirDate Then
                                                        blDayMatch = False
                                                    End If
                                                End If
                                                If blDayMatch And (tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) > 0) Then
                                                    'Match found, Book spot
                                                    If mBookSpot(ilSdf, ilImport, "A", blPreviewMode) Then
                                                        If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW = "W" Then
                                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk - 1
                                                            If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk <= 0 Then
                                                                For ilDay = 0 To 6 Step 1
                                                                    tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(ilDay) = 0
                                                                Next ilDay
                                                            End If
                                                        Else
                                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) - 1
                                                        End If
                                                        tmImportSpotInfo(ilImport).bMatched = True
                                                        If Not blPreviewMode Then
                                                            grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX)) + 1
                                                        End If
                                                        tmNetSpotInfo(ilSdf).bMatched = True
                                                        If blPreviewMode Then
                                                            mBuildIidf "C", ilSdf, ilImport, "N"
                                                        End If
                                                        blMatch = True
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next ilDPTime
                                    If blMatch Then
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next ilImport
                    If Not blMatch Then
                        For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                            If tmImportSpotInfo(ilImport).bMatched <> True Then
                                If (tmImportSpotInfo(ilImport).iLen = tmClf.iLen) And ((ilPass = 1) Or ((ilPass = 0) And (tmImportSpotInfo(ilImport).lRate = llAcqRate))) Then
                                    'Test times
                                    For ilTime = 0 To UBound(llStartTime) - 1 Step 1
                                        If (tmImportSpotInfo(ilImport).lAirTime >= llStartTime(ilTime)) And (tmImportSpotInfo(ilImport).lAirTime <= llEndTime(ilTime)) Then
                                            'Test day
                                            If (tmImportSpotInfo(ilImport).lAirDate >= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lStartDate) And (tmImportSpotInfo(ilImport).lAirDate <= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lEndDate) Then
                                                blDayMatch = True
                                                If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW <> "W" Then
                                                    If llDate <> tmImportSpotInfo(ilImport).lAirDate Then
                                                        blDayMatch = False
                                                    End If
                                                End If
                                                If blDayMatch And (tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) > 0) Then
                                                    'Match found, Book spot
                                                    If mBookSpot(ilSdf, ilImport, "A", blPreviewMode) Then
                                                        If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW = "W" Then
                                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk - 1
                                                            If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk <= 0 Then
                                                                For ilDay = 0 To 6 Step 1
                                                                    tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(ilDay) = 0
                                                                Next ilDay
                                                            End If
                                                        Else
                                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) - 1
                                                        End If
                                                        tmImportSpotInfo(ilImport).bMatched = True
                                                        If Not blPreviewMode Then
                                                            grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX)) + 1
                                                        End If
                                                        tmNetSpotInfo(ilSdf).bMatched = True
                                                        If blPreviewMode Then
                                                            mBuildIidf "C", ilSdf, ilImport, "N"
                                                        End If
                                                        blMatch = True
                                                        Exit For
                                                    End If
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    Next ilTime
                                End If
                                If blMatch Then
                                    Exit For
                                End If
                            End If
                        Next ilImport
                    End If
                    If blMatch Then
                        Exit For
                    End If
                Next ilPass
            End If
            'If Not blMatch Then
            '    ilRet = mAddIidf("M", ilSdf, 0)
            '    grdMatchedResult.TextMatrix(llMatchRow, NETWORKCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, NETWORKCOUNTINDEX)) + 1
            'Else
            '    grdMatchedResult.TextMatrix(llMatchRow, MATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MATCHCOUNTINDEX)) + 1
            '    tmNetSpotInfo(ilSdf).bMatched = True
            'End If
        Next ilSdf
'        'Any Spots not resolved?
'        ilUnresolvedNet = 0
'        ilUnresolvedLines = -1
'        For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
'            If tmNetSpotInfo(ilSdf).bMatched = False Then
'                ilUnresolvedNet = ilUnresolvedNet + 1
'                If ilUnresolvedLines = -1 Then
'                    ilUnresolvedLines = 1
'                Else
'                    For ilLoop = 0 To ilSdf - 1 Step 1
'                        If tmNetSpotInfo(ilLoop).bMatched = False Then
'                            If tmNetSpotInfo(ilSdf).tSdf.iLineNo <> tmNetSpotInfo(ilLoop).tSdf.iLineNo Then
'                                ilUnresolvedLines = ilUnresolvedLines + 1
'                                Exit For
'                            End If
'                        End If
'                    Next ilLoop
'                End If
'            End If
'        Next ilSdf
'        ilUnresovledStn = 0
'        For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
'            If tmImportSpotInfo(ilImport).bMatched <> True Then
'                ilUnresovledStn = ilUnresovledStn + 1
'            End If
'        Next ilImport
'        If (ilUnresolvedNet = ilUnresovledStn) Or (ilUnresolvedLines = 1) Then
'            For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
'                If (tmNetSpotInfo(ilSdf).bMatched = False) And (tmNetSpotInfo(ilSdf).lCffIndex >= 0) Then
'                    tmSdf = tmNetSpotInfo(ilSdf).tSdf
'                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
'                    blMatch = False
'                    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
'                        If (tmImportSpotInfo(ilImport).bMatched <> True) Then
'                            If tmImportSpotInfo(ilImport).iLen = tmSdf.iLen Then
'                                'Test day
'                                If (tmImportSpotInfo(ilImport).lAirDate >= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lStartDate) And (tmImportSpotInfo(ilImport).lAirDate <= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lEndDate) Then
'                                    blDayMatch = True
'                                    If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW <> "W" Then
'                                        If llDate <> tmImportSpotInfo(ilImport).lAirDate Then
'                                            blDayMatch = False
'                                        End If
'                                    End If
'                                    If blDayMatch And (tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) > 0) Then
'                                        'Match found, Book spot
'                                        If mBookSpot(ilSdf, ilImport, "O") Then
'                                            If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW = "W" Then
'                                                tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk - 1
'                                                If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk <= 0 Then
'                                                    For ilDay = 0 To 6 Step 1
'                                                        tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(ilDay) = 0
'                                                    Next ilDay
'                                                End If
'                                            Else
'                                                tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) - 1
'                                            End If
'                                            tmImportSpotInfo(ilImport).bMatched = True
'                                            grdMatchedResult.TextMatrix(llMatchRow, MRCOMPLIANTINDEX) = "No"
'                                            grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX)) + 1
'                                            tmNetSpotInfo(ilSdf).bMatched = True
'                                            blMatch = True
'                                        End If
'                                    End If
'                                    Exit For
'                                End If
'                            End If
'                        End If
'                    Next ilImport
'                End If
'            Next ilSdf
'            For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
'                If (tmNetSpotInfo(ilSdf).bMatched = False) And (tmNetSpotInfo(ilSdf).lCffIndex >= 0) Then
'                    tmSdf = tmNetSpotInfo(ilSdf).tSdf
'                    blMatch = False
'                    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
'                        If tmImportSpotInfo(ilImport).bMatched <> True Then
'                            If tmImportSpotInfo(ilImport).iLen = tmSdf.iLen Then
'                                'Test day
'                                If (tmImportSpotInfo(ilImport).lAirDate >= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lEndDate) And (tmImportSpotInfo(ilImport).lAirDate <= tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).lEndDate) Then
'                                    'Match found, Book spot
'                                    If mBookSpot(ilSdf, ilImport, "O") Then
'                                        If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).sDW = "W" Then
'                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk - 1
'                                            If tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iSpotsWk <= 0 Then
'                                                For ilDay = 0 To 6 Step 1
'                                                    tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(ilDay) = 0
'                                                Next ilDay
'                                            End If
'                                        Else
'                                            tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) = tmCffInfo(tmNetSpotInfo(ilSdf).lCffIndex).iDay(gWeekDayLong(tmImportSpotInfo(ilImport).lAirDate)) - 1
'                                        End If
'                                        tmImportSpotInfo(ilImport).bMatched = True
'                                        grdMatchedResult.TextMatrix(llMatchRow, MRCOMPLIANTINDEX) = "No"
'                                        grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX)) + 1
'                                        tmNetSpotInfo(ilSdf).bMatched = True
'                                        blMatch = True
'                                    End If
'                                    Exit For
'                                End If
'                            End If
'                        End If
'                    Next ilImport
'                End If
'            Next ilSdf
'        End If
        If Not blPreviewMode Then
            For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
                If tmNetSpotInfo(ilSdf).bMatched = False Then
                    ilRet = mAddIidf("M", ilSdf, 0, "N")
                    'If tmNetSpotInfo(ilSdf).lAcqCost > 0 Then
                        grdMatchedResult.TextMatrix(llMatchRow, MRNETCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRNETCOUNTINDEX)) + 1
                    'End If
                End If
            Next ilSdf
            For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                If tmImportSpotInfo(ilImport).bMatched = False Then
                    ilRet = mAddIidf("I", 0, ilImport, "N")
                    grdMatchedResult.TextMatrix(llMatchRow, MRSTNCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRSTNCOUNTINDEX)) + 1
                End If
            Next ilImport
        Else
            'Rebuild Iidf
            
            For ilSdf = 0 To UBound(tmNetSpotInfo) - 1 Step 1
                If tmNetSpotInfo(ilSdf).bMatched = False Then
                    mBuildIidf "M", ilSdf, 0, "N"
                End If
            Next ilSdf
            For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                If tmImportSpotInfo(ilImport).bMatched = False Then
                    mBuildIidf "I", 0, ilImport, "N"
                End If
            Next ilImport
        End If
    Else
        If Not blPreviewMode Then
            For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                If tmImportSpotInfo(ilImport).bMatched = False Then
                    ilRet = mAddIidf("I", 0, ilImport, "N")
                    grdMatchedResult.TextMatrix(llMatchRow, MRSTNCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(llMatchRow, MRSTNCOUNTINDEX)) + 1
                End If
            Next ilImport
        Else
            For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
                If tmImportSpotInfo(ilImport).bMatched = False Then
                    mBuildIidf "I", 0, ilImport, "N"
                End If
            Next ilImport
        End If
    End If
    If Not blPreviewMode Then
        llMatchRow = llMatchRow + 1
    End If
    mMatchSpots = True
End Function

Private Sub mClearIihfAndIidf(ilVefCode As Integer, llChfCode As Long, llInvStartdate As Long)
    Dim ilRet As Integer
    Dim slSchDate As String
    Dim ilGameNo As Integer
    
   
    tmIihfSrchKey2.lChfCode = llChfCode
    tmIihfSrchKey2.iVefCode = ilVefCode
    gPackDateLong llInvStartdate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    If UCase$(Trim$(tmIihf.sFileName)) <> UCase$(Trim$(smIihfFileName)) Then
        Exit Sub
    End If
    Do
        tmIidfSrchKey1.lCode = tmIihf.lCode
        ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        ilRet = btrDelete(hmIidf)
        If tmIidf.sSpotMatchType = "C" Then
            tmSdfSrchKey3.lCode = tmIidf.lSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                'Reset spot to missed
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                imSelectedDay = gWeekDayStr(slSchDate)
                ilGameNo = 0
                ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                tmSdfSrchKey3.lCode = tmIidf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    tmSdf.iDate(0) = tmIidf.iOrigSpotDate(0)
                    tmSdf.iDate(1) = tmIidf.iOrigSpotDate(1)
                    tmSdf.iTime(0) = tmIidf.iOrigSpotTime(0)
                    tmSdf.iTime(1) = tmIidf.iOrigSpotTime(1)
                    tmSdf.sPtType = "0"
                    tmSdf.lCopyCode = 0
                    tmSdf.iRotNo = 0
                    ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                End If
            End If
        End If
    Loop While ilRet = BTRV_ERR_NONE
    ilRet = btrDelete(hmIihf)
End Sub

Private Function mAddIihf(ilVefCode As Integer, llChfCode As Long, llAmfCode As Long, llInvStartdate As Long) As Integer
    Dim ilRet As Integer
    
    tmIihf.lCode = 0
    tmIihf.iVefCode = ilVefCode
    tmIihf.lChfCode = llChfCode
    gPackDateLong llInvStartdate, tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1)
    tmIihf.sFileName = smIihfFileName
    tmIihf.sStnEstimateNo = smEstimateNumber
    tmIihf.sStnInvoiceNo = smInvoiceNumber
    tmIihf.sStnContractNo = smContractNumber
    tmIihf.sStnAdvtName = smAdvertiserName
    tmIihf.lAmfCode = llAmfCode
    tmIihf.sSourceForm = smSourceForm
    tmIihf.sUnused = ""
    ilRet = btrInsert(hmIihf, tmIihf, imIihfRecLen, INDEXKEY0)
    mAddIihf = ilRet
End Function


Private Function mAddIidf(slType As String, ilSdfIndex As Integer, ilImportIndex As Integer, slAgyCompliant As String) As Integer
    Dim ilRet As Integer
    Dim tlSdf As SDF
    Dim tlImportSpotInfo As IMPORTSPOTINFO
    Dim llCifCode As Long
    
    tmIidf.lCode = 0
    tmIidf.lIihfCode = tmIihf.lCode
    tmIidf.sSpotMatchType = slType
    If slType = "C" Then
        tlSdf = tmNetSpotInfo(ilSdfIndex).tSdf
        tlImportSpotInfo = tmImportSpotInfo(ilImportIndex)
        tmIidf.lSdfCode = tlSdf.lCode
        gPackDateLong tlImportSpotInfo.lAirDate, tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTimeLong tlImportSpotInfo.lAirTime, tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = tlImportSpotInfo.iLen
        tmIidf.lStnCpfCode = mAddCpf(tlImportSpotInfo.sISCI)
        llCifCode = mAddOrUpdateCif(tmIidf.lStnCpfCode, tlSdf.iAdfCode, tmIidf.iStnSpotLen)
        tmIidf.lStnDPStartTime = tlImportSpotInfo.lDPStartTime
        tmIidf.lStnDPEndTime = tlImportSpotInfo.lDPEndTime
        tmIidf.sStnDPDays = tlImportSpotInfo.sDPDays
        tmIidf.iOrigSpotDate(0) = tlSdf.iDate(0)
        tmIidf.iOrigSpotDate(1) = tlSdf.iDate(1)
        tmIidf.iOrigSpotTime(0) = tlSdf.iTime(0)
        tmIidf.iOrigSpotTime(1) = tlSdf.iTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tmNetSpotInfo(ilSdfIndex).lAcqCost >= 0 Then
            tmIidf.sStnRate = gLongToStrDec(tmNetSpotInfo(ilSdfIndex).lAcqCost, 2)
        End If
    ElseIf slType = "I" Then
        tlImportSpotInfo = tmImportSpotInfo(ilImportIndex)
        tmIidf.lSdfCode = 0
        gPackDateLong tlImportSpotInfo.lAirDate, tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTimeLong tlImportSpotInfo.lAirTime, tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = tlImportSpotInfo.iLen
        'tmIidf.sStnISCI = tlImportSpotInfo.sISCI
        tmIidf.lStnCpfCode = mAddCpf(tlImportSpotInfo.sISCI)
        llCifCode = mAddOrUpdateCif(tmIidf.lStnCpfCode, 0, tmIidf.iStnSpotLen)
        tmIidf.lStnDPStartTime = tlImportSpotInfo.lDPStartTime
        tmIidf.lStnDPEndTime = tlImportSpotInfo.lDPEndTime
        tmIidf.sStnDPDays = tlImportSpotInfo.sDPDays
        gPackDate "1/1/1970", tmIidf.iOrigSpotDate(0), tmIidf.iOrigSpotDate(1)
        gPackTime "12AM", tmIidf.iOrigSpotTime(0), tmIidf.iOrigSpotTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tlImportSpotInfo.lRate >= 0 Then
            tmIidf.sStnRate = gLongToStrDec(tlImportSpotInfo.lRate, 2)
        End If
    Else
        tlSdf = tmNetSpotInfo(ilSdfIndex).tSdf
        tmIidf.lSdfCode = tlSdf.lCode
        gPackDate "1/1/1970", tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTime "12AM", tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = 0
        tmIidf.lStnCpfCode = 0
        tmIidf.lStnDPStartTime = 0
        tmIidf.lStnDPEndTime = 0
        tmIidf.sStnDPDays = ""
        tmIidf.iOrigSpotDate(0) = tlSdf.iDate(0)
        tmIidf.iOrigSpotDate(1) = tlSdf.iDate(1)
        tmIidf.iOrigSpotTime(0) = tlSdf.iTime(0)
        tmIidf.iOrigSpotTime(1) = tlSdf.iTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tmNetSpotInfo(ilSdfIndex).lAcqCost > 0 Then
            tmIidf.sStnRate = gLongToStrDec(tmNetSpotInfo(ilSdfIndex).lAcqCost, 2)
        End If
    End If
    tmIidf.sUnused = ""
    ilRet = btrInsert(hmIidf, tmIidf, imIidfRecLen, INDEXKEY0)
    mAddIidf = ilRet
    
End Function

Private Sub mBuildIidf(slType As String, ilSdfIndex As Integer, ilImportIndex As Integer, slAgyCompliant As String)
    Dim ilRet As Integer
    Dim tlSdf As SDF
    Dim tlImportSpotInfo As IMPORTSPOTINFO
    Dim llCifCode As Long
    
    tmIidf.lIihfCode = tmIihf.lCode
    tmIidf.sSpotMatchType = slType
    If slType = "C" Then
        tlSdf = tmNetSpotInfo(ilSdfIndex).tSdf
        tlImportSpotInfo = tmImportSpotInfo(ilImportIndex)
        tmIidf.lSdfCode = tlSdf.lCode
        gPackDateLong tlImportSpotInfo.lAirDate, tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTimeLong tlImportSpotInfo.lAirTime, tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = tlImportSpotInfo.iLen
        tmIidf.lStnCpfCode = mAddCpf(tlImportSpotInfo.sISCI)
        'llCifCode = mAddOrUpdateCif(tmIidf.lStnCpfCode, tlSdf.iadfCode, tmIidf.iStnSpotLen)
        tmIidf.lStnDPStartTime = tlImportSpotInfo.lDPStartTime
        tmIidf.lStnDPEndTime = tlImportSpotInfo.lDPEndTime
        tmIidf.sStnDPDays = tlImportSpotInfo.sDPDays
        tmIidf.iOrigSpotDate(0) = tlSdf.iDate(0)
        tmIidf.iOrigSpotDate(1) = tlSdf.iDate(1)
        tmIidf.iOrigSpotTime(0) = tlSdf.iTime(0)
        tmIidf.iOrigSpotTime(1) = tlSdf.iTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tmNetSpotInfo(ilSdfIndex).lAcqCost >= 0 Then
            tmIidf.sStnRate = gLongToStrDec(tmNetSpotInfo(ilSdfIndex).lAcqCost, 2)
        End If
    ElseIf slType = "I" Then
        tlImportSpotInfo = tmImportSpotInfo(ilImportIndex)
        tmIidf.lSdfCode = 0
        gPackDateLong tlImportSpotInfo.lAirDate, tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTimeLong tlImportSpotInfo.lAirTime, tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = tlImportSpotInfo.iLen
        'tmIidf.sStnISCI = tlImportSpotInfo.sISCI
        tmIidf.lStnCpfCode = mAddCpf(tlImportSpotInfo.sISCI)
        'llCifCode = mAddOrUpdateCif(tmIidf.lStnCpfCode, 0, tmIidf.iStnSpotLen)
        tmIidf.lStnDPStartTime = tlImportSpotInfo.lDPStartTime
        tmIidf.lStnDPEndTime = tlImportSpotInfo.lDPEndTime
        tmIidf.sStnDPDays = tlImportSpotInfo.sDPDays
        gPackDate "1/1/1970", tmIidf.iOrigSpotDate(0), tmIidf.iOrigSpotDate(1)
        gPackTime "12AM", tmIidf.iOrigSpotTime(0), tmIidf.iOrigSpotTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tlImportSpotInfo.lRate >= 0 Then
            tmIidf.sStnRate = gLongToStrDec(tlImportSpotInfo.lRate, 2)
        End If
    Else
        tlSdf = tmNetSpotInfo(ilSdfIndex).tSdf
        tmIidf.lSdfCode = tlSdf.lCode
        gPackDate "1/1/1970", tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
        gPackTime "12AM", tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
        tmIidf.iStnSpotLen = 0
        tmIidf.lStnCpfCode = 0
        tmIidf.lStnDPStartTime = 0
        tmIidf.lStnDPEndTime = 0
        tmIidf.sStnDPDays = ""
        tmIidf.iOrigSpotDate(0) = tlSdf.iDate(0)
        tmIidf.iOrigSpotDate(1) = tlSdf.iDate(1)
        tmIidf.iOrigSpotTime(0) = tlSdf.iTime(0)
        tmIidf.iOrigSpotTime(1) = tlSdf.iTime(1)
        tmIidf.sAgyCompliant = slAgyCompliant
        tmIidf.sStnRate = ""
        If tmNetSpotInfo(ilSdfIndex).lAcqCost > 0 Then
            tmIidf.sStnRate = gLongToStrDec(tmNetSpotInfo(ilSdfIndex).lAcqCost, 2)
        End If
    End If
    tmIidf.sUnused = ""
    tmIidfDetail(UBound(tmIidfDetail)) = tmIidf
    ReDim Preserve tmIidfDetail(0 To UBound(tmIidfDetail) + 1) As IIDF
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mFindAvail                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Private Function mFindAvail(ilVefCode As Integer, slSchDate As String, slFindTime As String, ilGameNo As Integer, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slSchTime, ilAvailIndex)
'   Where:
'       slSchDate(I)- Scheduled Date
'       slSchTime(I)- Time that avail is to be found at
'       ilFindAdjAvail(I)- Find closest avail to specified time
'       llSsfRecPos(O)- Ssf record position
'       ilAvailIndex(O)- Index into Ssf where avail is located
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim llSchDate As Long
    Dim llTime As Long
    Dim llTstTime As Long
    Dim llFndAdjTime As Long
    Dim ilLoop As Integer
    llTime = CLng(gTimeToCurrency(slFindTime, False))
    llSchDate = gDateValue(slSchDate)
    imSelectedDay = gWeekDayStr(slSchDate)
    lmSsfDate(imSelectedDay) = 0
    ilRet = gObtainSsfForDateOrGame(ilVefCode, llSchDate, slFindTime, ilGameNo, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTstTime
            If llTime = llTstTime Then 'Replace
                ilAvailIndex = ilLoop
                mFindAvail = True
                Exit Function
            ElseIf (llTstTime < llTime) And (ilFindAdjAvail) Then
                ilAvailIndex = ilLoop
                llFndAdjTime = llTstTime
            ElseIf (llTime < llTstTime) And (ilFindAdjAvail) Then
                If llFndAdjTime = -1 Then
                    ilAvailIndex = ilLoop
                    mFindAvail = True
                    Exit Function
                Else
                    If (llTime - llFndAdjTime) < (llTstTime - llTime) Then
                        mFindAvail = True
                        Exit Function
                    Else
                        ilAvailIndex = ilLoop
                        mFindAvail = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    If (llFndAdjTime <> -1) And (ilFindAdjAvail) Then
        mFindAvail = True
        Exit Function
    End If
    mFindAvail = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailRoom                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if room exist for    *
'*                      spot within avail              *
'*                                                     *
'*******************************************************
Private Function mAvailRoom(ilVefCode As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mAvailRoom(ilAvailIndex)
'   where:
'       ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'       ilRet(O)- True=Avail has room; False=insufficient room within avail
'
'       tmSdf(I)- spot records
'
'       Code later: ask if avail should be overbooked
'                   If so, create a version zero (0) of the library with the new
'                   units/seconds
'
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
    
    
    imVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If imVpfIndex = -1 Then
        mAvailRoom = False
        Exit Function
    End If
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    '10/27/11: Disallow more then 31 spots in any avail
    If tmAvail.iNoSpotsThis >= 31 Then
        'ilRet = MsgBox("Move not allowed because Avail contains the maximum number of spots (31).", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
       LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            mAvailRoom = True
            Exit Function
        End If
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(imVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
                'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
                '    ilSpotLen = 0
                'End If
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
        End If
    Next ilSpotIndex
    ilSpotLen = tmSdf.iLen
    If (tgVpf(imVpfIndex).sSSellOut = "T") Then
        ilSpotUnits = ilSpotLen \ 30
        If ilSpotUnits <= 0 Then
            ilSpotUnits = 1
        End If
        ilSpotLen = 0
    Else
        ilSpotUnits = 1
        'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
        '    ilSpotLen = 0
        'End If
    End If
    ilNewUnit = 0
    ilNewSec = 0
    If (tgVpf(imVpfIndex).sSSellOut = "M") Then
        If (ilSpotLen + ilSecSold <> ilAvailSec) Or (ilSpotUnits + ilUnitsSold <> ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    Else
        If (ilSpotLen + ilSecSold > ilAvailSec) Or (ilSpotUnits + ilUnitsSold > ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    End If
    If (tgVpf(imVpfIndex).sSOverBook <> "Y") Then
        'ilRet = MsgBox("Move not allowed because Avail would be Overbooked.", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        '5/20/11
        If (tmAvail.iOrigUnit = 0) And (tmAvail.iOrigLen = 0) Then
            tmAvail.iOrigUnit = tmAvail.iAvInfo And &H1F
            tmAvail.iOrigLen = tmAvail.iLen
        End If
        tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        tmAvail.iLen = ilNewSec
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailRoom = False
        Exit Function
    End If
    mAvailRoom = True
    Exit Function
End Function

Private Function mBookSpot(ilSdfIndex As Integer, ilImportIndex As Integer, slInAgyCompliant As String, Optional blPreviewMode As Boolean = False) As Boolean
    Dim tlSdf As SDF
    Dim tlImportSpotInfo As IMPORTSPOTINFO
    Dim slAirDate As String
    Dim slAirTime As String
    Dim ilAvailIndex As Integer
    Dim ilBkQH As Integer
    Dim ilRet As Integer
    Dim llSdfRecPos As Long
    Dim slRet As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilPriceLevel As Integer
    Dim llCpfCode As Long
    Dim llCifCode As Long
    Dim slSdfDate As String
    Dim slAgyCompliant As String
    
    mBookSpot = False
    If blPreviewMode Then
        mBookSpot = True
        Exit Function
    End If
    slAgyCompliant = slInAgyCompliant
    tlSdf = tmNetSpotInfo(ilSdfIndex).tSdf
    tlImportSpotInfo = tmImportSpotInfo(ilImportIndex)
    ilVefCode = tlSdf.iVefCode
    ilVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If imVpfIndex = -1 Then
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    slAirDate = Format(tlImportSpotInfo.lAirDate, "m/d/yy")
    slAirTime = gFormatTimeLong(tlImportSpotInfo.lAirTime, "A", "1")
    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    If Not mAvailRoom(ilVefCode, ilAvailIndex) Then
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    tmSdfSrchKey3.lCode = tlSdf.lCode
    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
    If slAgyCompliant = "O" Then  'A=Aired as Sold; O=Aired Outside
        slRet = "O"
    Else
        slRet = "S"
        gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slSdfDate
        If gDateValue(gObtainPrevMonday(slAirDate)) <> gDateValue(gObtainPrevMonday(slSdfDate)) Then
            slRet = "O"
            slAgyCompliant = "O"
        End If
    End If
    'Test if time within daypart, if not set to Outside
    ilBkQH = IMPORTINVOICESPOT
    ilPriceLevel = 0
    ilRet = gBookSpot(slRet, hmSdf, tlSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmRdf, ilVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmSxf, hmGsf)
    mBookSpot = ilRet
    If ilRet Then
        ilRet = mAddIidf("C", ilSdfIndex, ilImportIndex, slAgyCompliant)
    Else
        'ilRet = mAddIidf("I", ilSdfIndex, ilImportIndex, "N")
        'ilRet = mAddIidf("M", ilSdfIndex, ilImportIndex, "N")
        Exit Function
    End If
    tmSdfSrchKey3.lCode = tlSdf.lCode
    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        If gDateValue(slAirDate) <= lmLastStdMnthBilled Then
            tlSdf.sBill = "Y"
        End If
        tlSdf.sAffChg = "Y"
        gPackTimeLong tlImportSpotInfo.lAirTime, tlSdf.iTime(0), tlSdf.iTime(1)
        llCpfCode = mAddCpf(tlImportSpotInfo.sISCI)
        llCifCode = mAddOrUpdateCif(llCpfCode, tlSdf.iAdfCode, tlSdf.iLen)
        If llCifCode > 0 Then
            tlSdf.sPtType = "1"
            tlSdf.lCopyCode = llCifCode
        Else
            tlSdf.sPtType = "0"
            tlSdf.lCopyCode = 0
            tlSdf.iRotNo = 0
        End If
        ilRet = btrUpdate(hmSdf, tlSdf, imSdfRecLen)
    End If
End Function



Private Function mGetISCI(tlSdf As SDF) As String
    Dim ilRet As Integer
    
    mGetISCI = ""
    If tlSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.lcpfCode > 0 Then
                tmCpfSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    mGetISCI = Trim$(tmCpf.sISCI)
                End If
            End If
        End If
    End If

End Function



Private Sub mPopDetail(Optional blPreviewMode As Boolean = False)
    Dim ilRet As Integer
    Dim llIihfCode As Long
    Dim llSpotRow As Long
    Dim llStationRow As Long
    Dim llMatchRow As Long
    Dim slDays As String
    Dim slTimeRange As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slDateRange As String
    Dim slEDIDays As String
    Dim slAcqRate As String
    Dim ilIidf As Integer
    
    mClearGrid grdNetworkSpots
    mClearGrid grdStationSpots
    mClearGrid grdMatchedSpots
    
    llSpotRow = grdNetworkSpots.FixedRows
    llStationRow = grdStationSpots.FixedRows
    llMatchRow = grdMatchedSpots.FixedRows
    If Not blPreviewMode Then
        ReDim tmIidfDetail(0 To 0) As IIDF
        lacDetail.Caption = "Details for: " & grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETADVERTISERINDEX) & ", Invoice from " & grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNSTATIONINDEX)
        llIihfCode = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRIIHFCODEINDEX))
        tmIihfSrchKey0.lCode = llIihfCode
        ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        tmIidfSrchKey1.lCode = llIihfCode
        ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmIidf.lIihfCode = llIihfCode)
            tmIidfDetail(UBound(tmIidfDetail)) = tmIidf
            ReDim Preserve tmIidfDetail(0 To UBound(tmIidfDetail) + 1) As IIDF
            ilRet = btrGetNext(hmIidf, tmIidf, imIidfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        lacDetail.Caption = "Details for: " & grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRADVERTISERINDEX) & ", Invoice from " & grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRSTATIONINDEX) & " Match: " & grdCntrNetwork.TextMatrix(lmCRowSelected, CADVERTISERINDEX) & " Contract: " & grdCntrNetwork.TextMatrix(lmCRowSelected, CCONTRACTINDEX)
        llIihfCode = Val(grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRIIHFCODEINDEX))
    End If
    For ilIidf = 0 To UBound(tmIidfDetail) - 1 Step 1
        tmIidf = tmIidfDetail(ilIidf)
        If (tmIidf.sSpotMatchType = "M") Or (tmIidf.sSpotMatchType = "C") Then
            tmSdfSrchKey3.lCode = tmIidf.lSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            mDetermineDP slDays, slTimeRange, slDateRange, slEDIDays, slAcqRate
        End If
        If tmIidf.sSpotMatchType = "M" Then
            If llSpotRow >= grdNetworkSpots.Rows Then
                grdNetworkSpots.AddItem ""
            End If
            grdNetworkSpots.RowHeight(llSpotRow) = fgBoxGridH + 15
            grdNetworkSpots.TextMatrix(llSpotRow, NDPDAYSINDEX) = slDays
            grdNetworkSpots.TextMatrix(llSpotRow, NDPTIMEINDEX) = slTimeRange
            grdNetworkSpots.TextMatrix(llSpotRow, NLENGTHINDEX) = tmSdf.iLen
            grdNetworkSpots.TextMatrix(llSpotRow, NACQRATEINDEX) = slAcqRate
            grdNetworkSpots.TextMatrix(llSpotRow, NDATESINDEX) = slDateRange
            grdNetworkSpots.TextMatrix(llSpotRow, NSDFCODEINDEX) = tmSdf.lCode
            llSpotRow = llSpotRow + 1
        ElseIf tmIidf.sSpotMatchType = "I" Then
            If llStationRow >= grdStationSpots.Rows Then
                grdStationSpots.AddItem ""
            End If
            grdStationSpots.RowHeight(llStationRow) = fgBoxGridH + 15
            gUnpackDate tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), slAirDate
            gUnpackTime tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), "A", "4", slAirTime
            grdStationSpots.TextMatrix(llStationRow, SDATEINDEX) = slAirDate
            grdStationSpots.TextMatrix(llStationRow, STIMEINDEX) = slAirTime
            grdStationSpots.TextMatrix(llStationRow, SLENGTHINDEX) = tmIidf.iStnSpotLen
            grdStationSpots.TextMatrix(llStationRow, SACQRATEINDEX) = Trim$(tmIidf.sStnRate)
            grdStationSpots.TextMatrix(llStationRow, SISCIINDEX) = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'Trim$(tmIidf.sStnISCI)
            grdStationSpots.TextMatrix(llStationRow, SIIDFCODEINDEX) = tmIidf.lCode
            llStationRow = llStationRow + 1
        Else
            If llMatchRow >= grdMatchedSpots.Rows Then
                grdMatchedSpots.AddItem ""
            End If
            grdMatchedSpots.RowHeight(llMatchRow) = fgBoxGridH + 15
            grdMatchedSpots.TextMatrix(llMatchRow, MNETDPDAYSINDEX) = slDays
            grdMatchedSpots.TextMatrix(llMatchRow, MNETDPTIMEINDEX) = slTimeRange
            grdMatchedSpots.TextMatrix(llMatchRow, MNETDATESINDEX) = slDateRange
            gUnpackDate tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), slAirDate
            gUnpackTime tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), "A", "4", slAirTime
            grdMatchedSpots.TextMatrix(llMatchRow, MSTNDATEINDEX) = slAirDate
            grdMatchedSpots.TextMatrix(llMatchRow, MSTNTIMEINDEX) = slAirTime
            If tmIidf.sAgyCompliant = "O" Then  'A=Aired as Sold; O=Aired Outside
                grdMatchedSpots.TextMatrix(llMatchRow, MSTNCOMPLIANTINDEX) = "No"
            Else
                grdMatchedSpots.TextMatrix(llMatchRow, MSTNCOMPLIANTINDEX) = ""
            End If
            grdMatchedSpots.TextMatrix(llMatchRow, MLENGTHINDEX) = tmIidf.iStnSpotLen
            grdMatchedSpots.TextMatrix(llMatchRow, MACQRATEINDEX) = slAcqRate
            grdMatchedSpots.TextMatrix(llMatchRow, MSTNISCIINDEX) = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'Trim$(tmIidf.sStnISCI)
            grdMatchedSpots.TextMatrix(llMatchRow, MSDFCODEINDEX) = tmSdf.lCode
            grdMatchedSpots.TextMatrix(llMatchRow, MIIDFCODEINDEX) = tmIidf.lCode
            grdMatchedSpots.TextMatrix(llMatchRow, MSSORTINDEX) = ""
            llMatchRow = llMatchRow + 1
        End If
    Next ilIidf
    mMSSortCol MSTNTIMEINDEX
    mMSSortCol MSTNDATEINDEX
    mMSSortCol MNETDPTIMEINDEX
End Sub

Private Sub mClearGrid(grdGrid As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    grdGrid.TopRow = grdGrid.FixedRows
    grdGrid.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdGrid.FixedRows To grdGrid.Rows - 1 Step 1
        grdGrid.Row = llRow
        grdGrid.RowHeight(llRow) = fgBoxGridH + 15
        For llCol = 0 To grdGrid.Cols - 1 Step 1
            grdGrid.Col = llCol
            grdGrid.CellBackColor = vbWhite
            grdGrid.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdGrid.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub


Private Function mAddiihfAndIidf(ilVefCode As Integer, llChfCode As Long, llAmfCode As Long, llInvStartdate As Long, llNoMatchRow As Long, slStatus As String) As Integer
    Dim ilRet As Integer
    Dim ilImport As Integer
    
    ''Set result to File does not exist in grid
    'If llNoMatchRow >= grdNoMatchedResult.Rows Then
    '    grdNoMatchedResult.AddItem ""
    'End If
    ilRet = mAddIihf(ilVefCode, llChfCode, llAmfCode, llInvStartdate)
    If ilRet <> BTRV_ERR_NONE Then
        mAddToNoMatchGrid llNoMatchRow, 0, "Unable to Add Import Header File, ilRet = " & ilRet
        mAddiihfAndIidf = False
        Exit Function
    End If
    mAddToNoMatchGrid llNoMatchRow, tmIihf.lCode, slStatus
    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
        ilRet = mAddIidf("I", -1, ilImport, "N")
    Next ilImport
    'llNoMatchRow = llNoMatchRow + 1
    mAddiihfAndIidf = True
End Function

Private Sub mAddToNoMatchGrid(llNoMatchRow As Long, llIihfCode As Long, slStatus As String)
    Dim ilLen As Integer
    Dim blLenFound As Boolean
    Dim ilImport As Integer
    Dim ilCount As Integer
    
    If llNoMatchRow >= grdNoMatchedResult.Rows Then
        grdNoMatchedResult.AddItem ""
    End If
    grdNoMatchedResult.RowHeight(llNoMatchRow) = fgBoxGridH + 15
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRSTATIONINDEX) = smCallLetters
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRADVERTISERINDEX) = smAdvertiserName
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRESTIMATEINDEX) = smEstimateNumber
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCONTRACTINDEX) = smContractNumber
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRINVOICEINDEX) = smInvoiceNumber
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRFILENAMEINDEX) = smIihfFileName
    ReDim ilStnLen(0 To 0) As Integer
    For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
        blLenFound = False
        For ilLen = 0 To UBound(ilStnLen) - 1 Step 1
            If ilStnLen(ilLen) = tmImportSpotInfo(ilImport).iLen Then
                blLenFound = True
                Exit For
            End If
        Next ilLen
        If Not blLenFound Then
            ilStnLen(UBound(ilStnLen)) = tmImportSpotInfo(ilImport).iLen
            ReDim Preserve ilStnLen(0 To UBound(ilStnLen) + 1) As Integer
        End If
    Next ilImport
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) = ""
    For ilLen = 0 To UBound(ilStnLen) - 1 Step 1
        ilCount = 0
        For ilImport = 0 To UBound(tmImportSpotInfo) - 1 Step 1
            If tmImportSpotInfo(ilImport).iLen = ilStnLen(ilLen) Then
                ilCount = ilCount + 1
            End If
        Next ilImport
        If grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) = "" Then
            grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) = ilStnLen(ilLen) & "s: " & ilCount
        Else
            grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) = grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRCOUNTINDEX) & "; " & ilStnLen(ilLen) & "s: " & ilCount
        End If
    Next ilLen
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRSTATUSINDEX) = slStatus
    grdNoMatchedResult.TextMatrix(llNoMatchRow, NMRIIHFCODEINDEX) = llIihfCode
    llNoMatchRow = llNoMatchRow + 1
End Sub
Private Sub mAddToMatchGrid(llMatchRow As Long, llIihfCode As Long)
    Dim ilAdf As Integer
    Dim slDPName As String
    Dim slEDIDays As String
    Dim slDays(0 To 6) As String
    
    If llMatchRow >= grdMatchedResult.Rows Then
        grdMatchedResult.AddItem ""
    End If
    grdMatchedResult.RowHeight(llMatchRow) = fgBoxGridH + 15
    ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
    If ilAdf <> -1 Then
        grdMatchedResult.TextMatrix(llMatchRow, MRNETADVERTISERINDEX) = Trim$(tgCommAdf(ilAdf).sName)
    Else
        grdMatchedResult.TextMatrix(llMatchRow, MRNETADVERTISERINDEX) = ""
    End If
    grdMatchedResult.TextMatrix(llMatchRow, MRNETESTIMATEINDEX) = Trim$(tmChf.sAgyEstNo) & Trim$(tmChf.sTitle)
    grdMatchedResult.TextMatrix(llMatchRow, MRNETCONTRACTINDEX) = tmChf.lCntrNo
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNSTATIONINDEX) = smCallLetters
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNADVERTISERINDEX) = smAdvertiserName
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNESTIMATEINDEX) = smEstimateNumber
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNCONTRACTINDEX) = smContractNumber
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNINVOICEINDEX) = smInvoiceNumber
    grdMatchedResult.TextMatrix(llMatchRow, MRSTATUSINDEX) = ""
    grdMatchedResult.TextMatrix(llMatchRow, MRMATCHCOUNTINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchRow, MRNETCOUNTINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchRow, MRSTNCOUNTINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchRow, MRCOMPLIANTINDEX) = ""
    grdMatchedResult.TextMatrix(llMatchRow, MRFILENAMEINDEX) = smIihfFileName
    grdMatchedResult.TextMatrix(llMatchRow, MRSELECTEDINDEX) = "0"
    grdMatchedResult.TextMatrix(llMatchRow, MRIIHFCODEINDEX) = llIihfCode
    grdMatchedResult.TextMatrix(llMatchRow, MRCHFCODEINDEX) = tmChf.lCode
    grdMatchedResult.TextMatrix(llMatchRow, MRSORTINDEX) = ""
End Sub


Private Sub mDetermineDP(slDays As String, slTimeRange As String, slDateRange As String, slEDIDays As String, slAcqRate As String)
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim llSdfDate As Long
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilDay As Integer
    Dim ilDP As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim llLnStartDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    
    slDays = ""
    slTimeRange = ""
    slDateRange = ""
    slAcqRate = ""
    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
    tmChfSrchKey0.lCode = tmSdf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmClfSrchKey0.lChfCode = tmSdf.lChfCode
        tmClfSrchKey0.iLine = tmSdf.iLineNo
        tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnStartDate
        ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRdf <> -1 Then
            tmRdf = tgMRdf(ilRdf)
        End If
        If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
            ReDim llStartTime(0 To 1) As Long
            ReDim llEndTime(0 To 1) As Long
            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "4", slStartTime
            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "4", slEndTime
            slTimeRange = slStartTime & "-" & slEndTime
        Else
            ReDim llStartTime(0 To 0) As Long
            ReDim llEndTime(0 To 0) As Long
            If ilRdf <> -1 Then
                For ilDP = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                    If (tmRdf.iStartTime(0, ilDP) <> 1) Or (tmRdf.iStartTime(1, ilDP) <> 0) Then
                        gUnpackTime tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), "A", "4", slStartTime
                        gUnpackTime tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), "A", "4", slEndTime
                        gUnpackTimeLong tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), False, llRdfStartTime
                        gUnpackTimeLong tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), True, llRdfEndTime
                        If llRdfStartTime < llRdfEndTime Then
                            If slTimeRange = "" Then
                                slTimeRange = slStartTime & "-" & slEndTime
                            Else
                                slTimeRange = slTimeRange & "; " & slStartTime & "-" & slEndTime
                            End If
                        Else
                            If slTimeRange = "" Then
                                slTimeRange = slStartTime & "-" & "12AM"
                            Else
                                slTimeRange = slTimeRange & "; " & slStartTime & "-" & "12AM"
                            End If
                            slTimeRange = slTimeRange & "; " & "12AM" & "-" & slEndTime
                        End If
                    End If
                Next ilDP
            End If
        End If
        slAcqRate = gLongToStrDec(tmClf.lAcquisitionCost, 2)
        tmCffSrchKey0.lChfCode = tmClf.lChfCode
        tmCffSrchKey0.iClfLine = tmClf.iLine
        tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
        tmCffSrchKey0.iPropVer = tmClf.iPropVer
        gPackDateLong llLnStartDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmCff.iClfLine = tmClf.iLine) And (tmCff.lChfCode = tmClf.lChfCode)
            gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llCffStartDate
            gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llCffEndDate
            If (llCffStartDate <= llSdfDate) And (llCffEndDate >= llSdfDate) Then
                'slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                If tmCff.sDyWk <> "W" Then
                    For ilDay = 0 To 6 Step 1
                        If ilDay <> gWeekDayLong(llSdfDate) Then
                            tmCff.iDay(ilDay) = 0
                        End If
                    Next ilDay
                    slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                    slStartDate = Format(llSdfDate, "m/d/yy")
                    slEndDate = slStartDate
                Else
                    slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                    For ilDay = 0 To 6 Step 1
                        If tmCff.iDay(ilDay) > 0 Then
                            slStartDate = Format(llSdfDate - (gWeekDayLong(llSdfDate) - ilDay), "m/d/yy")
                            Exit For
                        End If
                    Next ilDay
                    For ilDay = 6 To 0 Step -1
                        If tmCff.iDay(ilDay) > 0 Then
                            slEndDate = Format(llSdfDate + (ilDay - gWeekDayLong(llSdfDate)), "m/d/yy")
                            Exit For
                        End If
                    Next ilDay
                End If
                If gDateValue(slStartDate) = gDateValue(slEndDate) Then
                    slDateRange = slStartDate
                Else
                    slDateRange = slStartDate & "-" & slEndDate
                End If
                Exit Sub
            End If
            ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
End Sub

Private Sub mUndoReconcile()
    Dim llIidfCode As Long
    Dim llSdfCode As Long
    Dim ilGameNo As Integer
    Dim ilRet As Integer
    Dim slSchDate As String
    Dim llRow As Long
    Dim llSpotRow As Long
    Dim llStationRow As Long
    Dim slDays As String
    Dim slTimeRange As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slDateRange As String
    Dim slEDIDays As String
    Dim slAcqRate As String
    
    If lmMRowSelected < grdMatchedSpots.FixedRows Then
        Exit Sub
    End If
    llIidfCode = Val(grdMatchedSpots.TextMatrix(lmMRowSelected, MIIDFCODEINDEX))
    tmIidfSrchKey0.lCode = llIidfCode
    ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    llSdfCode = Val(grdMatchedSpots.TextMatrix(lmMRowSelected, MSDFCODEINDEX))
    tmSdfSrchKey3.lCode = tmIidf.lSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        'Reset spot to missed
        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
        imSelectedDay = gWeekDayStr(slSchDate)
        ilGameNo = 0
        ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
        tmSdfSrchKey3.lCode = llSdfCode
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmSdf.iDate(0) = tmIidf.iOrigSpotDate(0)
            tmSdf.iDate(1) = tmIidf.iOrigSpotDate(1)
            tmSdf.iTime(0) = tmIidf.iOrigSpotTime(0)
            tmSdf.iTime(1) = tmIidf.iOrigSpotTime(1)
            tmSdf.sPtType = "0"
            tmSdf.lCopyCode = 0
            tmSdf.iRotNo = 0
            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
        End If
    End If
    tmIidf.sSpotMatchType = "I"
    tmIidf.lSdfCode = 0
    gPackDate "1/1/1970", tmIidf.iOrigSpotDate(0), tmIidf.iOrigSpotDate(1)
    gPackTime "12AM", tmIidf.iOrigSpotTime(0), tmIidf.iOrigSpotTime(1)
    tmIidf.sAgyCompliant = "N"
    ilRet = btrUpdate(hmIidf, tmIidf, imIidfRecLen)
    
    'Add to Unreconciled Station Spots
    llStationRow = grdStationSpots.FixedRows
    For llRow = grdStationSpots.FixedRows To grdStationSpots.Rows - 1 Step 1
        If grdStationSpots.TextMatrix(llRow, SDATEINDEX) <> "" Then
            llStationRow = llRow + 1
        End If
    Next llRow
    If llStationRow >= grdStationSpots.Rows Then
        grdStationSpots.AddItem ""
    End If
    grdStationSpots.RowHeight(llStationRow) = fgBoxGridH + 15
    gUnpackDate tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), slAirDate
    gUnpackTime tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), "A", "4", slAirTime
    grdStationSpots.TextMatrix(llStationRow, SDATEINDEX) = slAirDate
    grdStationSpots.TextMatrix(llStationRow, STIMEINDEX) = slAirTime
    grdStationSpots.TextMatrix(llStationRow, SLENGTHINDEX) = tmIidf.iStnSpotLen
    grdStationSpots.TextMatrix(llStationRow, SACQRATEINDEX) = tmIidf.sStnRate
    grdStationSpots.TextMatrix(llStationRow, SISCIINDEX) = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'Trim$(tmIidf.sStnISCI)
    grdStationSpots.TextMatrix(llStationRow, SIIDFCODEINDEX) = tmIidf.lCode
        
    tmIidf.lCode = 0
    tmIidf.lIihfCode = tmIidf.lIihfCode
    tmIidf.sSpotMatchType = "M"
    tmIidf.lSdfCode = tmSdf.lCode
    gPackDate "1/1/1970", tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1)
    gPackTime "12AM", tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1)
    tmIidf.lStnCpfCode = 0
    tmIidf.iOrigSpotDate(0) = tmSdf.iDate(0)
    tmIidf.iOrigSpotDate(1) = tmSdf.iDate(1)
    tmIidf.iOrigSpotTime(0) = tmSdf.iTime(0)
    tmIidf.iOrigSpotTime(1) = tmSdf.iTime(1)
    tmIidf.sAgyCompliant = "N"
    tmIidf.sStnRate = ""
    tmIidf.sUnused = ""
    ilRet = btrInsert(hmIidf, tmIidf, imIidfRecLen, INDEXKEY0)
    
    llSpotRow = grdNetworkSpots.FixedRows
    For llRow = grdNetworkSpots.FixedRows To grdNetworkSpots.Rows - 1 Step 1
        If grdNetworkSpots.TextMatrix(llRow, NDPDAYSINDEX) <> "" Then
            llSpotRow = llRow + 1
        End If
    Next llRow
    'Add to Unreconciled Network Spots
    If llSpotRow >= grdNetworkSpots.Rows Then
        grdNetworkSpots.AddItem ""
    End If
    mDetermineDP slDays, slTimeRange, slDateRange, slEDIDays, slAcqRate
    grdNetworkSpots.RowHeight(llSpotRow) = fgBoxGridH + 15
    grdNetworkSpots.TextMatrix(llSpotRow, NDPDAYSINDEX) = slDays
    grdNetworkSpots.TextMatrix(llSpotRow, NDPTIMEINDEX) = slTimeRange
    grdNetworkSpots.TextMatrix(llSpotRow, NLENGTHINDEX) = tmSdf.iLen
    grdNetworkSpots.TextMatrix(llSpotRow, NACQRATEINDEX) = slAcqRate
    grdNetworkSpots.TextMatrix(llSpotRow, NDATESINDEX) = slDateRange
    grdNetworkSpots.TextMatrix(llSpotRow, NSDFCODEINDEX) = tmSdf.lCode
    
    'Remove from grdMatchedSpots
    grdMatchedSpots.RemoveItem lmMRowSelected
    
    'Change counts in grdMatchedResult
    If lmMRRowSelected >= grdMatchedResult.FixedRows Then
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRMATCHCOUNTINDEX)) - 1
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETCOUNTINDEX)) + 1
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCOUNTINDEX)) + 1
    End If

    lmMRowSelected = -1
    mDetailSetCommands

End Sub

Private Sub mReconcile()
    Dim llSdfCode As Long
    Dim llIidfCode As Long
    Dim llRow As Long
    Dim ilRet As Integer
    Dim llMatchedRow As Long
    Dim slDays As String
    Dim slTimeRange As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slDateRange As String
    Dim slEDIDays As String
    Dim slAcqRate As String
    Dim slAgyCompliant As String
    Dim ilPos As Integer
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDPTime As Integer
    Dim slDPTimes() As String
    
    If lmNRowSelected < grdNetworkSpots.FixedRows Then
        Exit Sub
    End If
    If lmSRowSelected < grdStationSpots.FixedRows Then
        Exit Sub
    End If
    llSdfCode = grdNetworkSpots.TextMatrix(lmNRowSelected, NSDFCODEINDEX)
    tmSdfSrchKey3.lCode = llSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    llIidfCode = Val(grdStationSpots.TextMatrix(lmSRowSelected, SIIDFCODEINDEX))
    tmIidfSrchKey0.lCode = llIidfCode
    ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ReDim tmImportSpotInfo(0 To 1) As IMPORTSPOTINFO
    gUnpackDateLong tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), tmImportSpotInfo(0).lAirDate
    gUnpackTimeLong tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), False, tmImportSpotInfo(0).lAirTime
    tmImportSpotInfo(0).iLen = tmIidf.iStnSpotLen
    tmImportSpotInfo(0).sISCI = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'tmIidf.sStnISCI
    ilRet = btrDelete(hmIidf)
    
    tmIidfSrchKey2.lCode = tmSdf.lCode
    ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmIidf.lSdfCode = tmSdf.lCode)
        If tmIidf.sSpotMatchType = "M" Then
            ilRet = btrDelete(hmIidf)
            Exit Do
        End If
        ilRet = btrGetNext(hmIidf, tmIidf, imIidfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    slAgyCompliant = mGetCompliantStatus(tmImportSpotInfo(0).lAirDate, tmImportSpotInfo(0).lAirTime)
    
    ReDim tmNetSpotInfo(0 To 1) As NETSPOTINFO
    tmNetSpotInfo(0).tSdf = tmSdf
    
    ilRet = mBookSpot(0, 0, slAgyCompliant)
    
    llMatchedRow = grdMatchedSpots.FixedRows
    For llRow = grdMatchedSpots.FixedRows To grdMatchedSpots.Rows - 1 Step 1
        If grdMatchedSpots.TextMatrix(llRow, NDPDAYSINDEX) <> "" Then
            llMatchedRow = llRow + 1
        End If
    Next llRow
    'Add to Unreconciled Network Spots
    If llMatchedRow >= grdMatchedSpots.Rows Then
        grdMatchedSpots.AddItem ""
    End If
    grdMatchedSpots.RowHeight(llMatchedRow) = fgBoxGridH + 15
    mDetermineDP slDays, slTimeRange, slDateRange, slEDIDays, slAcqRate
    grdMatchedSpots.TextMatrix(llMatchedRow, MNETDPDAYSINDEX) = slDays
    grdMatchedSpots.TextMatrix(llMatchedRow, MNETDPTIMEINDEX) = slTimeRange
    grdMatchedSpots.TextMatrix(llMatchedRow, MNETDATESINDEX) = slDateRange
    gUnpackDate tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), slAirDate
    gUnpackTime tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), "A", "4", slAirTime
    grdMatchedSpots.TextMatrix(llMatchedRow, MSTNDATEINDEX) = slAirDate
    grdMatchedSpots.TextMatrix(llMatchedRow, MSTNTIMEINDEX) = slAirTime
    If tmIidf.sAgyCompliant = "O" Then  'A=Aired as Sold; O=Aired Outside
        grdMatchedSpots.TextMatrix(llMatchedRow, MSTNCOMPLIANTINDEX) = "No"
    Else
        grdMatchedSpots.TextMatrix(llMatchedRow, MSTNCOMPLIANTINDEX) = ""
    End If
    grdMatchedSpots.TextMatrix(llMatchedRow, MLENGTHINDEX) = tmIidf.iStnSpotLen
    grdMatchedSpots.TextMatrix(llMatchedRow, MACQRATEINDEX) = slAcqRate
    grdMatchedSpots.TextMatrix(llMatchedRow, MSTNISCIINDEX) = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'Trim$(tmIidf.sStnISCI)
    grdMatchedSpots.TextMatrix(llMatchedRow, MSDFCODEINDEX) = tmSdf.lCode
    grdMatchedSpots.TextMatrix(llMatchedRow, MIIDFCODEINDEX) = tmIidf.lCode
    grdMatchedSpots.TextMatrix(llMatchedRow, MSSORTINDEX) = ""
    
    'Remove from grdMatchedSpots
    grdNetworkSpots.RemoveItem lmNRowSelected
    grdStationSpots.RemoveItem lmSRowSelected
    
    'Change counts in grdMatchedResult
    If lmMRRowSelected >= grdMatchedResult.FixedRows Then
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRMATCHCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRMATCHCOUNTINDEX)) + 1
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRNETCOUNTINDEX)) - 1
        grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCOUNTINDEX) = Val(grdMatchedResult.TextMatrix(lmMRRowSelected, MRSTNCOUNTINDEX)) - 1
    End If
    
    lmNRowSelected = -1
    lmSRowSelected = -1
    mDetailSetCommands
End Sub

Private Sub mMoveFile(slType As String, slInFileName As String, llSpotDate As Long)
    'slType(I): P=Processed; U=Unreadable
    'slSubType(I): Pdf or Txt
    Dim slMovePath As String
    Dim fs As New FileSystemObject
    Dim slInvoiceStartDate As String
    Dim ilPass As Integer
    Dim slSubType As String
    Dim slFileName As String
    
    On Error GoTo mMoveFileErr
    If llSpotDate = 99999999 Then
        slInvoiceStartDate = "_" & Format(Now, "MMDDYYYY")
    Else
        slInvoiceStartDate = "_" & Format(gObtainStartStd(Format(llSpotDate, "m/d/yy")), "MMDDYYYY")
    End If
    For ilPass = 0 To 1 Step 1
        If ilPass = 0 Then
            slSubType = "Pdf"
            slFileName = slInFileName
        Else
            slSubType = "Txt"
            slFileName = Replace(slInFileName, ".Pdf", ".Txt")
        End If
        If slType = "U" Then
            slMovePath = sgBrowserDrivePath & "StationInvoices-Unreadable-" & slSubType & slInvoiceStartDate
        Else
            slMovePath = sgBrowserDrivePath & "StationInvoices-Processed-" & slSubType & slInvoiceStartDate
        End If
        If Not fs.FolderExists(slMovePath) Then
            fs.CreateFolder (slMovePath)
        End If
        If fs.FILEEXISTS(slMovePath & "\" & slFileName) Then
            fs.DeleteFile slMovePath & "\" & slFileName, True
        End If
        fs.MoveFile sgBrowserDrivePath & slFileName, slMovePath & "\" & slFileName
    Next ilPass
    Set fs = Nothing
    Exit Sub
mMoveFileErr:
    If ilPass = 0 Then
        If slType = "U" Then
            MsgBox "Unable to move file " & slFileName & " to " & sgBrowserDrivePath & "StationInvoices-Unreadable" & slInvoiceStartDate, vbOKOnly + vbInformation, "Warning"
        Else
            MsgBox "Unable to move file " & slFileName & " to " & sgBrowserDrivePath & "StationInvoices-Processed" & slInvoiceStartDate, vbOKOnly + vbInformation, "Warning"
        End If
    End If
    Exit Sub
End Sub

Private Sub mShowPreviousResults(slPDFFileName As String, llMatchResultRow As Long, llNoMatchResultRow As Long)
    Dim ilRet As Integer
    Dim slInvoiceDate As String
    Dim ilPos As Integer
    Dim ilAdf As Integer
    Dim ilVef As Integer
    Dim ilCountC As Integer
    Dim ilCountI As Integer
    Dim ilCountM As Integer
    Dim llAirTime As Long
    Dim llAirDate As Long
    Dim blCompliant As Boolean
    Dim ilLen As Integer
    
    smIihfFileName = slPDFFileName
    ilPos = InStrRev(sgBrowserDrivePath, "_", -1, vbTextCompare)
    If ilPos <= 0 Then
        Exit Sub
    End If
    slInvoiceDate = Mid$(sgBrowserDrivePath, ilPos + 1)
    If right(slInvoiceDate, 1) = "\" Then
        slInvoiceDate = Left$(slInvoiceDate, Len(slInvoiceDate) - 1)
    End If
    slInvoiceDate = Left$(slInvoiceDate, 2) & "/" & Mid$(slInvoiceDate, 3, 2) & "/" & Mid$(slInvoiceDate, 5)
    tmIihfSrchKey3.sFileName = slPDFFileName
    'gPackDate slInvoiceDate, tmIihfSrchKey3.iInvStartDate(0), tmIihfSrchKey3.iInvStartDate(1)
    gPackDate "1/1/1970", tmIihfSrchKey3.iInvStartDate(0), tmIihfSrchKey3.iInvStartDate(1)
    ilRet = btrGetGreaterOrEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (Trim$(tmIihf.sFileName) = smIihfFileName)
        ilVef = gBinarySearchVef(tmIihf.iVefCode)
        If ilVef <> -1 Then
            smCallLetters = Trim$(tgMVef(ilVef).sName)
            smAdvertiserName = ""
            If tmIihf.lChfCode > 0 Then
                tmChfSrchKey0.lCode = tmIihf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            smEstimateNumber = Trim$(tmIihf.sStnEstimateNo)
            smContractNumber = Trim$(tmIihf.sStnContractNo)
            smInvoiceNumber = Trim$(tmIihf.sStnInvoiceNo)
            smAdvertiserName = Trim$(tmIihf.sStnAdvtName)
            ilCountC = 0
            ilCountI = 0
            ilCountM = 0
            blCompliant = True
            ReDim ilStnLen(0 To 0) As Integer
            tmIidfSrchKey1.lCode = tmIihf.lCode
            ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmIidf.lIihfCode = tmIihf.lCode)
                If tmIidf.sSpotMatchType = "M" Then
                    ilCountM = ilCountM + 1
                ElseIf tmIidf.sSpotMatchType = "I" Then
                    ilStnLen(ilCountI) = tmIidf.iStnSpotLen
                    ilCountI = ilCountI + 1
                    ReDim Preserve ilStnLen(0 To ilCountI) As Integer
                Else
                    ilCountC = ilCountC + 1
                    gUnpackDateLong tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), llAirDate
                    gUnpackTimeLong tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), False, llAirTime
                    tmSdfSrchKey3.lCode = tmIidf.lSdfCode
                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        If mGetCompliantStatus(llAirDate, llAirTime) = "O" Then
                            blCompliant = False
                        End If
                    End If
                End If
                ilRet = btrGetNext(hmIidf, tmIidf, imIidfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (ilCountC > 0) Or (ilCountM > 0) Then
                mAddToMatchGrid llMatchResultRow, tmIihf.lCode
                grdMatchedResult.TextMatrix(llMatchResultRow, MRMATCHCOUNTINDEX) = ilCountC
                grdMatchedResult.TextMatrix(llMatchResultRow, MRNETCOUNTINDEX) = ilCountM
                grdMatchedResult.TextMatrix(llMatchResultRow, MRSTNCOUNTINDEX) = ilCountI
                If Not blCompliant Then
                    grdMatchedResult.TextMatrix(llMatchResultRow, MRCOMPLIANTINDEX) = "No"
                End If
                llMatchResultRow = llMatchResultRow + 1
            Else
                ReDim tmImportSpotInfo(0 To ilCountI) As IMPORTSPOTINFO
                For ilLen = 0 To UBound(ilStnLen) - 1 Step 1
                    tmImportSpotInfo(ilLen).iLen = ilStnLen(ilLen)
                Next ilLen
                mAddToNoMatchGrid llNoMatchResultRow, tmIihf.lCode, ""
            End If
        End If
        ilRet = btrGetNext(hmIihf, tmIihf, imIihfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub

Private Function mGetCompliantStatus(llAirDate As Long, llAirTime As Long) As String
    Dim slDays As String
    Dim slTimeRange As String
    Dim slDateRange As String
    Dim slEDIDays As String
    Dim ilPos As Integer
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDPTime As Integer
    Dim slAcqRate As String
    Dim slDPTimes() As String
    
    mGetCompliantStatus = "O"
    mDetermineDP slDays, slTimeRange, slDateRange, slEDIDays, slAcqRate
    ilPos = InStr(1, slDateRange, "-", vbTextCompare)
    If ilPos > 0 Then
        llStartDate = gDateValue(Left$(slDateRange, ilPos - 1))
        llEndDate = gDateValue(Mid$(slDateRange, ilPos + 1))
    Else
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llStartDate
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llEndDate
    End If
    If (llAirDate >= llStartDate) And (llAirDate <= llEndDate) Then
        slDPTimes = Split(slTimeRange, ";")
        If IsArray(slDPTimes) Then
            For ilDPTime = 0 To UBound(slDPTimes) Step 1
                ilPos = InStr(1, slDPTimes(ilDPTime), "-", vbTextCompare)
                If ilPos > 0 Then
                    llStartTime = gTimeToLong(Left$(slDPTimes(ilDPTime), ilPos - 1), False)
                    llEndTime = gTimeToLong(Mid$(slDPTimes(ilDPTime), ilPos + 1), False)
                    If (llAirTime >= llStartTime) And (llAirTime <= llEndTime) Then
                        ilDay = gWeekDayLong(llAirDate)
                        If Mid$(slEDIDays, ilDay + 1, 1) = "Y" Then
                            mGetCompliantStatus = "A"
                            Exit For
                        End If
                    End If
                End If
            Next ilDPTime
        End If
    End If
    
End Function

Private Sub mPopCntr()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilChf As Integer
    Dim ilVefCode As Integer
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim llIihfCode As Long
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim ilRdf As Integer
    Dim ilClf As Long
    Dim ilSpotCount As Integer
    Dim blPreviouslyPosted As Boolean
    Dim ilLen As Integer
    Dim blLenFound As Boolean
    
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbHourglass
    llRow = grdCntrStation.FixedRows
    grdCntrStation.Row = llRow
    For llCol = NMRSTATIONINDEX To NMRIIHFCODEINDEX Step 1
        grdCntrStation.Col = llCol
        grdCntrStation.CellBackColor = LIGHTYELLOW
        grdCntrStation.TextMatrix(llRow, llCol) = grdNoMatchedResult.TextMatrix(lmNMRRowSelected, llCol)
    Next llCol
    lacPossibleCntr.Caption = "Unposted and Manually Posted Contracts Sold " & grdCntrStation.TextMatrix(llRow, NMRSTATIONINDEX)
    llIihfCode = Val(grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRIIHFCODEINDEX))
    tmIihfSrchKey0.lCode = llIihfCode
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ilVefCode = tmIihf.iVefCode
        gUnpackDateLong tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), llInvStartdate
        llInvEndDate = gDateValue(gObtainEndStd(Format(llInvStartdate, "m/d/yy")))
        mGetContracts ilVefCode, llInvStartdate, llInvEndDate, blPreviouslyPosted
    Else
        ReDim tmUnpostedCntrInfo(0 To 0) As UNPOSTEDCNTRINFO
    End If
    'Obtain contracts
    lmCRowSelected = -1
    cmcCntrMatch.Enabled = False
    imCLastResultColSorted = -1
    imCLastResultSort = -1
    mClearGrid grdCntrNetwork
    llRow = grdCntrNetwork.FixedRows
    If UBound(tmUnpostedCntrInfo) > 0 Then
        For ilChf = 0 To UBound(tmUnpostedCntrInfo) - 1 Step 1
            tmChfSrchKey0.lCode = tmUnpostedCntrInfo(ilChf).lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilSpotCount = 0
                If llRow >= grdCntrNetwork.Rows Then
                    grdCntrNetwork.AddItem ""
                End If
                grdCntrNetwork.RowHeight(llRow) = fgBoxGridH + 15
                ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
                If ilAdf <> -1 Then
                    grdCntrNetwork.TextMatrix(llRow, CADVERTISERINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                End If
                grdCntrNetwork.TextMatrix(llRow, CESTIMATEINDEX) = Trim$(tmChf.sAgyEstNo) & Trim$(tmChf.sTitle)
                grdCntrNetwork.TextMatrix(llRow, CCONTRACTINDEX) = tmChf.lCntrNo
                grdCntrNetwork.TextMatrix(llRow, CPRODUCTINDEX) = Trim$(tmChf.sProduct)
                If Trim$(tmUnpostedCntrInfo(ilChf).sSourceForm) = "T" Then
                    grdCntrNetwork.TextMatrix(llRow, CSTATUSINDEX) = "Manually Posted- Times"
                ElseIf Trim$(tmUnpostedCntrInfo(ilChf).sSourceForm) = "C" Then
                    grdCntrNetwork.TextMatrix(llRow, CSTATUSINDEX) = "Manually Posted- Counts"
                Else
                    grdCntrNetwork.TextMatrix(llRow, CSTATUSINDEX) = ""
                End If
                grdCntrNetwork.TextMatrix(llRow, CPREVIEWINDEX) = ""
                grdCntrNetwork.Row = llRow
                grdCntrNetwork.Col = CPREVIEWINDEX
                grdCntrNetwork.CellBackColor = GRAY
                ReDim llClfCode(0 To 0) As Long
                ReDim ilClfLen(0 To 0) As Integer
                tmClfSrchKey1.lChfCode = tmChf.lCode
                tmClfSrchKey1.iVefCode = ilVefCode
                ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iVefCode = ilVefCode)
                    'ilRdf = gBinarySearchRdf(tmClf.iRdfcode)
                    'If ilRdf <> -1 Then
                    '    tmRdf = tgMRdf(ilRdf)
                    '    If grdCntrNetwork.TextMatrix(llRow, CDPINFOINDEX) = "" Then
                    '        grdCntrNetwork.TextMatrix(llRow, CDPINFOINDEX) = Trim$(tmRdf.sName)
                    '    Else
                    '        grdCntrNetwork.TextMatrix(llRow, CDPINFOINDEX) = grdCntrNetwork.TextMatrix(llRow, CDPINFOINDEX) & ";" & Trim$(tmRdf.sName)
                    '    End If
                    'End If
                    llClfCode(UBound(llClfCode)) = tmClf.lCode
                    ReDim Preserve llClfCode(0 To UBound(llClfCode) + 1) As Long
                    blLenFound = False
                    For ilLen = 0 To UBound(ilClfLen) - 1 Step 1
                        If ilClfLen(ilLen) = tmClf.iLen Then
                            blLenFound = True
                            Exit For
                        End If
                    Next ilLen
                    If Not blLenFound Then
                        ilClfLen(UBound(ilClfLen)) = tmClf.iLen
                        ReDim Preserve ilClfLen(0 To UBound(ilClfLen) + 1) As Integer
                    End If
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX) = ""
                For ilLen = 0 To UBound(ilClfLen) - 1 Step 1
                    ilSpotCount = 0
                    For ilClf = 0 To UBound(llClfCode) - 1 Step 1
                        ilSpotCount = ilSpotCount + mDetermineOrderedSpotCount(tmChf.lCode, llClfCode(ilClf), llInvStartdate, llInvEndDate, ilClfLen(ilLen))
                    Next ilClf
                    If grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX) = "" Then
                        grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX) = ilClfLen(ilLen) & "s: " & ilSpotCount
                    Else
                        grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX) = grdCntrNetwork.TextMatrix(llRow, CORDERCOUNTINDEX) & "; " & ilClfLen(ilLen) & "s: " & ilSpotCount
                    End If
                Next ilLen
                
                grdCntrNetwork.TextMatrix(llRow, CSELECTEDINDEX) = "0"
                grdCntrNetwork.TextMatrix(llRow, CSORTINDEX) = ""
                grdCntrNetwork.TextMatrix(llRow, CCHFCODEINDEX) = tmChf.lCode
                llRow = llRow + 1
            End If
        Next ilChf
    ElseIf blPreviouslyPosted Then
        If llRow >= grdCntrNetwork.Rows Then
            grdCntrNetwork.AddItem ""
        End If
        grdCntrNetwork.RowHeight(llRow) = fgBoxGridH + 15
        grdCntrNetwork.TextMatrix(llRow, CCONTRACTINDEX) = "Already Posted"
    End If
    mMousePointer grdMatchedResult, grdNoMatchedResult, vbDefault
    frcCntr.Visible = True
    frcResult.Visible = False
End Sub

Private Function mGetContracts(ilVefCode As Integer, llInvStartdate As Long, llInvEndDate As Long, blPreviouslyPosted As Boolean) As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llChfCode As Long
    Dim ilChf As Integer
    Dim llRow As Long
    Dim blFound As Boolean
    Dim slStr As String
    Dim slSourceForm As String
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tmUnpostedCntrInfo(0 To 0) As UNPOSTEDCNTRINFO
    blPreviouslyPosted = False
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)  'Extract operation record size
    tmSdfSrchKey1.iVefCode = ilVefCode
    slDate = Format$(llInvStartdate, "m/d/yy")
    gPackDate slDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""   'slType
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tmSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)
        ' And on the records where the date is equal to the passed in log date
        slDate = Format$(llInvStartdate, "m/d/yy")
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        slDate = Format$(llInvEndDate, "m/d/yy")
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        'tlCharTypeBuff.sType = "M"    'Extract all non-matching records
        'ilOffset = gFieldOffset("Sdf", "SdfSchStatus")
        'ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        ilRet = btrExtAddField(hmSdf, 0, ilExtLen) 'Extract the whole record
        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tmSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                blFound = False
                For ilChf = 0 To UBound(tmUnpostedCntrInfo) - 1 Step 1
                    If tmUnpostedCntrInfo(ilChf).lChfCode = tmSdf.lChfCode Then
                        blFound = True
                        blPreviouslyPosted = True
                        Exit For
                    End If
                Next ilChf
                If Not blFound Then
                    
                    'For llRow = grdMatchedResult.FixedRows To grdMatchedResult.Rows - 1 Step 1
                    '    slStr = Trim$(grdMatchedResult.TextMatrix(llRow, MRSTNSTATIONINDEX))
                    '    If slStr <> "" Then
                    '        If Val(grdMatchedResult.TextMatrix(llRow, MRCHFCODEINDEX)) = tmSdf.lChfCode Then
                    '            blFound = True
                    '            Exit For
                    '        End If
                    '    End If
                    'Next llRow
                    slSourceForm = ""
                    tmIihfSrchKey2.lChfCode = tmSdf.lChfCode
                    tmIihfSrchKey2.iVefCode = tmSdf.iVefCode
                    gPackDateLong llInvStartdate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
                    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
                            slSourceForm = Trim$(tmIihf.sSourceForm)
                        Else
                            blFound = True
                        End If
                    End If
                End If
                If Not blFound Then
                    tmUnpostedCntrInfo(UBound(tmUnpostedCntrInfo)).lChfCode = tmSdf.lChfCode
                    tmUnpostedCntrInfo(UBound(tmUnpostedCntrInfo)).sSourceForm = slSourceForm
                    ReDim Preserve tmUnpostedCntrInfo(0 To UBound(tmUnpostedCntrInfo) + 1) As UNPOSTEDCNTRINFO
                End If
                ilExtLen = Len(tmSdf)
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If

End Function

Private Sub mProcessUserMatch(Optional blPreviewMode As Boolean = False)
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim llIihfCode As Long
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim ilUpper As Integer
    Dim llRow As Integer
    Dim llMatchRow As Long
    Dim llNoMatchRow As Long
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim ilIidf As Integer
    
    llChfCode = Val(grdCntrNetwork.TextMatrix(lmCRowSelected, CCHFCODEINDEX))
    tmChfSrchKey0.lCode = llChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
    If ilAdf = -1 Then
        Exit Sub
    End If
    ilAdfCode = tgCommAdf(ilAdf).iCode
    'smAdvertiserName = Trim$(tgCommAdf(ilAdf).sName)
    smAdvertiserName = Trim$(grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRADVERTISERINDEX))
    llIihfCode = Val(grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRIIHFCODEINDEX))
    tmIihfSrchKey0.lCode = llIihfCode
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ilVefCode = tmIihf.iVefCode
        If tmIihf.lAmfCode > 0 Then
            tmAmfSrchKey0.lCode = tmIihf.lAmfCode
            ilRet = btrGetEqual(hmAmf, tmAmf, imAmfRecLen, tmAmfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                smAdvertiserName = Trim$(tmAmf.sStationAdvtName)
                ilAdfCode = tmAmf.iAdfCode
            End If
        End If
        gUnpackDateLong tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), llInvStartdate
        llInvEndDate = gDateValue(gObtainEndStd(Format(llInvStartdate, "m/d/yy")))
        smCallLetters = Trim$(grdNoMatchedResult.TextMatrix(lmNMRRowSelected, NMRSTATIONINDEX))
        smIihfFileName = Trim$(tmIihf.sFileName)
        smEstimateNumber = Trim$(tmIihf.sStnEstimateNo)
        smInvoiceNumber = Trim$(tmIihf.sStnInvoiceNo)
        smContractNumber = Trim$(tmIihf.sStnContractNo)
        smSourceForm = Trim$(tmIihf.sSourceForm)
        smNetContractNumber = ""
        ReDim tmImportSpotInfo(0 To 0) As IMPORTSPOTINFO
        ReDim llIldfCode(0 To 0) As Long
        tmIidfSrchKey1.lCode = llIihfCode
        ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmIidf.lIihfCode = llIihfCode)
            llIldfCode(UBound(llIldfCode)) = tmIidf.lCode
            ReDim Preserve llIldfCode(0 To UBound(llIldfCode) + 1) As Long
            ilRet = btrGetNext(hmIidf, tmIidf, imIidfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        For ilIidf = 0 To UBound(llIldfCode) - 1 Step 1
            tmIidfSrchKey0.lCode = llIldfCode(ilIidf)
            ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                If (tmIidf.sSpotMatchType = "C") Or (tmIidf.sSpotMatchType = "I") Then
                    ilUpper = UBound(tmImportSpotInfo)
                    gUnpackDateLong tmIidf.iStnSpotAirDate(0), tmIidf.iStnSpotAirDate(1), tmImportSpotInfo(ilUpper).lAirDate
                    gUnpackTimeLong tmIidf.iStnSpotAirTime(0), tmIidf.iStnSpotAirTime(1), False, tmImportSpotInfo(ilUpper).lAirTime
                    tmImportSpotInfo(ilUpper).iLen = tmIidf.iStnSpotLen
                    tmImportSpotInfo(ilUpper).sISCI = Trim$(mGetStnISCI(tmIidf.lStnCpfCode)) 'tmIidf.sStnISCI
                    If Trim$(tmIidf.sStnDPDays) = "" Then
                        tmImportSpotInfo(ilUpper).lDPStartTime = -1
                        tmImportSpotInfo(ilUpper).lDPEndTime = -1
                        tmImportSpotInfo(ilUpper).sDPDays = ""
                    Else
                        tmImportSpotInfo(ilUpper).lDPStartTime = tmIidf.lStnDPStartTime
                        tmImportSpotInfo(ilUpper).lDPEndTime = tmIidf.lStnDPEndTime
                        tmImportSpotInfo(ilUpper).sDPDays = tmIidf.sStnDPDays
                    End If
                    tmImportSpotInfo(ilUpper).bMatched = False
                    tmImportSpotInfo(ilUpper).lRate = -1
                    ReDim Preserve tmImportSpotInfo(0 To ilUpper + 1) As IMPORTSPOTINFO
                End If
                If Not blPreviewMode Then
                    ilRet = btrDelete(hmIidf)
                End If
            End If
            'tmIidfSrchKey1.lCode = llIihfCode
            'ilRet = btrGetEqual(hmIidf, tmIidf, imIidfRecLen, tmIidfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        Next ilIidf
        If Not blPreviewMode Then
            ilRet = btrDelete(hmIihf)
            grdNoMatchedResult.RemoveItem lmNMRRowSelected
            lmNMRRowSelected = -1
        End If
        llNoMatchRow = grdNoMatchedResult.FixedRows
        For llRow = grdNoMatchedResult.FixedRows To grdNoMatchedResult.Rows - 1 Step 1
            If grdNoMatchedResult.TextMatrix(llRow, NMRFILENAMEINDEX) <> "" Then
                llNoMatchRow = llRow + 1
            End If
        Next llRow
        llMatchRow = grdMatchedResult.FixedRows
        For llRow = grdMatchedResult.FixedRows To grdMatchedResult.Rows - 1 Step 1
            If grdMatchedResult.TextMatrix(llRow, MRNETADVERTISERINDEX) <> "" Then
                llMatchRow = llRow + 1
            End If
        Next llRow
        ilRet = mMatchSpots(llMatchRow, llNoMatchRow, llInvStartdate, llInvEndDate, llChfCode, ilAdfCode, blPreviewMode)
        If Not blPreviewMode Then
            mUpdateApf llMatchRow - 1
            mSetCommands
        End If
    End If
End Sub

Private Sub mAddMapAdf(ilVefCode As Integer, llAmfCode As Long)
    Dim ilAdf As Integer
    Dim slAdvertiser As String
    Dim ilRet As Integer
    ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
    If ilAdf <> -1 Then
        slAdvertiser = UCase$(Trim$(smAdvertiserName))
        If UCase$(Trim$(tgCommAdf(ilAdf).sName)) <> slAdvertiser Then
            'Test if in amf.  If not, add it
            tmAmfSrchKey2.iVefCode = ilVefCode
            tmAmfSrchKey2.sStationAdvtName = slAdvertiser
            ilRet = btrGetEqual(hmAmf, tmAmf, imAmfRecLen, tmAmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                tmAmf.lCode = 0
                tmAmf.sStationAdvtName = smAdvertiserName
                tmAmf.iVefCode = ilVefCode
                tmAmf.iAdfCode = tgCommAdf(ilAdf).iCode
                tmAmf.sUnused = ""
                ilRet = btrInsert(hmAmf, tmAmf, imAmfRecLen, INDEXKEY0)
                If ilRet = BTRV_ERR_NONE Then
                    llAmfCode = tmAmf.lCode
                End If
            End If
        End If
    End If
End Sub

Private Function mAddCpf(slISCI As String) As Long
    Dim ilRet As Integer
    
    tmCpfSrchKey1.sISCI = Trim$(slISCI)
    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        mAddCpf = tmCpf.lCode
    Else
        tmCpf.lCode = 0
        tmCpf.sName = ""
        tmCpf.sISCI = slISCI
        tmCpf.sCreative = ""
        tmCpf.iRotEndDate(0) = 0
        tmCpf.iRotEndDate(1) = 0
        tmCpf.lSifCode = 0
        ilRet = btrInsert(hmCpf, tmCpf, imCpfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            mAddCpf = 0
        Else
            mAddCpf = tmCpf.lCode
        End If
    End If
    Exit Function
End Function
Private Function mAddOrUpdateCif(llCpfCode As Long, ilAdfCode As Integer, ilLen As Integer) As Long
    Dim ilRet As Integer
    Dim slDate As String
    
    If llCpfCode <= 0 Then
        mAddOrUpdateCif = 0
        Exit Function
    End If
    tmCifSrchKey2.lCode = llCpfCode
    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        If ilAdfCode > 0 Then
            If tmCif.iAdfCode <= 0 Then
                tmCif.iAdfCode = ilAdfCode
                ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
            End If
        End If
        mAddOrUpdateCif = tmCif.lCode
        Exit Function
    End If
    'Add copy inventory
    tmCif.lCode = 0  'Autoincrement
    tmCif.iMcfCode = 0
    tmCif.sName = ""
    tmCif.sCut = ""
    tmCif.sReel = ""
    tmCif.iLen = ilLen
    tmCif.iEtfCode = 0
    tmCif.iEnfCode = 0
    tmCif.iAdfCode = ilAdfCode
    tmCif.lcpfCode = llCpfCode
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.iMnfAnn = 0
    tmCif.sHouse = "Y"
    tmCif.sCleared = "Y"
    tmCif.lCsfCode = 0
    tmCif.iNoTimesAir = 0
    tmCif.sCartDisp = "N"
    tmCif.sTapeDisp = "N"
    tmCif.sPurged = "A"
    tmCif.iPurgeDate(0) = 0
    tmCif.iPurgeDate(1) = 0
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.iUrfCode = tgUrf(0).iCode
    tmCif.sPrint = "N"
    tmCif.iLangMnfCode = 0
    tmCif.sUnused = ""
    ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddOrUpdateCif = tmCif.lCode
    Else
        mAddOrUpdateCif = 0
    End If
    Exit Function
End Function

Private Function mGetStnISCI(llStnCpfCode As Long) As String
    Dim ilRet As Integer
    tmCpfSrchKey.lCode = llStnCpfCode
    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        mGetStnISCI = Trim$(tmCpf.sISCI)
    Else
        mGetStnISCI = ""
    End If

End Function

Private Function mUpdateApf(llChgRowNo As Long)
    Dim llCntrNo As Long
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim slCallLetters As String
    Dim llRow As Long
    Dim slInvDate As String
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim llSdfDate As Long
    Dim llApfInvDate As Long
    Dim ilRet As Integer
    Dim ilApf As Integer
    Dim ilVef As Integer
    Dim llIihfCode As Long
    Dim llStartRow As Long
    Dim llEndRow As Long
    ReDim llApfCode(0 To 0) As Long
    
    If (Asc(tgSaf(0).sFeatures3) And REQSTATIONPOSTING) = REQSTATIONPOSTING Then 'Require to Post spot prior to invoicing
        mUpdateApf = True
        Exit Function
    End If
    If (Asc(tgSaf(0).sFeatures2) And PAYMENTONCOLLECTION) <> PAYMENTONCOLLECTION Then 'Payment on Collection
        mUpdateApf = True
        Exit Function
    End If
    If llChgRowNo >= grdMatchedResult.FixedRows Then
        llStartRow = llChgRowNo
        llEndRow = llChgRowNo
    Else
        llStartRow = grdMatchedResult.FixedRows
        llEndRow = grdMatchedResult.Rows - 1
    End If
    For llRow = llStartRow To llEndRow Step 1
        slCallLetters = UCase$(Trim$(grdMatchedResult.TextMatrix(llRow, MRSTNSTATIONINDEX)))
        If (slCallLetters <> "") Then
            ilVefCode = -1
            llIihfCode = Val(grdMatchedResult.TextMatrix(llRow, MRIIHFCODEINDEX))
            tmIihfSrchKey0.lCode = llIihfCode
            ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilVefCode = tmIihf.iVefCode
                gUnpackDate tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), slInvDate
                llInvStartdate = gDateValue(slInvDate)
                llInvEndDate = gDateValue(gObtainEndStd(slInvDate))
            End If
            If ilVefCode <> -1 Then
                ReDim llApfCode(0 To 0) As Long
                llCntrNo = Val(grdMatchedResult.TextMatrix(llRow, MRNETCONTRACTINDEX))
                llChfCode = Val(grdMatchedResult.TextMatrix(llRow, MRCHFCODEINDEX))
                tmApfSrchKey4.lCntrNo = llCntrNo
                gPackDate "1/1/1970", tmApfSrchKey4.iFullyPaidDate(0), tmApfSrchKey4.iFullyPaidDate(1)
                'ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORWRITE)
                ilRet = btrGetGreaterOrEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmApf.lCntrNo = llCntrNo)
                    gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llApfInvDate
                    If (llInvEndDate = llApfInvDate) And (tmApf.iVefCode = ilVefCode) And (tmApf.lSbfCode = 0) Then
                        llApfCode(UBound(llApfCode)) = tmApf.lCode
                        ReDim Preserve llApfCode(0 To UBound(llApfCode) + 1) As Long
                    End If
                    ilRet = btrGetNext(hmApf, tmApf, imApfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
                For ilApf = 0 To UBound(llApfCode) - 1 Step 1
                    tmApfSrchKey0.lCode = llApfCode(ilApf)
                    ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        'Find lines
                        'tmApf.iAiredSpotCount = 0
                        'tmClfSrchKey1.lChfCode = llChfCode
                        'tmClfSrchKey1.iVefCode = ilVefCode
                        'ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iVefCode = ilVefCode)
                        '    If tmApf.lAcquisitionCost = tmClf.lAcquisitionCost Then
                        '        'Update aired count
                        '        tmSdfSrchKey0.iVefCode = ilVefCode
                        '        tmSdfSrchKey0.lChfCode = llChfCode
                        '        tmSdfSrchKey0.iLineNo = tmClf.iLine
                        '        tmSdfSrchKey0.lFsfCode = 0
                        '        gPackDateLong llInvStartdate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
                        '        tmSdfSrchKey0.sSchStatus = ""
                        '        gPackTime "12AM", tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
                        '        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        '        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = tmClf.iLine)
                        '            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                        '            If llSdfDate > llInvEndDate Then
                        '                Exit Do
                        '            End If
                        '            'If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        '            If mIncludeSpot(llInvStartdate, llInvEndDate) Then
                        '                tmApf.iAiredSpotCount = tmApf.iAiredSpotCount + 1
                        '            End If
                        '            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        '        Loop
                        '    End If
                        '    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        'Loop
                        tmApf.iAiredSpotCount = mGetAiredCount(llChfCode, ilVefCode, llInvStartdate, llInvEndDate)
                        tmApf.sStationCntrNo = grdMatchedResult.TextMatrix(llRow, MRSTNCONTRACTINDEX)
                        tmApf.sStationInvNo = grdMatchedResult.TextMatrix(llRow, MRSTNINVOICEINDEX)
                        ilRet = btrUpdate(hmApf, tmApf, imApfRecLen)
                    End If
                Next ilApf
            End If
        End If
    Next llRow
    mUpdateApf = True
End Function

Private Function mClearApfAirCount(llCntrNo As Long, llIihfCode As Long)
    Dim ilVefCode As Integer
    Dim slInvDate As String
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim llApfInvDate As Long
    Dim ilRet As Integer
    Dim ilApf As Integer
    ReDim llApfCode(0 To 0) As Long
    
    If (Asc(tgSaf(0).sFeatures3) And REQSTATIONPOSTING) = REQSTATIONPOSTING Then 'Require to Post spot prior to invoicing
        mClearApfAirCount = True
        Exit Function
    End If
    If (Asc(tgSaf(0).sFeatures2) And PAYMENTONCOLLECTION) <> PAYMENTONCOLLECTION Then 'Payment on Collection
        mClearApfAirCount = True
        Exit Function
    End If
    ilVefCode = -1
    tmIihfSrchKey0.lCode = llIihfCode
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ilVefCode = tmIihf.iVefCode
        gUnpackDate tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1), slInvDate
        llInvStartdate = gDateValue(slInvDate)
        llInvEndDate = gDateValue(gObtainEndStd(slInvDate))
    End If
    If ilVefCode <> -1 Then
        ReDim llApfCode(0 To 0) As Long
        tmApfSrchKey4.lCntrNo = llCntrNo
        gPackDate "1/1/1970", tmApfSrchKey4.iFullyPaidDate(0), tmApfSrchKey4.iFullyPaidDate(1)
        'ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORWRITE)
        ilRet = btrGetGreaterOrEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmApf.lCntrNo = llCntrNo)
            gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llApfInvDate
            If (llInvEndDate = llApfInvDate) And (tmApf.iVefCode = ilVefCode) And (tmApf.lSbfCode = 0) Then
                llApfCode(UBound(llApfCode)) = tmApf.lCode
                ReDim Preserve llApfCode(0 To UBound(llApfCode) + 1) As Long
            End If
            ilRet = btrGetNext(hmApf, tmApf, imApfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        For ilApf = 0 To UBound(llApfCode) - 1 Step 1
            tmApfSrchKey0.lCode = llApfCode(ilApf)
            ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                'Find lines
                tmApf.iAiredSpotCount = 0
                ilRet = btrUpdate(hmApf, tmApf, imApfRecLen)
            End If
        Next ilApf
    End If
    mClearApfAirCount = True
End Function
Private Function mDetermineOrderedSpotCount(llChfCode As Long, llClfCode As Long, llInvStartdate As Long, llInvEndDate As Long, ilLen As Integer) As Integer
    Dim ilRet As Integer
    Dim llClfStartDate As Long
    Dim llClfEndDate As Long
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilDay As Integer
    Dim ilNoSpots As Integer
    Dim llDate As Long
    Dim tlClf As CLF
    
    ilNoSpots = 0
    tmChfSrchKey0.lCode = llChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mDetermineOrderedSpotCount = 0
        Exit Function
    End If
    tmClfSrchKey2.lCode = llClfCode
    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mDetermineOrderedSpotCount = 0
        Exit Function
    End If
    If tmClf.iLen <> ilLen Then
        mDetermineOrderedSpotCount = 0
        Exit Function
    End If
    gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llClfStartDate    'Week Start date
    gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llClfEndDate    'Week Start date
    If llClfStartDate > llClfEndDate Then
        mDetermineOrderedSpotCount = 0
        Exit Function
    End If
    If tmClf.sType = "H" Then
        'Check if Package line is CBS
        tmClfSrchKey0.lChfCode = tmChf.lCode
        tmClfSrchKey0.iLine = tmClf.iPkLineNo
        tmClfSrchKey0.iCntRevNo = tmClf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmClf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tmChf.lCode) And (tlClf.iLine = tmClf.iPkLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
            ilRet = btrGetNext(hmClf, tlClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tmChf.lCode) And (tlClf.iLine = tmClf.iPkLineNo) Then
            gUnpackDateLong tlClf.iStartDate(0), tlClf.iStartDate(1), llClfStartDate    'Week Start date
            gUnpackDateLong tlClf.iEndDate(0), tlClf.iEndDate(1), llClfEndDate    'Week Start date
            If llClfStartDate > llClfEndDate Then
                mDetermineOrderedSpotCount = 0
                Exit Function
            End If
        End If
    End If
    tmCffSrchKey0.lChfCode = tmChf.lCode
    tmCffSrchKey0.iClfLine = tmClf.iLine
    tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
    tmCffSrchKey0.iPropVer = tmClf.iPropVer
    tmCffSrchKey0.iStartDate(0) = 0
    tmCffSrchKey0.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = tmChf.lCode) And (tmCff.iClfLine = tmClf.iLine)
        If (tmCff.iCntRevNo = tmClf.iCntRevNo) And (tmCff.iPropVer = tmClf.iPropVer) And (tmCff.sDelete <> "Y") Then
            gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llCffStartDate    'Week Start date
            gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llCffEndDate    'Week Start date
            If llCffEndDate >= llInvStartdate Then
                For llDate = llInvStartdate To llInvEndDate Step 7
                    If (llDate + 6 >= llCffStartDate) And (llDate <= llCffEndDate) Then
                        If (tmCff.iSpotsWk <> 0) Or (tmCff.iXSpotsWk <> 0) Then 'Weekly
                            ilNoSpots = ilNoSpots + tmCff.iSpotsWk + tmCff.iXSpotsWk
                        Else    'Daily
                            For ilDay = 0 To 6 Step 1
                                If (llDate + ilDay >= llCffStartDate) And (llDate + ilDay <= llCffEndDate) Then
                                    ilNoSpots = ilNoSpots + tmCff.iDay(ilDay)
                                End If
                            Next ilDay
                        End If
                    End If
                Next llDate
            End If
            If llCffStartDate > llInvEndDate Then
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDetermineOrderedSpotCount = ilNoSpots
End Function

Private Function mRemoveExtraBlanks(slInStr As String) As String
    Dim slReturn As String
    Dim ilChar As Integer
    Dim slChar As String
    Dim slStr As String
    
    slReturn = ""
    slStr = Trim$(slInStr)
    For ilChar = 1 To Len(slStr) Step 1
        slChar = Mid(slStr, ilChar, 1)
        If Trim$(slChar) = "" Then
            If Trim$(Mid(slStr, ilChar - 1, 1)) <> "" Then
                slReturn = slReturn & slChar
            End If
        Else
            slReturn = slReturn & slChar
        End If
    Next ilChar
    mRemoveExtraBlanks = slReturn
End Function
Private Function mRemoveBlanks(slInStr As String) As String
    Dim slReturn As String
    Dim ilChar As Integer
    Dim slChar As String
    Dim slStr As String
    
    slReturn = ""
    slStr = Trim$(slInStr)
    For ilChar = 1 To Len(slStr) Step 1
        slChar = Mid(slStr, ilChar, 1)
        If Trim$(slChar) <> "" Then
            slReturn = slReturn & slChar
        End If
    Next ilChar
    mRemoveBlanks = slReturn
End Function

Private Sub mSetDetailButtons(Optional blPreviewMode As Boolean = False)
    If Not blPreviewMode Then
        cmcUndoReconcile.Visible = True
        cmcReconcile.Visible = True
        cmcReturn.Move frcDetail.Width / 2 - (3 * cmcReturn.Width) / 2, frcDetail.Height - cmcReturn.Height - 60
        cmcUndoReconcile.Move cmcReturn.Left + cmcReturn.Width + cmcUndoReconcile.Width / 2, cmcReturn.Top
        cmcReconcile.Move frcDetail.Width / 2 - cmcReconcile.Width / 2, (grdNetworkSpots.Top + grdNetworkSpots.Height) + (grdMatchedSpots.Top - (grdNetworkSpots.Top + grdNetworkSpots.Height)) / 2 - cmcReconcile.Height / 2
    Else
        cmcUndoReconcile.Visible = False
        cmcReconcile.Visible = False
        cmcReturn.Move frcDetail.Width / 2 - cmcReturn.Width / 2, frcDetail.Height - cmcReturn.Height - 60
    End If
End Sub

Private Function mIncludeSpot(llMonthStart As Long, llMonthEnd As Long) As Boolean
    Dim blIncludeSpot As Boolean
    Dim ilRet As Integer
    Dim llMissedDate As Long
    
    blIncludeSpot = True
    If tmSdf.sSpotType = "X" Then
        blIncludeSpot = False
    ElseIf tmSdf.sSchStatus = "M" Then
        blIncludeSpot = False
    Else
        If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
            If tmSdf.lSmfCode > 0 Then
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                    If (llMissedDate < llMonthStart) Or (llMissedDate > llMonthEnd) Then
                        blIncludeSpot = False
                    End If
                Else
                    blIncludeSpot = False
                End If
            Else
                blIncludeSpot = False
            End If
        End If
    End If
    mIncludeSpot = blIncludeSpot
End Function

Private Function mGetAiredCount(llChfCode As Long, ilVefCode As Integer, llInvStartdate As Long, llInvEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim llSdfDate As Long
    Dim ilAiredSpotCount As Integer
    
    'Required Files be opened: chf, clf, sdf, smf
    'Find lines
    ilAiredSpotCount = 0
    tmClfSrchKey1.lChfCode = llChfCode
    tmClfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iVefCode = ilVefCode)
        If tmApf.lAcquisitionCost = tmClf.lAcquisitionCost Then
            'Update aired count
            tmSdfSrchKey0.iVefCode = ilVefCode
            tmSdfSrchKey0.lChfCode = llChfCode
            tmSdfSrchKey0.iLineNo = tmClf.iLine
            tmSdfSrchKey0.lFsfCode = 0
            gPackDateLong llInvStartdate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
            tmSdfSrchKey0.sSchStatus = ""
            gPackTime "12AM", tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = tmClf.iLine)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                If llSdfDate > llInvEndDate Then
                    Exit Do
                End If
                If mIncludeSpot(llInvStartdate, llInvEndDate) Then
                    ilAiredSpotCount = ilAiredSpotCount + 1
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mGetAiredCount = ilAiredSpotCount

End Function

Private Sub mClearValues(llStartDate As Long, llEndDate As Long)
    smSourceForm = ""
    smInvoiceNumber = ""
    smContractNumber = ""
    smAdvertiserName = ""
    smEstimateNumber = ""
    smCallLetters = ""
    smNetContractNumber = ""
    llStartDate = 99999999
    llEndDate = 0
    lmInvStartDate = llStartDate
    ReDim tmImportSpotInfo(0 To 0) As IMPORTSPOTINFO
End Sub
