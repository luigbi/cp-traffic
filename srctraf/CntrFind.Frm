VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form CntrFind 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   9420
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9420
   Begin VB.CommandButton cmcRefresh 
      Appearance      =   0  'Flat
      Caption         =   "&Refresh"
      Height          =   285
      Left            =   5355
      TabIndex        =   40
      Top             =   5670
      Width           =   945
   End
   Begin VB.PictureBox pbcPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
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
      Height          =   1230
      Left            =   2745
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   39
      Top             =   2190
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox pbcAction 
      Height          =   5280
      Left            =   60
      ScaleHeight     =   5220
      ScaleWidth      =   9165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   9225
      Begin VB.CheckBox ckcShowISCI 
         Caption         =   "Show ISCI"
         Height          =   210
         Left            =   90
         TabIndex        =   38
         Top             =   4965
         Width           =   1155
      End
      Begin VB.OptionButton rbcShow 
         Caption         =   "Packages"
         Height          =   210
         Index           =   1
         Left            =   1260
         TabIndex        =   37
         Top             =   30
         Width           =   1275
      End
      Begin VB.OptionButton rbcShow 
         Caption         =   "Vehicles"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   30
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.PictureBox plcCalendar 
         Appearance      =   0  'Flat
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
         Height          =   1770
         Left            =   3300
         ScaleHeight     =   1740
         ScaleWidth      =   1965
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1995
         Begin VB.CommandButton cmcCalUp 
            Appearance      =   0  'Flat
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1635
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   45
            Width           =   285
         End
         Begin VB.CommandButton cmcCalDn 
            Appearance      =   0  'Flat
            Caption         =   "s"
            BeginProperty Font 
               Name            =   "Monotype Sorts"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   45
            Width           =   285
         End
         Begin VB.PictureBox pbcCalendar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ClipControls    =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   1440
            Left            =   45
            Picture         =   "CntrFind.frx":0000
            ScaleHeight     =   1410
            ScaleWidth      =   1845
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   255
            Width           =   1875
            Begin VB.Label lacDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   510
               TabIndex        =   27
               Top             =   405
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lacCalName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   210
            Left            =   330
            TabIndex        =   30
            Top             =   30
            Width           =   1305
         End
      End
      Begin VB.ListBox lbcSpotDates 
         Height          =   1110
         ItemData        =   "CntrFind.frx":2E1A
         Left            =   90
         List            =   "CntrFind.frx":2E1C
         TabIndex        =   33
         Top             =   3795
         Width           =   1110
      End
      Begin VB.ListBox lbcLineDates 
         Height          =   690
         ItemData        =   "CntrFind.frx":2E1E
         Left            =   90
         List            =   "CntrFind.frx":2E20
         TabIndex        =   32
         Top             =   2550
         Width           =   1110
      End
      Begin VB.CommandButton cmcAdvtCntr 
         Appearance      =   0  'Flat
         Caption         =   "&Get Advertiser"
         Height          =   285
         Left            =   7620
         TabIndex        =   31
         Top             =   2070
         Width           =   1470
      End
      Begin VB.TextBox edcSearch 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   1155
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox edcTo 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   585
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox edcFrom 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   585
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.ComboBox cbcSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   60
         TabIndex        =   5
         Top             =   930
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox edcFDDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Index           =   0
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1710
         Width           =   885
      End
      Begin VB.CommandButton cmcFDDropDown 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1905
         Picture         =   "CntrFind.frx":2E22
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1710
         Width           =   195
      End
      Begin VB.TextBox edcFDDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Index           =   1
         Left            =   1035
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2070
         Width           =   885
      End
      Begin VB.CommandButton cmcFDDropDown 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1905
         Picture         =   "CntrFind.frx":2F1C
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2070
         Width           =   195
      End
      Begin VB.ComboBox cbcMatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   2550
      End
      Begin VB.CheckBox ckcVehicleSports 
         Caption         =   "All Sports"
         Height          =   210
         Left            =   4335
         TabIndex        =   19
         Top             =   2070
         Width           =   1380
      End
      Begin VB.CheckBox ckcVehicleNonSports 
         Caption         =   "All Non-Sports"
         Height          =   210
         Left            =   2715
         TabIndex        =   18
         Top             =   2070
         Width           =   1545
      End
      Begin VB.ListBox lbcVehicles 
         Height          =   1950
         ItemData        =   "CntrFind.frx":3016
         Left            =   2715
         List            =   "CntrFind.frx":3018
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   75
         Width           =   3000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAdvtCntr 
         Height          =   1950
         Left            =   5835
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   75
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3440
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLines 
         Height          =   960
         Left            =   1260
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2415
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   16
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
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
         _Band(0).Cols   =   16
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpots 
         Height          =   1590
         Left            =   1260
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3525
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   2805
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
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
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lacSpotInfo 
         Caption         =   "Spot Info:"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   3555
         Width           =   1005
      End
      Begin VB.Label lacLineInfo 
         Caption         =   "Line Info:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2325
         Width           =   1035
      End
      Begin VB.Label lacTo 
         Caption         =   "To"
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   1350
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lacFrom 
         Caption         =   "From"
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   975
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lacSearch 
         Caption         =   "Search For:"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lacMatchOn 
         Caption         =   "Match On:"
         Height          =   210
         Left            =   105
         TabIndex        =   2
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label lacFStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   60
         TabIndex        =   11
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label lacFEndDate 
         Appearance      =   0  'Flat
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   60
         TabIndex        =   14
         Top             =   2055
         Width           =   900
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4170
      TabIndex        =   23
      Top             =   5670
      Width           =   945
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
      Left            =   1860
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   24
      Top             =   -15
      Width           =   75
   End
   Begin VB.Image imcPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8790
      Picture         =   "CntrFind.frx":301A
      Top             =   5580
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Find"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "CntrFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CntrFind.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Height adjustment factor

'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassAll As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imChgMode As Integer
Dim imBSMode As Integer

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=End Date; 3=Transmission Date
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar

Dim imBypassFocus As Integer

Dim lmHighlightRow As Long
Dim lmLineHighlightRow As Long

Dim imSetAll As Integer
Dim imAllClicked As Integer

Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String
Dim imSelectedVefCode() As Integer

Dim tmProductProtectionCode() As SORTCODE
Dim smProductProtectionCodeTag As String

Dim tmSalesperson() As SORTCODE
Dim smSalespersonTag As String

Dim tmTeam() As MNF
Dim smTeamTag As String

Dim imMoPrt() As Integer
Dim imTuPrt() As Integer
Dim imWePrt() As Integer
Dim imThPrt() As Integer
Dim imFrPrt() As Integer
Dim imSaPrt() As Integer
Dim imSuPrt() As Integer

'Spot detail record information
Dim hmSdfPsf As Integer
Dim hmSdf As Integer        'Spot detail file handle
Dim hmPsf As Integer
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim tmSdfSrchKey5 As LONGKEY0    'SDF record image (key 3)
Dim imSdfRecLen As Integer  'SDF record length
Dim tmSpotFindInfo() As SPOTACTIONINFO
Private tmCntrFindSpot() As CNTRFINDSPOT
'Contract header record information
Dim hmCHF As Integer        'Contract file handle
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey0 As LONGKEY0
Dim tmChfSrchKey1 As CHFKEY1

'Contract header record information
Dim hmClf As Integer        'Contract file handle
Dim imClfRecLen As Integer  'CHF record length
Dim tmClf As CLF            'CHF record image
Dim tmClfSrchKey0 As CLFKEY0
Dim tmFindClf() As CLFLIST

'Contract header record information
Dim hmCff As Integer        'Contract file handle
Dim tmFindCFF() As CFFLIST

'Contract header record information
Dim hmCif As Integer        'Contract file handle
Dim imCifRecLen As Integer  'CHF record length
Dim tmCif As CIF            'CHF record image
Dim tmCifSrchKey0 As LONGKEY0

'Contract header record information
Dim hmCpf As Integer        'Contract file handle
Dim imCpfRecLen As Integer  'CHF record length
Dim tmCpf As CPF            'CHF record image
Dim tmCpfSrchKey0 As LONGKEY0

'Buyer record information
Private hmPnf As Integer        'Contract file handle
Private imPnfRecLen As Integer  'CHF record length
Private tmPnf As PNF            'CHF record image
Private tmPnfSrchKey0 As INTKEY0

Private hmSmf As Integer
Private imSmfRecLen As Integer  'CHF record length
Private tmSmf As SMF            'CHF record image
Private tmSmfSrchKey2 As LONGKEY0
Private tmSmfSrchKey4 As SMFKEY4

'Game Schedule
Dim hmGsf As Integer
Dim tmGsf As GSF
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer
Dim tmGsfInfo() As GSFINFO

Private hmVsf As Integer

Private imLastAdvtCntrColSorted As Integer
Private imLastAdvtCntrSort As Integer
Private lmLastClickedAdvtCntrRow As Long
Private lmScrollTop As Long

Private imLastLineColSorted As Integer
Private imLastLineSort As Integer
Private lmLastClickedLineRow As Long

Private imDefaultMatchIndex As Integer
Private imMatchIndex As Integer

Private imDefaultSearchIndex As Integer
Private imSearchIndex As Integer

Private imAdfCode As Integer
Private imAgfCode As Integer
Private imMnfComp As Integer
Private imSlfCode As Integer
Private smContractType As String
Private smBuyerName As String
Private lmFromRate As Long
Private lmToRate As Long
Private imRdfCode As Integer
Private lmCntrNo As Long
Private smEstNo As String

Private bmFirstCallToVpfFind As Boolean

Private rst_Chf As ADODB.Recordset

Const ADVTINDEX = 0
Const CNTRINDEX = 1
Const CASHTRADEINDEX = 2
Const HOLDINDEX = 3
Const SORTINDEX = 4
Const CHFCODEINDEX = 5

Const LINEINDEX = 0
Const VEHICLEINDEX = 1
Const DAYPARTINDEX = 2
Const DATESINDEX = 3
Const NOPERWKINDEX = 4
Const MOINDEX = 5
Const TUINDEX = 6
Const WEINDEX = 7
Const THINDEX = 8
Const FRINDEX = 9
Const SAINDEX = 10
Const SUINDEX = 11
Const RATEINDEX = 12
Const LINESORTINDEX = 13
Const VEFCODEINDEX = 14
Const CLFCODEINDEX = 15

Const MOSPOTINDEX = 0
Const TUSPOTINDEX = 1
Const WESPOTINDEX = 2
Const THSPOTINDEX = 3
Const FRSPOTINDEX = 4
Const SASPOTINDEX = 5
Const SUSPOTINDEX = 6
Const DATAEXISTINDEX = 7






Private Sub cbcMatch_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
        gSetMousePointer grdSpots, grdSpots, vbHourglass
        imChgMode = True
        ilLoopCount = 0
        Do
            If ilLoopCount > 0 Then
                If cbcMatch.ListIndex >= 0 Then
                    cbcMatch.Text = cbcMatch.List(cbcMatch.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcMatch.Text <> "" Then
                gManLookAhead cbcMatch, imBSMode, imDefaultMatchIndex
            End If
            imMatchIndex = cbcMatch.ListIndex
        Loop While imMatchIndex <> cbcMatch.ListIndex
        If imMatchIndex >= 0 Then
            lacSearch.Visible = False
            cbcSearch.Visible = False
            lacFrom.Visible = False
            edcFrom.Visible = False
            lacTo.Visible = False
            edcTo.Visible = False
            edcSearch.Visible = False
            grdAdvtCntr.TextMatrix(0, ADVTINDEX) = "Advertiser/Product"
            Select Case cbcMatch.ItemData(imMatchIndex)
                Case 1  'Advertiser
                    cbcSearch.Clear
                    mAdvtPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                    grdAdvtCntr.TextMatrix(0, ADVTINDEX) = "Product"
                Case 2  'Agency
                    cbcSearch.Clear
                    mAgencyPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 3  'Product Protection
                    cbcSearch.Clear
                    mProductProtectionPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 4  'Salesperson
                    cbcSearch.Clear
                    mSPersonPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 5  'Contract Type
                    cbcSearch.Clear
                    mContractTypePop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 6  'Buyer Name'
                    cbcSearch.Clear
                    mBuyerPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 7  'Rate Range
                    lacFrom.Visible = True
                    edcFrom.Visible = True
                    lacTo.Visible = True
                    edcTo.Visible = True
                Case 8  'Daypart
                    cbcSearch.Clear
                    mDaypartPop
                    lacSearch.Visible = True
                    cbcSearch.Visible = True
                Case 9  'Contract #
                    lacSearch.Visible = True
                    edcSearch.Visible = True
                Case 10  'Contract #
                    lacSearch.Visible = True
                    edcSearch.Visible = True
            End Select
        End If
        mClearGrid
        gSetMousePointer grdSpots, grdSpots, vbDefault
        gSetMousePointer grdAdvtCntr, grdLines, vbDefault
        Screen.MousePointer = vbDefault
        mSetCommands
        imChgMode = False
    End If
End Sub

Private Sub cbcMatch_Click()
    imDefaultMatchIndex = cbcMatch.ListIndex
    cbcMatch_Change
End Sub

Private Sub cbcMatch_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cbcMatch
    imDefaultMatchIndex = imMatchIndex
End Sub

Private Sub cbcMatch_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcMatch_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcMatch.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcSearch_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
        gSetMousePointer grdSpots, grdSpots, vbHourglass
        imChgMode = True
        ilLoopCount = 0
        Do
            If ilLoopCount > 0 Then
                If cbcSearch.ListIndex >= 0 Then
                    cbcSearch.Text = cbcSearch.List(cbcSearch.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcSearch.Text <> "" Then
                gManLookAhead cbcSearch, imBSMode, imDefaultSearchIndex
            End If
            imSearchIndex = cbcSearch.ListIndex
        Loop While imSearchIndex <> cbcSearch.ListIndex
        mClearGrid
        gSetMousePointer grdSpots, grdSpots, vbDefault
        gSetMousePointer grdAdvtCntr, grdLines, vbDefault
        Screen.MousePointer = vbDefault
        mSetCommands
        imChgMode = False
    End If
End Sub

Private Sub cbcSearch_Click()
    imDefaultSearchIndex = cbcSearch.ListIndex
    cbcSearch_Change
End Sub

Private Sub cbcSearch_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cbcMatch
    imDefaultSearchIndex = imSearchIndex
End Sub

Private Sub cbcSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcSearch_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSearch.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub ckcShowISCI_Click()
    mFillSpotGrid
End Sub

Private Sub ckcShowISCI_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcVehicleNonSports_Click()
    Dim llValue As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer

    If lbcVehicles.ListCount <= 0 Then
        Exit Sub
    End If
    If Not imSetAll Then
        Exit Sub
    End If
    ilValue = False
    If ckcVehicleNonSports.Value = vbChecked Then
        ilValue = True
    End If
    If (((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) And (ckcVehicleSports.Value = vbChecked) And (ilValue = True)) Or (((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) And (ilValue = True)) Then
        imAllClicked = True
        llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    Else
        imAllClicked = True
        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
            slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilIndex = gBinarySearchVef(ilVefCode)
            If ilIndex <> -1 Then
                If tgMVef(ilIndex).sType <> "G" Then
                    lbcVehicles.Selected(ilLoop) = ilValue
                End If
            End If
        Next ilLoop
        imAllClicked = False
    End If
    mClearGrid
End Sub

Private Sub ckcVehicleNonSports_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub ckcVehicleSports_Click()
    Dim llValue As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer

    If lbcVehicles.ListCount <= 0 Then
        Exit Sub
    End If
    If Not imSetAll Then
        Exit Sub
    End If
    ilValue = False
    If ckcVehicleSports.Value = vbChecked Then
        ilValue = True
    End If
    If (ckcVehicleNonSports.Value = vbChecked) And (ilValue = True) Then
        imAllClicked = True
        llRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    Else
        imAllClicked = True
        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
            slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilIndex = gBinarySearchVef(ilVefCode)
            If ilIndex <> -1 Then
                If tgMVef(ilIndex).sType = "G" Then
                    lbcVehicles.Selected(ilLoop) = ilValue
                End If
            End If
        Next ilLoop
        imAllClicked = False
    End If
    mClearGrid
End Sub

Private Sub ckcVehicleSports_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub


Private Sub cmcAdvtCntr_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slNameCode As String
    Dim slCode As String

    
    mClearGrid
    If imMatchIndex = -1 Then
        MsgBox "No 'Match On' has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
        Exit Sub
    End If
    If lbcVehicles.SelCount <= 0 Then
        MsgBox "No Vehicle(s) has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
        Exit Sub
    End If
    If Not mCheckFromDateTime() Then
        Exit Sub
    End If
    Select Case cbcMatch.ItemData(imMatchIndex)
        Case 1  'Advertiser
            If imSearchIndex < 0 Then
                MsgBox "No Advertiser has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            slNameCode = tgAdvertiser(imSearchIndex).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imAdfCode = Val(slCode)
        Case 2  'Agency
            If imSearchIndex < 0 Then
                MsgBox "No Adency has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            slNameCode = tgAgency(imSearchIndex).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imAgfCode = Val(slCode)
        Case 3  'Product Protection
            If imSearchIndex < 0 Then
                MsgBox "No Product Protection has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            slNameCode = tmProductProtectionCode(imSearchIndex).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMnfComp = Val(slCode)
        Case 4  'Salesperson
            If imSearchIndex < 0 Then
                MsgBox "No Salesperson has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            slNameCode = tmSalesperson(imSearchIndex).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imSlfCode = Val(slCode)
        Case 5  'Contract Type
            If imSearchIndex < 0 Then
                MsgBox "No Contract Type has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            Select Case cbcSearch.ItemData(imSearchIndex)
                Case 1  'Order
                    smContractType = "C"
                Case 2  'Remnant
                    smContractType = "T"
                Case 3  'Pre Inquiry
                    smContractType = "Q"
                Case 4  'Direct Response
                    smContractType = "R"
                Case 5  'Promo
                    smContractType = "M"
                Case 6  'PSA
                    smContractType = "S"
                Case 7  'Reservation
                    smContractType = "V"
            End Select
        Case 6  'Buyer Name'
            If imSearchIndex < 0 Then
                MsgBox "No Buyer has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            smBuyerName = Trim$(cbcSearch.List(imSearchIndex))
        Case 7  'Rate Range
            lacFrom.Visible = True
            edcFrom.Visible = True
            lacTo.Visible = True
            edcTo.Visible = True
            If (Trim$(edcFrom.Text) = "") Then
                MsgBox "From Rate Range Not Defined", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            If (Trim$(edcTo.Text) = "") Then
                MsgBox "From Rate Range Not Defined", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            lmFromRate = gStrDecToLong(edcFrom.Text, 2)
            lmToRate = gStrDecToLong(edcTo.Text, 2)
        Case 8  'Daypart
            If imSearchIndex < 0 Then
                MsgBox "No Daypart Name has been selected", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            imRdfCode = cbcSearch.ItemData(imSearchIndex)
        Case 9  'Contract #
            If (Trim$(edcSearch.Text) = "") Then
                MsgBox "Contract # Not Defined", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            lmCntrNo = Val(edcSearch.Text)
        Case 10  'Estimate #
            If (Trim$(edcSearch.Text) = "") Then
                MsgBox "Estimate # Not Defined", vbOKOnly + vbApplicationModal, "Get Advertiser"
                Exit Sub
            End If
            smEstNo = UCase(Trim$(edcSearch.Text))
    End Select
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    ilRet = mBuildSpots("A")
    grdAdvtCntr.Redraw = False
    grdAdvtCntr.Row = 0
    For llCol = ADVTINDEX To CNTRINDEX Step 1
        grdAdvtCntr.Col = llCol
        grdAdvtCntr.CellBackColor = LIGHTBLUE
    Next llCol
    grdAdvtCntr.RowHeight(0) = fgBoxGridH + 15
    llRow = grdAdvtCntr.FixedRows
    For ilLoop = 0 To UBound(tmSpotFindInfo) - 1 Step 1
        tmChfSrchKey0.lCode = tmSpotFindInfo(ilLoop).lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            '2/17/18: Allow programmatic Buys
            tmChf.sSource = ""
            If gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode(), tmChf.sSource) Then
                ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
                If ilAdf <> -1 Then
                    If llRow >= grdAdvtCntr.Rows Then
                        grdAdvtCntr.AddItem ""
                    End If
                    grdAdvtCntr.RowHeight(llRow) = fgBoxGridH + 15
                    For llCol = ADVTINDEX To HOLDINDEX Step 1
                        grdAdvtCntr.Row = llRow
                        grdAdvtCntr.Col = llCol
                        grdAdvtCntr.CellBackColor = vbWhite
                        grdAdvtCntr.CellForeColor = vbBlue
                    Next llCol
                    If Trim$(tmChf.sProduct) <> "" Then
                        If cbcMatch.ItemData(imMatchIndex) = 1 Then
                            grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) = Trim$(tmChf.sProduct)
                        Else
                            grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) = Trim$(tgCommAdf(ilAdf).sName) & "/" & Trim$(tmChf.sProduct)
                        End If
                    Else
                        grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                    End If
                    grdAdvtCntr.TextMatrix(llRow, CNTRINDEX) = Trim$(str(tmChf.lCntrNo))
                    If tmChf.iPctTrade = 0 Then
                        grdAdvtCntr.TextMatrix(llRow, CASHTRADEINDEX) = "C"
                    ElseIf tmChf.iPctTrade = 100 Then
                        grdAdvtCntr.TextMatrix(llRow, CASHTRADEINDEX) = "T"
                    Else
                        grdAdvtCntr.TextMatrix(llRow, CASHTRADEINDEX) = Trim$(str(tmChf.iPctTrade))
                    End If
                    If tmChf.sStatus = "H" Then
                        grdAdvtCntr.TextMatrix(llRow, HOLDINDEX) = "Hold"
                    ElseIf tmChf.sStatus = "O" Then
                        grdAdvtCntr.TextMatrix(llRow, HOLDINDEX) = "Order"
                    Else
                        grdAdvtCntr.TextMatrix(llRow, HOLDINDEX) = "Proposal"
                    End If
                    
                    grdAdvtCntr.TextMatrix(llRow, CHFCODEINDEX) = Trim$(str(tmSpotFindInfo(ilLoop).lChfCode))
                    llRow = llRow + 1
                End If
            End If
        End If
    Next ilLoop
    mAdvtCntrSortCol ADVTINDEX
    grdAdvtCntr.Row = 0
    grdAdvtCntr.Col = CHFCODEINDEX
    If grdAdvtCntr.Rows = grdAdvtCntr.FixedRows + 1 Then
        If grdAdvtCntr.TextMatrix(grdAdvtCntr.FixedRows, ADVTINDEX) <> "" Then
            lmHighlightRow = grdAdvtCntr.FixedRows
            mPaintRowColor
            mFillLineGrid
        End If
    End If
    grdAdvtCntr.Redraw = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
End Sub


Private Sub cmcAdvtCntr_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcFDDropDown_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcFDDropDown(Index).SelStart = 0
    edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    edcFDDropDown(Index).SetFocus
End Sub

Private Sub cmcFDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
        slStr = edcFDDropDown(Index).Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = Index + 1
    plcCalendar.Move edcFDDropDown(Index).Left, edcFDDropDown(Index).Top + edcFDDropDown(Index).Height
    gCtrlGotFocus ActiveControl

End Sub



Private Sub cmcRefresh_Click()
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdLines, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    mClearSpotGrid
    mGetCntrSpots
    mFillSpotGrid
    Screen.MousePointer = vbDefault
    gSetMousePointer grdLines, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
End Sub

Private Sub edcFDDropDown_Change(Index As Integer)
    Dim slStr As String
    
    slStr = edcFDDropDown(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mClearGrid
    mDatePop
    mSetCommands
End Sub

Private Sub edcFDDropDown_GotFocus(Index As Integer)
    Dim slStr As String
    
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
    End If
    imDateBox = Index + 1
    If (edcFDDropDown(Index).Text = "") And (Index = 1) Then
        slStr = edcFDDropDown(0).Text
        If slStr <> "" Then
            edcFDDropDown(Index).Text = gObtainNextSunday(slStr)
        End If
    End If
    plcCalendar.Move edcFDDropDown(Index).Left, edcFDDropDown(Index).Top + edcFDDropDown(Index).Height
    gCtrlGotFocus ActiveControl

End Sub

Private Sub edcFDDropDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcFDDropDown_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcFDDropDown(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcFDDropDown_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcFDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDDropDown(Index).Text = slDate
            End If
        End If
        edcFDDropDown(Index).SelStart = 0
        edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcFDDropDown(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcFDDropDown(Index).Text = slDate
            End If
        End If
        edcFDDropDown(Index).SelStart = 0
        edcFDDropDown(Index).SelLength = Len(edcFDDropDown(Index).Text)
    End If
End Sub



Private Sub edcFrom_Change()
    mClearGrid
    mSetCommands
End Sub

Private Sub edcFrom_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus edcFrom
End Sub

Private Sub edcSearch_Change()
    mClearGrid
    mSetCommands
End Sub

Private Sub edcSearch_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus edcSearch
End Sub

Private Sub edcTo_Change()
    mClearGrid
    mSetCommands
End Sub

Private Sub edcTo_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus edcFrom
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
    Me.KeyPreview = True
    CntrFind.Refresh
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

Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        'cmcCancel_Click
    End If
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
    Dim slDate As String
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    bmFirstCallToVpfFind = True
    imTerminate = False
    imFirstActivate = True
    imChgMode = False
    imSetAll = True
    imAllClicked = False
    imDateBox = -1
    imBypassFocus = False
    imBypassAll = False

    imcPrt.Picture = IconTraf!imcPrinter.Picture    'IconTraf!imcCamera.Picture

    lmLastClickedAdvtCntrRow = -1
    lmScrollTop = grdAdvtCntr.FixedRows
    imLastAdvtCntrColSorted = -1
    imLastAdvtCntrSort = -1
    lmHighlightRow = -1
    imLastLineColSorted = -1
    imLastLineSort = -1
    lmLineHighlightRow = -1
    imMatchIndex = -1
    imSearchIndex = -1
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    'CntrFind.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'pbcAction.BorderStyle = 0
    'pbcAction.Visible = False
    gCenterForm CntrFind
    CntrFind.Top = CntrFind.Top - 150
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", CntrFind
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSdfPsf = hmSdf
    
    hmPsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPsf, "", sgDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Psf.Btr)", CntrFind
    On Error GoTo 0
    
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", CntrFind
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", CntrFind
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    
    hmCpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", CntrFind
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    
    hmPnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pnf.Btr)", CntrFind
    On Error GoTo 0
    imPnfRecLen = Len(tmPnf)
    
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", CntrFind
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", CntrFind
    On Error GoTo 0
    
    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", CntrFind
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    
    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", CntrFind
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", CntrFind
    On Error GoTo 0
    
    ReDim Preserve tmGsfInfo(0 To 0) As GSFINFO
    
    mInitBox
    mVehPop
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        ckcVehicleSports.Visible = False
        ckcVehicleNonSports.Caption = "All Vehicles"
    End If
    If imTerminate Then
        Exit Sub
    End If
    If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        ckcVehicleSports.Visible = False
        ckcVehicleNonSports.Caption = "All Vehicles"
    End If
    slDate = Format$(gNow(), "m/d/yy")
    If ((Asc(tgSpf.sUsingFeatures4) And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY) Then
        edcFDDropDown(0).Text = slDate
    Else
        edcFDDropDown(0).Text = Format$(gDateValue(slDate) + 1, "m/d/yy")
    End If
    edcFDDropDown(1).Text = gObtainNextSunday(edcFDDropDown(0).Text)
    
    cbcMatch.AddItem "Advertiser"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 1
    cbcMatch.AddItem "Agency"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 2
    cbcMatch.AddItem "Product Protection"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 3
    cbcMatch.AddItem "Salesperson"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 4
    cbcMatch.AddItem "Contract Type"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 5
    cbcMatch.AddItem "Buyer Name"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 6
    cbcMatch.AddItem "Spot Rate Range"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 7
    cbcMatch.AddItem "Daypart Name"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 8
    cbcMatch.AddItem "Contract #"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 9
    cbcMatch.AddItem "Estimate #"
    cbcMatch.ItemData(cbcMatch.NewIndex) = 10
    
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    
    mSetCommands
    
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
    Exit Sub
         
    On Error GoTo 0
    imTerminate = True
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim ilLoop
    Dim ilCol As Integer
    Dim llRet As Long

    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcFDDropDown(0).Left, edcFDDropDown(0).Top + edcFDDropDown(0).Height

    imBSMode = False
    imCalType = 0   'Standard


    cmcCancel.Top = CntrFind.Height - cmcCancel.Height - 120
    pbcAction.Top = 285
    pbcAction.Left = 60
    pbcAction.Width = fmAdjFactorW * pbcAction.Width
    pbcAction.Height = cmcCancel.Top - pbcAction.Top - 120
    grdAdvtCntr.Width = fmAdjFactorW * grdAdvtCntr.Width
    grdAdvtCntr.Left = pbcAction.Width - grdAdvtCntr.Width - 180
    cmcAdvtCntr.Left = grdAdvtCntr.Left
    gGrid_IntegralHeight grdAdvtCntr, fgBoxGridH + 15
    cmcAdvtCntr.Top = grdAdvtCntr.Top + grdAdvtCntr.Height + 90
    
    lbcVehicles.Width = grdAdvtCntr.Left - lbcVehicles.Left - 120
    lbcVehicles.Height = grdAdvtCntr.Height - ckcVehicleNonSports.Height
    ckcVehicleNonSports.Left = lbcVehicles.Left
    ckcVehicleNonSports.Top = lbcVehicles.Top + lbcVehicles.Height + 90
    ckcVehicleSports.Top = ckcVehicleNonSports.Top
    ckcVehicleSports.Left = ckcVehicleNonSports.Left + ckcVehicleNonSports.Width + 180
    
    grdLines.Width = pbcAction.Width - lbcLineDates.Width - 360
    grdLines.Left = lbcLineDates.Left + lbcLineDates.Width + 90
    If ckcVehicleNonSports.Top + ckcVehicleNonSports.Height > cmcAdvtCntr.Top + cmcAdvtCntr.Height Then
        grdLines.Top = ckcVehicleNonSports.Top + ckcVehicleNonSports.Height + 90
    Else
        grdLines.Top = cmcAdvtCntr.Top + cmcAdvtCntr.Height + 90
    End If
    grdLines.Height = (pbcAction.Height - (grdLines.Top)) / 2
    
    grdSpots.Width = grdLines.Width
    grdSpots.Left = grdLines.Left
    
    cmcCancel.Left = CntrFind.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    cmcRefresh.Top = cmcCancel.Top
    cmcRefresh.Left = CntrFind.Width / 2 - cmcCancel.Width / 2 + cmcCancel.Width

    mSetGridColumns
    mSetGridTitles
    'gGrid_IntegralHeight grdAdvtCntr, fgBoxGridH + 15
    gGrid_IntegralHeight grdLines, fgBoxGridH + 15
    grdSpots.Top = grdLines.Top + grdLines.Height + 180
    grdSpots.Height = pbcAction.Height - grdSpots.Top - 60
    gGrid_IntegralHeight grdSpots, fgBoxGridH + 15
    lacLineInfo.Top = grdLines.Top
    lacLineInfo.Left = 60
    lbcLineDates.Height = grdLines.Height - lacLineInfo.Height
    lbcLineDates.Top = grdLines.Top + grdLines.Height - lbcLineDates.Height
    lacSpotInfo.Top = grdSpots.Top
    lacSpotInfo.Left = 60
    lbcSpotDates.Height = grdSpots.Height - lacSpotInfo.Height - ckcShowISCI.Height - 120
    lbcSpotDates.Top = lacSpotInfo.Top + lacSpotInfo.Height + 60
    ckcShowISCI.Top = lbcSpotDates.Top + lbcSpotDates.Height + 60
    ckcShowISCI.Left = lbcSpotDates.Left
    imcPrt.Move Me.Width - (3 * imcPrt.Width / 2), Me.Height - imcPrt.Height
    pbcPrinting.Move Me.Width / 2 - pbcPrinting.Width / 2, Me.Height / 2 - pbcPrinting.Height / 2
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
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
       
    igManUnload = YES
    Unload CntrFind
    igManUnload = NO
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase imMoPrt
    Erase imTuPrt
    Erase imWePrt
    Erase imThPrt
    Erase imFrPrt
    Erase imSaPrt
    Erase imSuPrt
    Erase imSelectedVefCode
    Erase tmSpotFindInfo
    Erase tmUserVehicle
    Erase tmProductProtectionCode
    Erase tmSalesperson
    Erase tmFindClf
    Erase tmFindCFF
    Erase tmCntrFindSpot
    Erase tmTeam
    Erase tmGsfInfo
    btrDestroy hmSdfPsf
    btrDestroy hmSdf
    btrDestroy hmPsf
    btrDestroy hmGsf
    btrDestroy hmSmf
    btrDestroy hmVsf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmCif
    btrDestroy hmCpf
    btrDestroy hmPnf
    
    rst_Chf.Close
    
    Set CntrFind = Nothing   'Remove data segment

End Sub
Private Sub grdAdvtCntr_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub grdLines_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub grdLines_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long

    For llRow = 0 To grdLines.FixedRows - 1 Step 1
        llHeight = grdLines.RowHeight(llRow)
    Next llRow
    grdLines.ToolTipText = ""
    If Y <= llHeight Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdLines, X, Y, llRow, llCol)
    If (llCol = VEHICLEINDEX) Or (llCol = DAYPARTINDEX) Or (llCol = DATESINDEX) Then
        grdLines.ToolTipText = Trim$(grdLines.TextMatrix(llRow, llCol))
    End If
End Sub

Private Sub grdLines_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdLines.RowHeight(0) Then
        If (grdLines.MouseCol <= DATESINDEX) Or (grdLines.MouseCol = RATEINDEX) Then
            grdLines.Col = grdLines.MouseCol
            mLineSortCol grdLines.Col
        End If
        grdLines.Row = 0
        grdLines.Col = CLFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdLines, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdLines.FixedRows Then
        Exit Sub
    End If
    mClearSpotGrid
    If grdLines.TextMatrix(llCurrentRow, LINEINDEX) = "" Then
        lmLineHighlightRow = -1
        imcPrt.Visible = False
    Else
        lmLineHighlightRow = llCurrentRow
        imcPrt.Visible = True
    End If
    mPaintLineRowColor
    If lmLineHighlightRow = -1 Then
        Exit Sub
    End If
    mDateSpotPop
    mFillSpotGrid
    Screen.MousePointer = vbDefault
    gSetMousePointer grdLines, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
End Sub

Private Sub grdSpots_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub grdSpots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long

    For llRow = 0 To grdSpots.FixedRows - 1 Step 1
        llHeight = grdSpots.RowHeight(llRow)
    Next llRow
    grdSpots.ToolTipText = ""
    If Y <= llHeight Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdSpots, X, Y, llRow, llCol)
    grdSpots.ToolTipText = Trim$(grdSpots.TextMatrix(llRow, llCol))

End Sub

Private Sub imcPrt_Click()
    Dim ilCurrentLineNo As Integer
    Dim ilLinesPerPage As Integer
    Dim slRecord As String
    Dim slHeading As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim llFlightStartDate As Long
    Dim llFlightEndDate As Long
    Dim ilCff As Integer
    Dim slRate As String
    Dim ilLineNo As Integer
    Dim ilGameNo As Integer
    Dim llPrintedDate As Long
    Dim ilSdf As Integer
    Dim slStr As String
    Dim slISCI As String
    Dim ilVef As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim llLoopDate As Long
    Dim ilMaxSpots As Integer
    Dim ilPass As Integer
    Dim ilIncludeSpot As Integer
    Dim ilPrtOrient As Integer
    Dim slFlightInfo As String
    Dim ilWkFound As Integer
    Dim ilMaxPrtWidth As Integer
    ReDim ilCurrentX(0 To 6) As Integer
    
    If (lmHighlightRow < grdAdvtCntr.FixedRows) Or (lmLineHighlightRow < grdLines.FixedRows) Then
        Exit Sub
    End If
    If (grdAdvtCntr.TextMatrix(lmHighlightRow, ADVTINDEX) = "") Or (grdLines.TextMatrix(lmLineHighlightRow, LINEINDEX) = "") Then
        Exit Sub
    End If
    
    pbcPrinting.Visible = True
    DoEvents
    
    ilCurrentLineNo = 0
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    'determine tab column positions
    ilPrtOrient = Printer.Orientation
    Printer.Orientation = vbPRORLandscape
    ilRet = 0
    On Error GoTo imcPrtErr:
    slHeading = "Counterpoint 'Find' Information for " & Trim$(tgUrf(0).sRept) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    '6/6/16: Replaced GoSub
    'GoSub mHeading1
    mHeading1 ilRet, slHeading, ilCurrentLineNo
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output Contract Information
    slRecord = "Advertiser: " & Trim$(grdAdvtCntr.TextMatrix(lmHighlightRow, ADVTINDEX)) & " Contract #: " & Trim$(grdAdvtCntr.TextMatrix(lmHighlightRow, CNTRINDEX))
    If Trim$(grdAdvtCntr.TextMatrix(lmHighlightRow, CASHTRADEINDEX)) = "C" Then
        slRecord = slRecord & " " & "Cash"
    ElseIf Trim$(grdAdvtCntr.TextMatrix(lmHighlightRow, CASHTRADEINDEX)) = "T" Then
        slRecord = slRecord & " " & "Trade"
    Else
        slRecord = slRecord & " " & Trim$(grdAdvtCntr.TextMatrix(lmHighlightRow, CASHTRADEINDEX))
    End If
    If grdAdvtCntr.TextMatrix(lmHighlightRow, HOLDINDEX) = "H" Then
        slRecord = slRecord & " " & "Hold"
    ElseIf grdAdvtCntr.TextMatrix(lmHighlightRow, HOLDINDEX) = "O" Then
        slRecord = slRecord & " " & "Order"
    Else
        slRecord = slRecord & " " & "Proposal"
    End If
    '6/6/16: Replaced GoSub
    'GoSub mLineOutput
    mLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output line information
    slRecord = "Line #: " & Trim$(grdLines.TextMatrix(lmLineHighlightRow, LINEINDEX)) & " Vehicle: " & Trim$(grdLines.TextMatrix(lmLineHighlightRow, VEHICLEINDEX)) & " Daypart Name:" & Trim$(grdLines.TextMatrix(lmLineHighlightRow, DAYPARTINDEX))
    '6/6/16: Replaced GoSub
    'GoSub mLineOutput
    mLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    ilLineNo = grdLines.TextMatrix(lmLineHighlightRow, LINEINDEX)
    If lbcSpotDates.ItemData(lbcSpotDates.ListIndex) = 0 Then
        llSDate = gDateValue(lbcLineDates.List(0))
        llEDate = gDateValue(lbcLineDates.List(lbcLineDates.ListCount - 1))
    Else
        ilGameNo = tmGsfInfo(lbcSpotDates.ItemData(0) - 1).iGameNo
        llSDate = 0
        llEDate = lbcSpotDates.ListCount - 1
    End If
    For llDate = llSDate To llEDate Step 7
        If ilGameNo = 0 Then
            ilWkFound = False
            For ilCff = LBound(tmFindCFF) To UBound(tmFindCFF) - 1 Step 1
                If tmFindCFF(ilCff).CffRec.iClfLine = ilLineNo Then
                    gUnpackDateLong tmFindCFF(ilCff).CffRec.iStartDate(0), tmFindCFF(ilCff).CffRec.iStartDate(1), llFlightStartDate
                    gUnpackDateLong tmFindCFF(ilCff).CffRec.iEndDate(0), tmFindCFF(ilCff).CffRec.iEndDate(1), llFlightEndDate
                    If llFlightEndDate < llFlightStartDate Then
                        slRecord = "CBS"
                        ilWkFound = True
                    Else
                        If (llDate >= llFlightStartDate) And (llDate <= llFlightEndDate) Then
                            ilWkFound = True
                            slRecord = Format$(llSDate, "m/d/yy") & "-" & Format$(llSDate + 6, "m/d/yy")
                            If tmFindCFF(ilCff).CffRec.sDyWk <> "D" Then
                                slRecord = slRecord & " #/Wk: " & Trim$(str(tmFindCFF(ilCff).CffRec.iSpotsWk)) & " Air Days"
                                For ilDay = 0 To 6 Step 1
                                    If llFlightStartDate + 6 <= llFlightEndDate Then
                                        If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                            slRecord = slRecord & " " & Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su")
                                        End If
                                    Else
                                        If llFlightStartDate + ilDay <= llFlightEndDate Then
                                            If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                slRecord = slRecord & " " & Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su")
                                            End If
                                        End If
                                    End If
                                Next ilDay
                            Else
                                slRecord = slRecord & " Spots/Day"
                                For ilDay = 0 To 6 Step 1
                                    If llFlightStartDate + 6 <= llFlightEndDate Then
                                        If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                            slRecord = slRecord & " " & Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su") & ":" & Trim$(str(tmFindCFF(ilCff).CffRec.iDay(ilDay)))
                                        End If
                                    Else
                                        If llFlightStartDate + ilDay <= llFlightEndDate Then
                                            If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                slRecord = slRecord & " " & Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su") & ":" & Trim$(str(tmFindCFF(ilCff).CffRec.iDay(ilDay)))
                                            End If
                                        End If
                                    End If
                                Next ilDay
                            End If
                            Select Case tmFindCFF(ilCff).CffRec.sPriceType
                                Case "T"    'True
                                    slRate = "$" & gLongToStrDec(tmFindCFF(ilCff).CffRec.lActPrice, 2)
                                Case "N"    'No Charge
                                    slRate = "N/C"
                                Case "M"    'MG Line
                                    slRate = "MG"
                                Case "B"    'Bonus
                                    slRate = "Bonus"
                                Case "S"    'Spinoff
                                    slRate = "Spinoff"
                                Case "P"    'Package
                                    slRate = "$" & gLongToStrDec(tmFindCFF(ilCff).CffRec.lActPrice, 2)
                                Case "R"    'Recapturable
                                    slRate = "Recapturable"
                                Case "A"    'ADU
                                    slRate = "ADU"
                            End Select
                            slRecord = slRecord & " " & slRate
                            '6/6/16: Replaced GoSub
                            'GoSub mLineOutput
                            mLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage
                            If ilRet <> 0 Then
                                Printer.EndDoc
                                On Error GoTo 0
                                pbcPrinting.Visible = False
                                Exit Sub
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next ilCff
        Else
            ilWkFound = True
        End If
        'Print Flight Info
        If ilWkFound Then
            For ilDay = 0 To 6 Step 1
                ilCurrentX(ilDay) = -1
                If ilGameNo = 0 Then
                    slRecord = Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su") & ": " & Format$(llDate + ilDay, "m/d/yy")
                Else
                    'slRecord = Trim$(lbcSpotDates.List(llDate + ilDay))
                    If llDate + ilDay < lbcSpotDates.ListCount Then
                        slRecord = "#" & Trim$(str(tmGsfInfo(lbcSpotDates.ItemData(llDate + ilDay) - 1).iGameNo))
                        slRecord = slRecord & " " & Trim$(tmGsfInfo(lbcSpotDates.ItemData(llDate + ilDay) - 1).sVisitName) & " @ " & Trim$(tmGsfInfo(lbcSpotDates.ItemData(llDate + ilDay) - 1).sHomeName)
                    Else
                        slRecord = ""
                    End If
                End If
                '6/6/16: Replaced GoSub
                'GoSub mPartialLineOutput
                mPartialLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage, ilDay, ilCurrentX(), ilMaxPrtWidth
                If ilRet <> 0 Then
                    Printer.EndDoc
                    On Error GoTo 0
                    pbcPrinting.Visible = False
                    Exit Sub
                End If
                ilCurrentX(ilDay) = Printer.CurrentX
            Next ilDay
            For ilDay = 5 To 0 Step -1
                ilCurrentX(ilDay + 1) = ilCurrentX(ilDay)
            Next ilDay
            ilCurrentX(0) = 0
            ilMaxPrtWidth = ilCurrentX(1) - ilCurrentX(0)
            If ilCurrentX(2) - ilCurrentX(1) < ilMaxPrtWidth Then
                ilMaxPrtWidth = ilCurrentX(2) - ilCurrentX(1)
            End If
            If ilCurrentX(3) - ilCurrentX(2) < ilMaxPrtWidth Then
                ilMaxPrtWidth = ilCurrentX(3) - ilCurrentX(2)
            End If
            If ilCurrentX(4) - ilCurrentX(3) < ilMaxPrtWidth Then
                ilMaxPrtWidth = ilCurrentX(4) - ilCurrentX(3)
            End If
            If ilCurrentX(5) - ilCurrentX(4) < ilMaxPrtWidth Then
                ilMaxPrtWidth = ilCurrentX(5) - ilCurrentX(4)
            End If
        End If
        'Get spots
        For ilPass = 1 To 6 Step 1
            ReDim imMoPrt(0 To 0) As Integer
            ReDim imTuPrt(0 To 0) As Integer
            ReDim imWePrt(0 To 0) As Integer
            ReDim imThPrt(0 To 0) As Integer
            ReDim imFrPrt(0 To 0) As Integer
            ReDim imSaPrt(0 To 0) As Integer
            ReDim imSuPrt(0 To 0) As Integer
            ilMaxSpots = 0
            ilDay = 0
            For llLoopDate = llDate To llDate + 6 Step 1
                If ilGameNo <> 0 Then
                    If llLoopDate < lbcSpotDates.ListCount Then
                        ilGameNo = tmGsfInfo(lbcSpotDates.ItemData(llLoopDate) - 1).iGameNo
                    Else
                        Exit For
                    End If
                End If
                For ilSdf = 0 To UBound(tmCntrFindSpot) - 1 Step 1
                    If tmCntrFindSpot(ilSdf).iLineNo = ilLineNo Then
                        ilIncludeSpot = False
                        Select Case ilPass
                            Case 1
                                If tmCntrFindSpot(ilSdf).sStatus = "S" Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                            Case 2
                                If (tmCntrFindSpot(ilSdf).sStatus = "G") Or (tmCntrFindSpot(ilSdf).sStatus = "O") Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iMissedGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sMissedDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                            Case 3
                                If tmCntrFindSpot(ilSdf).sStatus = "M" Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                            Case 4
                                If tmCntrFindSpot(ilSdf).sStatus = "X" Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                            Case 5
                                If tmCntrFindSpot(ilSdf).sStatus = "C" Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                            Case 6
                                If tmCntrFindSpot(ilSdf).sStatus = "H" Then
                                    If ilGameNo <> 0 Then
                                        If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                            ilIncludeSpot = True
                                        End If
                                    Else
                                        If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llLoopDate Then
                                            ilIncludeSpot = True
                                        End If
                                    End If
                                End If
                        End Select
                        If ilIncludeSpot Then
                            Select Case ilDay
                                Case 0
                                    mSavePrtIndex ilSdf, imMoPrt(), ilMaxSpots
                                Case 1
                                    mSavePrtIndex ilSdf, imTuPrt(), ilMaxSpots
                                Case 2
                                    mSavePrtIndex ilSdf, imWePrt(), ilMaxSpots
                                Case 3
                                    mSavePrtIndex ilSdf, imThPrt(), ilMaxSpots
                                Case 4
                                    mSavePrtIndex ilSdf, imFrPrt(), ilMaxSpots
                                Case 5
                                    mSavePrtIndex ilSdf, imSaPrt(), ilMaxSpots
                                Case 6
                                    mSavePrtIndex ilSdf, imSuPrt(), ilMaxSpots
                            End Select
                        End If
                    End If
                Next ilSdf
                ilDay = ilDay + 1
            Next llLoopDate
            If ilMaxSpots > 0 Then
                Select Case ilPass
                    Case 1
                        slStr = "Scheduled"
                    Case 2
                        slStr = "MG/Outside"
                    Case 3
                        slStr = "Missed"
                    Case 4
                        slStr = "Fill"
                    Case 5
                        slStr = "Canceled"
                    Case 6
                        slStr = "Hidden"
                End Select
                For ilDay = 0 To 6 Step 1
                    slRecord = ""
                    If ilDay = 3 Then
                        slRecord = slStr
                    End If
                    '6/6/16: Replaced GoSub
                    'GoSub mPartialLineOutput
                    mPartialLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage, ilDay, ilCurrentX(), ilMaxPrtWidth
                    If ilRet <> 0 Then
                        Printer.EndDoc
                        On Error GoTo 0
                        pbcPrinting.Visible = False
                        Exit Sub
                    End If
                Next ilDay
                mSetPrtIndex imMoPrt(), ilMaxSpots
                mSetPrtIndex imTuPrt(), ilMaxSpots
                mSetPrtIndex imWePrt(), ilMaxSpots
                mSetPrtIndex imThPrt(), ilMaxSpots
                mSetPrtIndex imFrPrt(), ilMaxSpots
                mSetPrtIndex imSaPrt(), ilMaxSpots
                mSetPrtIndex imSuPrt(), ilMaxSpots
                For ilLoop = 0 To ilMaxSpots - 1 Step 1
                    For ilDay = 0 To 6 Step 1
                        Select Case ilDay
                            Case 0
                                ilSdf = imMoPrt(ilLoop)
                            Case 1
                                ilSdf = imTuPrt(ilLoop)
                            Case 2
                                ilSdf = imWePrt(ilLoop)
                            Case 3
                                ilSdf = imThPrt(ilLoop)
                            Case 4
                                ilSdf = imFrPrt(ilLoop)
                            Case 5
                                ilSdf = imSaPrt(ilLoop)
                            Case 6
                                ilSdf = imSuPrt(ilLoop)
                        End Select
                        If ilSdf <> -1 Then
                            Select Case ilPass
                                Case 1
                                    slRecord = tmCntrFindSpot(ilSdf).sAirTime
                                Case 2
                                    slRecord = Trim$(tmCntrFindSpot(ilSdf).sAirDate) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirTime)
                                Case 3
                                    slRecord = tmCntrFindSpot(ilSdf).sAirTime
                                Case 4
                                    slRecord = tmCntrFindSpot(ilSdf).sAirTime
                                Case 5
                                    slRecord = tmCntrFindSpot(ilSdf).sAirTime
                                Case 6
                                    slRecord = tmCntrFindSpot(ilSdf).sAirTime
                            End Select
                        Else
                            slRecord = ""
                        End If
                        '6/6/16: Replaced GoSub
                        'GoSub mPartialLineOutput
                        mPartialLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage, ilDay, ilCurrentX(), ilMaxPrtWidth
                        If ilRet <> 0 Then
                            Printer.EndDoc
                            On Error GoTo 0
                            pbcPrinting.Visible = False
                            Exit Sub
                        End If
                    Next ilDay
                    If ilPass = 2 Then
                        For ilDay = 0 To 6 Step 1
                            Select Case ilDay
                                Case 0
                                    ilSdf = imMoPrt(ilLoop)
                                Case 1
                                    ilSdf = imTuPrt(ilLoop)
                                Case 2
                                    ilSdf = imWePrt(ilLoop)
                                Case 3
                                    ilSdf = imThPrt(ilLoop)
                                Case 4
                                    ilSdf = imFrPrt(ilLoop)
                                Case 5
                                    ilSdf = imSaPrt(ilLoop)
                                Case 6
                                    ilSdf = imSuPrt(ilLoop)
                            End Select
                            If ilSdf <> -1 Then
                                ilVef = gBinarySearchVef(tmCntrFindSpot(ilSdf).iAirVefCode)
                                If (ilVef <> -1) Then
                                    If ilGameNo > 0 Then
                                        slRecord = "#" & Trim$(str(tmCntrFindSpot(ilSdf).iGameNo)) & " " & Trim$(tgMVef(ilVef).sName)
                                    Else
                                        slRecord = Trim$(tgMVef(ilVef).sName)
                                    End If
                                Else
                                    slRecord = ""
                                End If
                            Else
                                slRecord = ""
                            End If
                            '6/6/16: Replaced GoSub
                            'GoSub mPartialLineOutput
                            mPartialLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage, ilDay, ilCurrentX(), ilMaxPrtWidth
                            If ilRet <> 0 Then
                                Printer.EndDoc
                                On Error GoTo 0
                                pbcPrinting.Visible = False
                                Exit Sub
                            End If
                        Next ilDay
                    End If
                    If ckcShowISCI.Value = vbChecked Then
                        For ilDay = 0 To 6 Step 1
                            Select Case ilDay
                                Case 0
                                    ilSdf = imMoPrt(ilLoop)
                                Case 1
                                    ilSdf = imTuPrt(ilLoop)
                                Case 2
                                    ilSdf = imWePrt(ilLoop)
                                Case 3
                                    ilSdf = imThPrt(ilLoop)
                                Case 4
                                    ilSdf = imFrPrt(ilLoop)
                                Case 5
                                    ilSdf = imSaPrt(ilLoop)
                                Case 6
                                    ilSdf = imSuPrt(ilLoop)
                            End Select
                            If ilSdf <> -1 Then
                                If tmCntrFindSpot(ilSdf).sPtType <> "I" Then
                                    tmCntrFindSpot(ilSdf).sISCI = mObtainISCI(tmCntrFindSpot(ilSdf).sPtType, tmCntrFindSpot(ilSdf).lCopyCode)
                                    tmCntrFindSpot(ilSdf).sPtType = "I"
                                End If
                                slRecord = Trim$(tmCntrFindSpot(ilSdf).sISCI)
                            Else
                                slRecord = ""
                            End If
                            '6/6/16: Replaced GoSub
                            'GoSub mPartialLineOutput
                            mPartialLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage, ilDay, ilCurrentX(), ilMaxPrtWidth
                            If ilRet <> 0 Then
                                Printer.EndDoc
                                On Error GoTo 0
                                pbcPrinting.Visible = False
                                Exit Sub
                            End If
                        Next ilDay
                    End If
                Next ilLoop
                
            End If
        Next ilPass
    Next llDate
    Printer.EndDoc
    Printer.Orientation = ilPrtOrient
    On Error GoTo 0
    pbcPrinting.Visible = False
    Exit Sub
'mHeading1:  'Output file name and date
'    Printer.Print slHeading
'    If ilRet <> 0 Then
'        Return
'    End If
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Printer.Print " "
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
'mLineOutput:
'    If ilCurrentLineNo >= ilLinesPerPage Then
'        Printer.NewPage
'        If ilRet <> 0 Then
'            Return
'        End If
'        ilCurrentLineNo = 0
'        GoSub mHeading1
'        If ilRet <> 0 Then
'            Return
'        End If
'    End If
'    Printer.Print slRecord
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
'mPartialLineOutput:
'    If (ilCurrentLineNo >= ilLinesPerPage) And (ilDay = 0) Then
'        Printer.NewPage
'        If ilRet <> 0 Then
'            Return
'        End If
'        ilCurrentLineNo = 0
'        '6/6/16: Replaced GoSub
'        'GoSub mHeading1
'        mHeading1 ilRet, slHeading, ilCurrentLineNo
'        If ilRet <> 0 Then
'            Return
'        End If
'    End If
'    If ilCurrentX(ilDay) = -1 Then
'        If ilDay <> 6 Then
'            Printer.Print slRecord; Spc(5),  'Tab,
'        Else
'            Printer.Print slRecord
'            ilCurrentLineNo = ilCurrentLineNo + 1
'        End If
'    Else
'        Do While Printer.TextWidth(slRecord) > ilMaxPrtWidth
'            slRecord = Left(slRecord, Len(slRecord) - 1)
'        Loop
'        Printer.CurrentX = ilCurrentX(ilDay)
'        If ilDay <> 6 Then
'            Printer.Print slRecord,
'        Else
'            Printer.Print slRecord
'        End If
'    End If
'    Return
imcPrtErr:
    ilRet = Err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub

Private Sub lacFEndDate_Click()
    plcCalendar.Visible = False
End Sub


Private Sub lacFrom_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lacFStartDate_Click()
    plcCalendar.Visible = False
End Sub


Private Sub lacLineInfo_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lacMatchOn_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lacScreen_Click()
    plcCalendar.Visible = False
End Sub


Private Sub lacSearch_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lacSpotInfo_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lacTo_Click()
    plcCalendar.Visible = False
End Sub

Private Sub lbcLineDates_Click()
    If imChgMode Then
        Exit Sub
    End If
    mFillFlightsInLineGrid
    'If lbcSpotDates.ListIndex < 0 Then
        If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
            lbcSpotDates.ListIndex = lbcLineDates.ListIndex
            mFillSpotGrid
        End If
    'End If
    mSetCommands
    If lbcSpotDates.ListIndex >= 0 Then
        cmcCancel.Caption = "&Done"
    End If
End Sub

Private Sub lbcLineDates_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcSpotDates_Click()
    Dim llSDate As Long
    Dim ilDay As Integer
    Dim slDay As String
    
    If lmLineHighlightRow < grdLines.FixedRows Then
        Exit Sub
    End If
    If lbcSpotDates.ListIndex = -1 Then
        Exit Sub
    End If
    If lbcSpotDates.ItemData(lbcSpotDates.ListIndex) = 0 Then
        llSDate = gDateValue(lbcSpotDates.List(lbcSpotDates.ListIndex))
        For ilDay = 0 To 6 Step 1
            slDay = Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su")
            grdSpots.TextMatrix(0, ilDay) = slDay & " " & Format$(llSDate + ilDay, "m/d/yy")
        Next ilDay
    Else
        For ilDay = 0 To 6 Step 1
            If lbcSpotDates.ListIndex + ilDay < lbcSpotDates.ListCount Then
                slDay = "#" & Trim$(str(tmGsfInfo(lbcSpotDates.ItemData(lbcSpotDates.ListIndex + ilDay) - 1).iGameNo))
                slDay = slDay & " " & Trim$(tmGsfInfo(lbcSpotDates.ItemData(lbcSpotDates.ListIndex + ilDay) - 1).sVisitName) & " @ " & Trim$(tmGsfInfo(lbcSpotDates.ItemData(lbcSpotDates.ListIndex + ilDay) - 1).sHomeName)
            Else
                slDay = ""
            End If
            grdSpots.TextMatrix(0, ilDay) = slDay
        Next ilDay
    End If
    If imChgMode Then
        Exit Sub
    End If
    mFillSpotGrid
    mSetCommands
    cmcCancel.Caption = "&Done"
End Sub

Private Sub lbcSpotDates_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcVehicles_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcVehicleNonSports.Value = vbUnchecked
        ckcVehicleSports.Value = vbUnchecked
        imSetAll = True
    End If
    mClearGrid
    mSetCommands
End Sub

Private Sub lbcVehicles_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcVehicles_Scroll()
    plcCalendar.Visible = False
End Sub


Private Sub pbcAction_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mCompPop                        *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate competitive list      *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mProductProtectionPop()
'
'   mCompPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String      'Competitive name, saved to determine if changed
    Dim ilIndex As Integer      'Competitive name, saved to determine if changed
    ilfilter(0) = CHARFILTER
    slFilter(0) = "C"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilRet = gIMoveListBox(CntrFind, cbcSearch, tmProductProtectionCode(), smProductProtectionCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCompPopErr
        gCPErrorMsg ilRet, "mProductProtectionPop (gIMoveListBox)", CntrFind
        On Error GoTo 0
    End If
    Exit Sub
mCompPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    Dim ilSetVehicle As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilLoop As Integer
    
    lbcVehicles.Clear
    smUserVehicleTag = ""
    If rbcShow(0).Value = True Then
        ''ilRet = gPopUserVehicleBox(Spots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, Traffic!lbcUserVehicle)
        'ilRet = gPopUserVehicleBox(Spots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, tmUserVehicle(), smUserVehicleTag)
        If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSPORT + VEHVIRTUAL + ACTIVEVEH ' Airing and all conventional vehicles (except with Log) and Log
            ilRet = gPopUserVehicleBox(CntrFind, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
        Else
            llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH ' Airing and all conventional vehicles (except with Log) and Log
            ilRet = gPopUserVehicleBox(CntrFind, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
        End If
    Else
        llFilter = VEHPACKAGE + VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH ' Airing and all conventional vehicles (except with Log) and Log
        ilRet = gPopUserVehicleBox(CntrFind, llFilter, lbcVehicles, tmUserVehicle(), smUserVehicleTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", CntrFind
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imDateBox = 1 Then
        slStr = edcFDDropDown(0).Text
    ElseIf imDateBox = 2 Then
        slStr = edcFDDropDown(1).Text
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imDateBox = 1 Then
                    edcFDDropDown(0).Text = Format$(llDate, "m/d/yy")
                    edcFDDropDown(0).SelStart = 0
                    edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
                    imBypassFocus = True
                    edcFDDropDown(0).SetFocus
                    Exit Sub
                ElseIf imDateBox = 2 Then
                    edcFDDropDown(1).Text = Format$(llDate, "m/d/yy")
                    edcFDDropDown(1).SelStart = 0
                    edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
                    imBypassFocus = True
                    edcFDDropDown(1).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox = 1 Then
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        edcFDDropDown(1).SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub cmcCalDn_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(0).SelStart = 0
        edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(1).SelStart = 0
        edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
        edcFDDropDown(1).SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(0).SelStart = 0
        edcFDDropDown(0).SelLength = Len(edcFDDropDown(0).Text)
        edcFDDropDown(0).SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcFDDropDown(1).SelStart = 0
        edcFDDropDown(1).SelLength = Len(edcFDDropDown(1).Text)
        edcFDDropDown(1).SetFocus
    End If
End Sub


Private Function mCheckFromDateTime()
    Dim slSDate As String
    Dim slEDate As String
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long

    mCheckFromDateTime = False
    slSDate = Trim$(edcFDDropDown(0).Text)
    If slSDate = "" Then
        MsgBox "From Start Date is Not Defined", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    If Not gValidDate(slSDate) Then
        MsgBox "From Start Date is Not Valid", vbOKOnly + vbApplicationModal, "Start Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    slEDate = Trim$(edcFDDropDown(1).Text)
    If slEDate = "" Then
        MsgBox "From End Date is Not Defined", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If Not gValidDate(slEDate) Then
        MsgBox "From End Date is Not Valid", vbOKOnly + vbApplicationModal, "End Date"
        edcFDDropDown(1).SetFocus
        Exit Function
    End If
    If gDateValue(slEDate) < gDateValue(slSDate) Then
        MsgBox "From End Date must be after or equal to Start Date", vbOKOnly + vbApplicationModal, "Date"
        edcFDDropDown(0).SetFocus
        Exit Function
    End If
    mCheckFromDateTime = True
End Function


Private Sub grdAdvtCntr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdAdvtCntr.RowHeight(0) Then
        If (grdAdvtCntr.MouseCol <= CNTRINDEX) Then
            grdAdvtCntr.Col = grdAdvtCntr.MouseCol
            mAdvtCntrSortCol grdAdvtCntr.Col
        End If
        grdAdvtCntr.Row = 0
        grdAdvtCntr.Col = CHFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdAdvtCntr, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdAdvtCntr.FixedRows Then
        Exit Sub
    End If
    mClearLineGrid
    mClearSpotGrid
    If grdAdvtCntr.TextMatrix(llCurrentRow, ADVTINDEX) = "" Then
        lmHighlightRow = -1
    Else
        lmHighlightRow = llCurrentRow
    End If
    mPaintRowColor
    If lmHighlightRow = -1 Then
        Exit Sub
    End If
    mFillLineGrid
    Screen.MousePointer = vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    gSetMousePointer grdSpots, grdSpots, vbDefault
End Sub

Private Sub grdAdvtCntr_Scroll()
    plcCalendar.Visible = False
    lmScrollTop = grdAdvtCntr.TopRow
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdAdvtCntr.ColWidth(SORTINDEX) = 0
    grdAdvtCntr.ColWidth(CHFCODEINDEX) = 0
    'grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.Width * 0.07
    grdAdvtCntr.ColWidth(CNTRINDEX) = grdAdvtCntr.Width * 0.2
    grdAdvtCntr.ColWidth(CASHTRADEINDEX) = grdAdvtCntr.Width * 0.09
    grdAdvtCntr.ColWidth(HOLDINDEX) = grdAdvtCntr.Width * 0.16
    
    grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To HOLDINDEX Step 1
        If ilCol <> ADVTINDEX Then
            grdAdvtCntr.ColWidth(ADVTINDEX) = grdAdvtCntr.ColWidth(ADVTINDEX) - grdAdvtCntr.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdAdvtCntr
    
    grdLines.ColWidth(LINESORTINDEX) = 0
    grdLines.ColWidth(VEFCODEINDEX) = 0
    grdLines.ColWidth(CLFCODEINDEX) = 0
    grdLines.ColWidth(LINEINDEX) = grdLines.Width * 0.05
    'grdLines.ColWidth(VEHICLEINDEX) = grdLines.Width * 0.07
    grdLines.ColWidth(DAYPARTINDEX) = grdLines.Width * 0.1
    grdLines.ColWidth(DATESINDEX) = grdLines.Width * 0.1
    grdLines.ColWidth(NOPERWKINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(MOINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(TUINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(WEINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(THINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(FRINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(SAINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(SUINDEX) = grdLines.Width * 0.05
    grdLines.ColWidth(RATEINDEX) = grdLines.Width * 0.1
    
    grdLines.ColWidth(VEHICLEINDEX) = grdLines.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To RATEINDEX Step 1
        If ilCol <> VEHICLEINDEX Then
            grdLines.ColWidth(VEHICLEINDEX) = grdLines.ColWidth(VEHICLEINDEX) - grdLines.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdLines

    grdSpots.ColWidth(DATAEXISTINDEX) = 0
    grdSpots.ColWidth(MOSPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(TUSPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(WESPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(THSPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(FRSPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(SASPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(SUSPOTINDEX) = grdSpots.Width * 0.14
    grdSpots.ColWidth(MOSPOTINDEX) = grdLines.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DATAEXISTINDEX Step 1
        If ilCol <> MOSPOTINDEX Then
            grdSpots.ColWidth(MOSPOTINDEX) = grdSpots.ColWidth(MOSPOTINDEX) - grdSpots.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdSpots

End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdAdvtCntr.TextMatrix(0, ADVTINDEX) = "Advertiser"
    grdAdvtCntr.TextMatrix(0, CNTRINDEX) = "Contract"
    grdAdvtCntr.TextMatrix(0, CASHTRADEINDEX) = "C/T"
    grdAdvtCntr.TextMatrix(0, HOLDINDEX) = "Status"

    grdLines.TextMatrix(0, LINEINDEX) = "Line"
    grdLines.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdLines.TextMatrix(0, DAYPARTINDEX) = "Daypart"
    grdLines.TextMatrix(0, DATESINDEX) = "Dates"
    grdLines.TextMatrix(0, NOPERWKINDEX) = "#/Wk"
    grdLines.TextMatrix(0, MOINDEX) = "Mo"
    grdLines.TextMatrix(0, TUINDEX) = "Tu"
    grdLines.TextMatrix(0, WEINDEX) = "We"
    grdLines.TextMatrix(0, THINDEX) = "Th"
    grdLines.TextMatrix(0, FRINDEX) = "Fr"
    grdLines.TextMatrix(0, SAINDEX) = "Sa"
    grdLines.TextMatrix(0, SUINDEX) = "Su"
    grdLines.TextMatrix(0, RATEINDEX) = "Spot Rate"

    grdSpots.TextMatrix(0, MOSPOTINDEX) = "Mo"
    grdSpots.TextMatrix(0, TUSPOTINDEX) = "Tu"
    grdSpots.TextMatrix(0, WESPOTINDEX) = "We"
    grdSpots.TextMatrix(0, THSPOTINDEX) = "Th"
    grdSpots.TextMatrix(0, FRSPOTINDEX) = "Fr"
    grdSpots.TextMatrix(0, SASPOTINDEX) = "Sa"
    grdSpots.TextMatrix(0, SUSPOTINDEX) = "Su"

End Sub

Private Sub mAdvtCntrSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        slStr = Trim$(grdAdvtCntr.TextMatrix(llRow, ADVTINDEX))
        If slStr <> "" Then
            If ilCol = CNTRINDEX Then
                slSort = grdAdvtCntr.TextMatrix(llRow, CNTRINDEX)
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdAdvtCntr.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdAdvtCntr.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastAdvtCntrColSorted) Or ((ilCol = imLastAdvtCntrColSorted) And (imLastAdvtCntrSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAdvtCntr.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAdvtCntr.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastAdvtCntrColSorted Then
        imLastAdvtCntrColSorted = SORTINDEX
    Else
        imLastAdvtCntrColSorted = -1
        imLastAdvtCntrSort = -1
    End If
    gGrid_SortByCol grdAdvtCntr, ADVTINDEX, SORTINDEX, imLastAdvtCntrColSorted, imLastAdvtCntrSort
    imLastAdvtCntrColSorted = ilCol
End Sub
Private Sub mLineSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    
    For llRow = grdLines.FixedRows To grdLines.Rows - 1 Step 1
        slStr = Trim$(grdLines.TextMatrix(llRow, ADVTINDEX))
        If slStr <> "" Then
            If ilCol = LINEINDEX Then
                slSort = grdLines.TextMatrix(llRow, LINEINDEX)
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = DATESINDEX Then
                slStr = grdLines.TextMatrix(llRow, DATESINDEX)
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slDate = Trim$(str(gDateValue(Trim$(Left(slStr, ilPos - 1)))))
                    Do While Len(slDate) < 6
                        slDate = "0" & slDate
                    Loop
                    slSort = slDate
                    slDate = Trim$(str(gDateValue(Trim$(Mid(slStr, ilPos + 1)))))
                    Do While Len(slDate) < 6
                        slDate = "0" & slDate
                    Loop
                    slSort = slSort & slDate
                Else
                    slSort = slStr
                    Do While Len(slSort) < 12
                        slSort = slSort & " "
                    Loop
                End If
            ElseIf ilCol = RATEINDEX Then
                slStr = grdLines.TextMatrix(llRow, RATEINDEX)
                ilPos = InStr(1, slStr, ".", vbTextCompare)
                If ilPos > 0 Then
                    slStr = Trim$(str(gStrDecToLong(slStr, 2)))
                    slSort = slStr
                    Do While Len(slSort) < 8
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = slStr
                    Do While Len(slSort) < 8
                        slSort = slSort & " "
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdLines.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdLines.TextMatrix(llRow, LINESORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastLineColSorted) Or ((ilCol = imLastLineColSorted) And (imLastLineSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdLines.TextMatrix(llRow, LINESORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdLines.TextMatrix(llRow, LINESORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastLineColSorted Then
        imLastLineColSorted = LINESORTINDEX
    Else
        imLastLineColSorted = -1
        imLastLineSort = -1
    End If
    gGrid_SortByCol grdLines, LINEINDEX, LINESORTINDEX, imLastLineColSorted, imLastLineSort
    imLastLineColSorted = ilCol
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    lmLastClickedAdvtCntrRow = -1
    imLastAdvtCntrColSorted = -1
    imLastAdvtCntrSort = -1
    grdAdvtCntr.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        grdAdvtCntr.Row = llRow
        For llCol = 0 To CHFCODEINDEX Step 1
            grdAdvtCntr.Col = llCol
            grdAdvtCntr.CellBackColor = vbWhite
            grdAdvtCntr.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdAdvtCntr.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    lmHighlightRow = -1
    mClearLineGrid
    mClearSpotGrid
End Sub
Private Sub mPaintRowColor()
    Dim llCol As Long
    Dim llRow As Long
    
    For llRow = grdAdvtCntr.FixedRows To grdAdvtCntr.Rows - 1 Step 1
        If grdAdvtCntr.TextMatrix(llRow, ADVTINDEX) <> "" Then
            For llCol = ADVTINDEX To HOLDINDEX Step 1
                grdAdvtCntr.Row = llRow
                grdAdvtCntr.Col = llCol
                If lmHighlightRow <> llRow Then
                    grdAdvtCntr.CellBackColor = vbWhite
                Else
                    grdAdvtCntr.CellBackColor = GRAY    'vbBlue
                End If
            Next llCol
        End If
    Next llRow
End Sub
Private Sub mPaintLineRowColor()
    Dim llCol As Long
    Dim llRow As Long
    
    For llRow = grdLines.FixedRows To grdLines.Rows - 1 Step 1
        If grdLines.TextMatrix(llRow, LINEINDEX) <> "" Then
            For llCol = LINEINDEX To RATEINDEX Step 1
                grdLines.Row = llRow
                grdLines.Col = llCol
                If lmLineHighlightRow <> llRow Then
                    grdLines.CellBackColor = vbWhite
                Else
                    grdLines.CellBackColor = GRAY    'vbBlue
                End If
            Next llCol
        End If
    Next llRow
End Sub


Private Function mBuildSpots(slAction As String) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilVpfIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilChf As Integer
    Dim ilVefCode As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slDays As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim llTSDate As Long
    Dim llTEDate As Long
    Dim llDate As Long
    Dim slTSTime As String
    Dim slTETime As String
    Dim llTSTime As Long
    Dim llTETime As Long
    Dim ilMaxFrontLoad As Integer
    Dim ilFLMax As Integer
    Dim slSQLQuery As String
    
    If cbcMatch.ItemData(imMatchIndex) = 9 Then
        tmChfSrchKey1.lCntrNo = lmCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmCntrNo) And (tmChf.sSchStatus <> "F")
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmCntrNo) And (tmChf.sSchStatus = "F") Then
            ReDim tmSpotFindInfo(0 To 1) As SPOTACTIONINFO
            tmSpotFindInfo(0).lChfCode = tmChf.lCode
            tmSpotFindInfo(0).iSchdSpotCount = 0
            tmSpotFindInfo(0).iMissedSpotCount = 0
            mBuildSpots = True
        Else
            ReDim tmSpotFindInfo(0 To 0) As SPOTACTIONINFO
            mBuildSpots = False
        End If
        Exit Function
    End If
    If cbcMatch.ItemData(imMatchIndex) = 10 Then
        'Find Orders and propoals
        ReDim tmSpotFindInfo(0 To 0) As SPOTACTIONINFO
        slSQLQuery = "Select distinct chfCntrNo, chfCode From chf_Contract_Header "
        slSQLQuery = slSQLQuery & " Where UCase(chfAgyEst) = '" & smEstNo & "'"
        slSQLQuery = slSQLQuery & " And chfDelete = 'N'"
        'slSQLQuery = slSQLQuery & " And chfSchStatus = 'F'"
        slSQLQuery = slSQLQuery & " Order By chfCntrNo Desc"
        Set rst_Chf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Chf.EOF
            tmSpotFindInfo(UBound(tmSpotFindInfo)).lChfCode = rst_Chf!chfCode
            tmSpotFindInfo(UBound(tmSpotFindInfo)).iSchdSpotCount = 0
            tmSpotFindInfo(UBound(tmSpotFindInfo)).iMissedSpotCount = 0
            ReDim Preserve tmSpotFindInfo(0 To UBound(tmSpotFindInfo) + 1) As SPOTACTIONINFO
            rst_Chf.MoveNext
        Loop
        mBuildSpots = True
        Exit Function
    End If
    ReDim imSelectedVefCode(0 To 0) As Integer
    slSDate = Trim$(edcFDDropDown(0).Text)
    slEDate = Trim$(edcFDDropDown(1).Text)
    slSTime = "12M"
    slETime = "12M"
    slDays = ""
    ReDim tmSpotFindInfo(0 To 0) As SPOTACTIONINFO
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            slNameCode = tmUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            imSelectedVefCode(UBound(imSelectedVefCode)) = ilVefCode
            ReDim Preserve imSelectedVefCode(0 To UBound(imSelectedVefCode) + 1) As Integer
            If bmFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(CntrFind, ilVefCode)
                bmFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(ilVefCode)
            End If
            If ilVpfIndex <> -1 Then
                btrExtClear hmSdfPsf   'Clear any previous extend operation
                ilExtLen = Len(tmSdf)
                llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
                btrExtClear hmSdfPsf   'Clear any previous extend operation
                tmSdfSrchKey1.iVefCode = ilVefCode
                gPackDate slSDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
                tmSdfSrchKey1.iTime(0) = 0
                tmSdfSrchKey1.iTime(1) = 0
                tmSdfSrchKey1.sSchStatus = ""
                ilRet = btrGetGreaterOrEqual(hmSdfPsf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    If ilRet = BTRV_ERR_NONE Then
                        Call btrExtSetBounds(hmSdfPsf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
                        tlIntTypeBuff.iType = ilVefCode
                        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                        ilRet = btrExtAddLogicConst(hmSdfPsf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                        If slSDate <> "" Then
                            gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Sdf", "SdfDate")
                            If slEDate <> "" Then
                                ilRet = btrExtAddLogicConst(hmSdfPsf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                            Else
                                ilRet = btrExtAddLogicConst(hmSdfPsf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                            End If
                        End If
                        If slEDate <> "" Then
                            gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                            ilOffSet = gFieldOffset("Sdf", "SdfDate")
                            ilRet = btrExtAddLogicConst(hmSdfPsf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        End If
                        ilRet = btrExtAddField(hmSdfPsf, 0, ilExtLen)  'Extract Name
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Function
                        End If
                        ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                                Exit Function
                            End If
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                            Loop
                            Do While ilRet = BTRV_ERR_NONE
                                If mCheckMatch() Then
                                    ilIndex = -1
                                    For ilLoop = 0 To UBound(tmSpotFindInfo) - 1 Step 1
                                        If tmSdf.lChfCode = tmSpotFindInfo(ilLoop).lChfCode Then
                                            ilIndex = ilLoop
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If ilIndex = -1 Then
                                        ilIndex = UBound(tmSpotFindInfo)
                                        tmSpotFindInfo(ilIndex).lChfCode = tmSdf.lChfCode
                                        tmSpotFindInfo(ilIndex).iSchdSpotCount = 0
                                        tmSpotFindInfo(ilIndex).iMissedSpotCount = 0
                                        ReDim Preserve tmSpotFindInfo(0 To ilIndex + 1) As SPOTACTIONINFO
                                    End If
                                    If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                        tmSpotFindInfo(ilIndex).iSchdSpotCount = tmSpotFindInfo(ilIndex).iSchdSpotCount + 1
                                    Else
                                        tmSpotFindInfo(ilIndex).iMissedSpotCount = tmSpotFindInfo(ilIndex).iMissedSpotCount + 1
                                    End If
                                End If
                                ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                                Loop
                            Loop
                        End If
                    End If
                End If
            End If
        End If
    Next ilVef
    'Check for contract where all spots moved from week spot looked into
    btrExtClear hmSmf   'Clear any previous extend operation
    ilExtLen = Len(tmSmf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSmf   'Clear any previous extend operation
    gPackDate slSDate, tmSmfSrchKey4.iMissedDate(0), tmSmfSrchKey4.iMissedDate(1)
    tmSmfSrchKey4.lChfCode = 0
    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        Call btrExtSetBounds(hmSmf, llNoRec, -1, "UC", "SMF", "") 'Set extract limits (all records)
        If slSDate <> "" Then
            gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
            If slEDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEDate <> "" Then
            gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
            ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmSmf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Function
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                For ilVef = 0 To UBound(imSelectedVefCode) - 1 Step 1
                    If imSelectedVefCode(ilVef) = tmSmf.iOrigSchVef Then
                        ilFound = True
                        Exit For
                    End If
                Next ilVef
                If ilFound Then
                    tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        If mCheckMatch() Then
                            ilIndex = -1
                            For ilLoop = 0 To UBound(tmSpotFindInfo) - 1 Step 1
                                If tmSmf.lChfCode = tmSpotFindInfo(ilLoop).lChfCode Then
                                    ilIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If ilIndex = -1 Then
                                ilIndex = UBound(tmSpotFindInfo)
                                tmSpotFindInfo(ilIndex).lChfCode = tmSmf.lChfCode
                                tmSpotFindInfo(ilIndex).iSchdSpotCount = 0
                                tmSpotFindInfo(ilIndex).iMissedSpotCount = 0
                                ReDim Preserve tmSpotFindInfo(0 To ilIndex + 1) As SPOTACTIONINFO
                            End If
                            tmSpotFindInfo(ilIndex).iSchdSpotCount = tmSpotFindInfo(ilIndex).iSchdSpotCount + 1
                        End If
                    End If
                End If
                ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If

    mBuildSpots = True
End Function

Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer

    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(Copy, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(CntrFind, cbcSearch, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", CntrFind
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAgencyPop                      *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency list box       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAgencyPop()
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
    'ilRet = gPopAgyBox(RptSel, lbcSelection, Traffic!lbcAgency)
    ilRet = gPopAgyBox(CntrFind, cbcSearch, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", CntrFind
        On Error GoTo 0
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonPop                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales office list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop()
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilType As Integer
    Dim ilDormant As Integer
    ilType = 0  'All
    ilDormant = True
    ilRet = gPopSalespersonBox(CntrFind, ilType, False, ilDormant, cbcSearch, tmSalesperson(), smSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gIMoveListBox)", Advt
        On Error GoTo 0
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mContractTypePop()
    cbcSearch.AddItem "Standard"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 1
    cbcSearch.AddItem "Remnant"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 2
    cbcSearch.AddItem "Per Inquiry"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 3
    cbcSearch.AddItem "Direct Response"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 4
    cbcSearch.AddItem "Promo"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 5
    cbcSearch.AddItem "PSA"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 6
    cbcSearch.AddItem "Reservation"
    cbcSearch.ItemData(cbcSearch.NewIndex) = 7
End Sub


Private Sub mDaypartPop()
    Dim ilRet As Integer
    Dim ilRdf As Integer
    
    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        cbcSearch.AddItem Trim$(tgMRdf(ilRdf).sName)
        cbcSearch.ItemData(cbcSearch.NewIndex) = tgMRdf(ilRdf).iCode
    Next ilRdf
End Sub

Private Sub mBuyerPop()
    Dim ilRet  As Integer
    Dim slName As String
    
    ilRet = btrGetFirst(hmPnf, tmPnf, imPnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If tmPnf.sType = "B" Then
            slName = Trim$(tmPnf.sName)
            gFindMatch slName, 0, cbcSearch
            If gLastFound(cbcSearch) < 0 Then
                cbcSearch.AddItem slName
                cbcSearch.ItemData(cbcSearch.NewIndex) = tmPnf.iCode
            End If
        End If
        ilRet = btrGetNext(hmPnf, tmPnf, imPnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub

Private Function mCheckMatch() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llPrice As Long
    Dim slPrice As String
    
    mCheckMatch = False
    tmChfSrchKey0.lCode = tmSdf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        Select Case cbcMatch.ItemData(imMatchIndex)
            Case 1  'Advertiser
                If tmChf.iAdfCode = imAdfCode Then
                    mCheckMatch = True
                End If
            Case 2  'Agency
                If tmChf.iAgfCode = imAgfCode Then
                    mCheckMatch = True
                End If
            Case 3  'Product Protection
                If (tmChf.iMnfComp(0) = imMnfComp) Or (tmChf.iMnfComp(1) = imMnfComp) Then
                    mCheckMatch = True
                End If
            Case 4  'Salesperson
                For ilLoop = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
                    If tmChf.iSlfCode(ilLoop) = imSlfCode Then
                        mCheckMatch = True
                        Exit For
                    End If
                Next ilLoop
            Case 5  'Contract Type
                If tmChf.sType = smContractType Then
                    mCheckMatch = True
                End If
            Case 6  'Buyer Name'
                If StrComp(Trim$(tmChf.sBuyer), smBuyerName, vbTextCompare) = 0 Then
                    mCheckMatch = True
                End If
            Case 7  'Rate Range
            
                tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                tmClfSrchKey0.iLine = tmSdf.iLineNo
                tmClfSrchKey0.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey0.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                    ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)  'get spot price here so that the flight can be accessed in cff
                    If InStr(1, slPrice, ".", vbTextCompare) > 0 Then
                        llPrice = gStrDecToLong(slPrice, 2)
                    Else
                        llPrice = 0
                    End If
                    If (llPrice >= lmFromRate) And (llPrice <= lmToRate) Then
                        mCheckMatch = True
                    End If
                End If
            Case 8  'Daypart
                'If imSearchIndex < 0 Then
                '    MsgBox "No Daypart Name has been selected", vbOkOnly + vbApplicationModal, "Get Advertiser"
                '    Exit Function
                'End If
                'imRdfCode = cbcSearch.ItemData(imSearchIndex)
                tmClfSrchKey0.lChfCode = tmSdf.lChfCode
                tmClfSrchKey0.iLine = tmSdf.iLineNo
                tmClfSrchKey0.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey0.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                    If (tmClf.iRdfCode = imRdfCode) Then
                        mCheckMatch = True
                    End If
                End If
            Case 9  'Contract #
                If tmChf.lCntrNo = lmCntrNo Then
                    mCheckMatch = True
                End If
            Case 10  'Estimate #
                If UCase(Trim$(tmChf.sAgyEstNo)) = smEstNo Then
                    mCheckMatch = True
                End If
        End Select
    End If
End Function

Private Sub mClearLineGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    lmLastClickedLineRow = -1
    imLastLineColSorted = -1
    imLastLineSort = -1
    grdLines.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdLines.FixedRows To grdLines.Rows - 1 Step 1
        grdLines.Row = llRow
        For llCol = 0 To CLFCODEINDEX Step 1
            grdLines.Col = llCol
            grdLines.CellBackColor = vbWhite
            grdLines.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdLines.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    lmLineHighlightRow = -1
End Sub

Private Sub mClearSpotGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    grdSpots.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdSpots.FixedRows To grdSpots.Rows - 1 Step 1
        grdSpots.Row = llRow
        For llCol = 0 To DATAEXISTINDEX Step 1
            grdSpots.Col = llCol
            grdSpots.CellBackColor = vbWhite
            grdSpots.CellForeColor = vbBlue
            grdSpots.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdSpots.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub


Private Sub mDatePop()
    Dim llDate As Long
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    lbcLineDates.Clear
    lbcSpotDates.Clear
    slStartDate = edcFDDropDown(0).Text
    llStartDate = gDateValue(gObtainPrevMonday(slStartDate))
    slEndDate = edcFDDropDown(1).Text
    llEndDate = gDateValue(gObtainNextSunday(slEndDate))
    For llDate = llStartDate To llEndDate Step 7
        lbcLineDates.AddItem Format$(llDate, "m/d/yy")
        If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
            lbcSpotDates.AddItem Format$(llDate, "m/d/yy")
            lbcSpotDates.ItemData(lbcSpotDates.NewIndex) = 0
        End If
    Next llDate
End Sub

Private Sub mFillLineGrid()
    Dim llChfCode As Long
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilShowLine As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slStrTime As String
    Dim ilVpfIndex As Integer
    Dim slGMedium As String
    
    If lmHighlightRow < grdAdvtCntr.FixedRows Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    llChfCode = grdAdvtCntr.TextMatrix(lmHighlightRow, CHFCODEINDEX)
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tmChf, tmFindClf(), tmFindCFF())
    grdLines.Redraw = False
    grdLines.Row = 0
    For llCol = LINEINDEX To RATEINDEX Step 1
        If (llCol <= DATESINDEX) Or (llCol = RATEINDEX) Then
            grdLines.Col = llCol
            grdLines.CellBackColor = LIGHTBLUE
        End If
    Next llCol
    grdLines.TopRow = 1
    grdLines.RowHeight(0) = fgBoxGridH + 15
    llRow = grdLines.FixedRows
    For ilClf = 0 To UBound(tmFindClf) - 1 Step 1
        ilShowLine = False
        If rbcShow(0).Value Then
            If (tmFindClf(ilClf).ClfRec.sType <> "O") Or (tmFindClf(ilClf).ClfRec.sType <> "A") Or (tmFindClf(ilClf).ClfRec.sType <> "E") Then
                ilShowLine = True
            End If
        Else
            If (tmFindClf(ilClf).ClfRec.sType = "O") Or (tmFindClf(ilClf).ClfRec.sType = "A") Or (tmFindClf(ilClf).ClfRec.sType = "E") Then
                ilShowLine = True
            End If
        End If
        If ilShowLine Then
            If llRow >= grdLines.Rows Then
                grdLines.AddItem ""
            End If
            grdLines.RowHeight(llRow) = fgBoxGridH + 15
            For llCol = LINEINDEX To RATEINDEX Step 1
                grdLines.Row = llRow
                grdLines.Col = llCol
                grdLines.CellBackColor = vbWhite
                grdLines.CellForeColor = vbBlue
            Next llCol
            ilVef = gBinarySearchVef(tmFindClf(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                grdLines.TextMatrix(llRow, LINEINDEX) = Trim$(str$(tmFindClf(ilClf).ClfRec.iLine))
                grdLines.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                If bmFirstCallToVpfFind Then
                    ilVpfIndex = gVpfFind(CntrFind, tmFindClf(ilClf).ClfRec.iVefCode)
                    bmFirstCallToVpfFind = False
                Else
                    ilVpfIndex = gVpfFindIndex(tmFindClf(ilClf).ClfRec.iVefCode)
                End If
                If ilVpfIndex <> -1 Then
                    slGMedium = tgVpf(ilVpfIndex).sGMedium
                Else
                    slGMedium = ""
                End If
                If ((tmFindClf(ilClf).ClfRec.iStartTime(0) <> 1) Or (tmFindClf(ilClf).ClfRec.iStartTime(1) <> 0)) And (slGMedium <> "S") Then
                    gUnpackTime tmFindClf(ilClf).ClfRec.iStartTime(0), tmFindClf(ilClf).ClfRec.iStartTime(1), "A", "1", slStartTime
                    gUnpackTime tmFindClf(ilClf).ClfRec.iEndTime(0), tmFindClf(ilClf).ClfRec.iEndTime(1), "A", "1", slEndTime
                    slStr = slStartTime & "-" & slEndTime
                Else
                    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If tmFindClf(ilClf).ClfRec.iRdfCode = tgMRdf(ilRdf).iCode Then
                            slStr = Trim$(tgMRdf(ilRdf).sName)
                            'Add times
                            slStrTime = ""
                            For ilLoop = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                If (tgMRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilLoop), tgMRdf(ilRdf).iStartTime(1, ilLoop), "A", "1", slStartTime
                                    gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilLoop), tgMRdf(ilRdf).iEndTime(1, ilLoop), "A", "1", slEndTime
                                    If slStartTime <> "" Then
                                        slStrTime = slStrTime & " " & slStartTime & "-" & slEndTime
                                    End If
                                End If
                            Next ilLoop
                            If Trim$(slStrTime) <> "" Then
                                slStr = slStr & ": " & Trim$(slStrTime)
                            End If
                            Exit For
                        End If
                    Next ilRdf
                End If
                grdLines.TextMatrix(llRow, DAYPARTINDEX) = slStr
                gUnpackDate tmFindClf(ilClf).ClfRec.iStartDate(0), tmFindClf(ilClf).ClfRec.iStartDate(1), slStartDate
                gUnpackDate tmFindClf(ilClf).ClfRec.iEndDate(0), tmFindClf(ilClf).ClfRec.iEndDate(1), slEndDate
                If gDateValue(slEndDate) >= gDateValue(slStartDate) Then
                    grdLines.TextMatrix(llRow, DATESINDEX) = slStartDate & "-" & slEndDate
                Else
                    grdLines.TextMatrix(llRow, DATESINDEX) = "CBS"
                End If
                grdLines.TextMatrix(llRow, NOPERWKINDEX) = ""
                grdLines.TextMatrix(llRow, MOINDEX) = ""
                grdLines.TextMatrix(llRow, TUINDEX) = ""
                grdLines.TextMatrix(llRow, WEINDEX) = ""
                grdLines.TextMatrix(llRow, THINDEX) = ""
                grdLines.TextMatrix(llRow, FRINDEX) = ""
                grdLines.TextMatrix(llRow, SAINDEX) = ""
                grdLines.TextMatrix(llRow, SUINDEX) = ""
                grdLines.TextMatrix(llRow, RATEINDEX) = ""
                grdLines.TextMatrix(llRow, VEFCODEINDEX) = tmFindClf(ilClf).ClfRec.iVefCode
                grdLines.TextMatrix(llRow, CLFCODEINDEX) = tmFindClf(ilClf).ClfRec.lCode
                llRow = llRow + 1
            End If
        End If
    Next ilClf
    If lbcLineDates.ListIndex < 0 Then
        If lbcLineDates.ListCount > 0 Then
            imChgMode = True
            lbcLineDates.ListIndex = 0
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
                lbcSpotDates.ListIndex = 0
            End If
            imChgMode = False
        End If
    End If
    If lbcLineDates.ListIndex >= 0 Then
        mFillFlightsInLineGrid
    End If
    mGetCntrSpots
    mLineSortCol LINEINDEX
    grdLines.Row = 0
    grdLines.Col = CLFCODEINDEX
    grdLines.Redraw = True
    gSetMousePointer grdSpots, grdSpots, vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    Screen.MousePointer = vbDefault

End Sub

Private Sub mFillFlightsInLineGrid()
    Dim ilCff As Integer
    Dim llDate As Long
    Dim llFlightStartDate As Long
    Dim llFlightEndDate As Long
    Dim llRow As Long
    Dim ilLineNo As Integer
    Dim ilDay As Integer
    Dim slRate As String
    
    If lbcLineDates.ListIndex < 0 Then
        llDate = -1
    Else
        llDate = gDateValue(lbcLineDates.List(lbcLineDates.ListIndex))
    End If
    For llRow = grdLines.FixedRows To grdLines.Rows - 1 Step 1
        If grdLines.TextMatrix(llRow, LINEINDEX) <> "" Then
            grdLines.TextMatrix(llRow, NOPERWKINDEX) = ""
            grdLines.TextMatrix(llRow, MOINDEX) = ""
            grdLines.TextMatrix(llRow, TUINDEX) = ""
            grdLines.TextMatrix(llRow, WEINDEX) = ""
            grdLines.TextMatrix(llRow, THINDEX) = ""
            grdLines.TextMatrix(llRow, FRINDEX) = ""
            grdLines.TextMatrix(llRow, SAINDEX) = ""
            grdLines.TextMatrix(llRow, SUINDEX) = ""
            grdLines.TextMatrix(llRow, RATEINDEX) = ""
            If llDate <> -1 Then
                ilLineNo = Val(grdLines.TextMatrix(llRow, LINEINDEX))
                For ilCff = LBound(tmFindCFF) To UBound(tmFindCFF) - 1 Step 1
                    If tmFindCFF(ilCff).CffRec.iClfLine = ilLineNo Then
                        gUnpackDateLong tmFindCFF(ilCff).CffRec.iStartDate(0), tmFindCFF(ilCff).CffRec.iStartDate(1), llFlightStartDate
                        gUnpackDateLong tmFindCFF(ilCff).CffRec.iEndDate(0), tmFindCFF(ilCff).CffRec.iEndDate(1), llFlightEndDate
                        If (llDate >= llFlightStartDate) And (llDate <= llFlightEndDate) Then
                            If tmFindCFF(ilCff).CffRec.sDyWk <> "D" Then
                                For ilDay = 0 To 6 Step 1
                                    If llFlightStartDate + 6 <= llFlightEndDate Then
                                        If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                            grdLines.TextMatrix(llRow, MOINDEX + ilDay) = "Y"
                                        End If
                                    Else
                                        If llFlightStartDate + ilDay <= llFlightEndDate Then
                                            If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                grdLines.TextMatrix(llRow, MOINDEX + ilDay) = "Y"
                                            End If
                                        End If
                                    End If
                                Next ilDay
                                grdLines.TextMatrix(llRow, NOPERWKINDEX) = Trim$(str(tmFindCFF(ilCff).CffRec.iSpotsWk))
                            Else
                                For ilDay = 0 To 6 Step 1
                                    If llFlightStartDate + 6 <= llFlightEndDate Then
                                        If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                            grdLines.TextMatrix(llRow, MOINDEX + ilDay) = Trim$(str(tmFindCFF(ilCff).CffRec.iDay(ilDay)))
                                        End If
                                    Else
                                        If llFlightStartDate + ilDay <= llFlightEndDate Then
                                            If tmFindCFF(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                grdLines.TextMatrix(llRow, MOINDEX + ilDay) = Trim$(str(tmFindCFF(ilCff).CffRec.iDay(ilDay)))
                                            End If
                                        End If
                                    End If
                                Next ilDay
                            End If
                            Select Case tmFindCFF(ilCff).CffRec.sPriceType
                                Case "T"    'True
                                    slRate = gLongToStrDec(tmFindCFF(ilCff).CffRec.lActPrice, 2)
                                Case "N"    'No Charge
                                    slRate = "N/C"
                                Case "M"    'MG Line
                                    slRate = "MG"
                                Case "B"    'Bonus
                                    slRate = "Bonus"
                                Case "S"    'Spinoff
                                    slRate = "Spinoff"
                                Case "P"    'Package
                                    slRate = gLongToStrDec(tmFindCFF(ilCff).CffRec.lActPrice, 2)
                                Case "R"    'Recapturable
                                    slRate = "Recapturable"
                                Case "A"    'ADU
                                    slRate = "ADU"
                            End Select
                            grdLines.TextMatrix(llRow, RATEINDEX) = slRate
                        End If
                    End If
                Next ilCff
            End If
        End If
    Next llRow
End Sub

Private Sub mGetCntrSpots()
    Dim llChfCode As Long
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slSortDate As String
    Dim slSortTime As String
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim slLineNo As String
    Dim llTime As Long
    Dim ilIncludeSpot As Integer
    Dim tlLongTypeBuff As POPLCODE          '6-12-02 long
    
    ReDim tmCntrFindSpot(0 To 0) As CNTRFINDSPOT
    slStartDate = edcFDDropDown(0).Text
    llStartDate = gDateValue(gObtainPrevMonday(slStartDate))
    slEndDate = edcFDDropDown(1).Text
    llEndDate = gDateValue(gObtainNextSunday(slEndDate))
    
    llChfCode = grdAdvtCntr.TextMatrix(lmHighlightRow, CHFCODEINDEX)
    
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    
    tmSdfSrchKey5.lCode = llChfCode
    ilRet = btrGetEqual(hmSdfPsf, tmSdf, imSdfRecLen, tmSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdfPsf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        tlLongTypeBuff.lCode = llChfCode
        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hmSdfPsf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)
        ilRet = btrExtAddField(hmSdfPsf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilIncludeSpot = True
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (tmSdf.sSpotType <> "X") Then
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llDate
                        If (llDate < llStartDate) Or (llDate > llEndDate) Then
                            ilIncludeSpot = False
                        Else
                            gUnpackTimeLong tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), False, llTime
                        End If
                    Else
                        ilIncludeSpot = False
                    End If
                Else
                    If (llDate < llStartDate) Or (llDate > llEndDate) Then
                        ilIncludeSpot = False
                    Else
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                    End If
                End If
                If ilIncludeSpot Then
                    ilUpper = UBound(tmCntrFindSpot)
                    slSortDate = Trim$(str$(llDate))
                    Do While Len(slSortDate) < 6
                        slSortDate = "0" & slSortDate
                    Loop
                    slSortTime = Trim$(str$(llTime))
                    Do While Len(slSortTime) < 5
                        slSortTime = "0" & slSortTime
                    Loop
                    slLineNo = Trim$(str(tmSdf.iLineNo))
                    Do While Len(slLineNo) < 5
                        slLineNo = "0" & slLineNo
                    Loop
                    tmCntrFindSpot(ilUpper).sKey = slLineNo & "|" & slSortDate & "|" & slSortTime
                    tmCntrFindSpot(ilUpper).iLineNo = tmSdf.iLineNo
                    tmCntrFindSpot(ilUpper).iGameNo = tmSdf.iGameNo
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), tmCntrFindSpot(ilUpper).sAirDate
                    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", tmCntrFindSpot(ilUpper).sAirTime
                    If tmSdf.sSpotType = "X" Then
                        tmCntrFindSpot(ilUpper).sStatus = "X"
                    Else
                        tmCntrFindSpot(ilUpper).sStatus = tmSdf.sSchStatus
                    End If
                    If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (tmSdf.sSpotType <> "X") Then
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), tmCntrFindSpot(ilUpper).sMissedDate
                        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", tmCntrFindSpot(ilUpper).sMissedTime
                        tmCntrFindSpot(ilUpper).iAirVefCode = tmSdf.iVefCode
                        tmCntrFindSpot(ilUpper).iMissedGameNo = tmSmf.iGameNo
                    End If
                    tmCntrFindSpot(ilUpper).sPtType = tmSdf.sPtType
                    tmCntrFindSpot(ilUpper).lCopyCode = tmSdf.lCopyCode
                    ReDim Preserve tmCntrFindSpot(0 To ilUpper + 1) As CNTRFINDSPOT
                End If
                ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdfPsf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    If UBound(tmCntrFindSpot) - 1 > 0 Then
        ArraySortTyp fnAV(tmCntrFindSpot(), 0), UBound(tmCntrFindSpot), 0, LenB(tmCntrFindSpot(0)), 0, LenB(tmCntrFindSpot(0).sKey), 0
    End If
End Sub

Private Sub mFillSpotGrid()
    Dim llRow As Integer
    Dim llCol As Long
    Dim ilLineNo As Integer
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilPass As Long
    Dim llTitleRow As Long
    Dim llSpotRow As Long
    Dim ilSdf As Integer
    Dim ilDay As Integer
    Dim ilVef As Integer
    Dim ilIncludeSpot As Integer
    Dim ilSpotCount As Integer
    Dim slISCI As String
    Dim ilGameNo As Integer
    ReDim ilTitleShown(0 To 5) As Integer
    
    If lmLineHighlightRow < grdLines.FixedRows Then
        Exit Sub
    End If
    If lbcSpotDates.ListIndex = -1 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdAdvtCntr, grdLines, vbHourglass
    gSetMousePointer grdSpots, grdSpots, vbHourglass
    ilLineNo = grdLines.TextMatrix(lmLineHighlightRow, LINEINDEX)
    If lbcSpotDates.ItemData(lbcSpotDates.ListIndex) = 0 Then
        llSDate = gDateValue(lbcSpotDates.List(lbcSpotDates.ListIndex))
        llEDate = llSDate + 6
        ilGameNo = 0
    Else
        ilGameNo = tmGsfInfo(lbcSpotDates.ItemData(lbcSpotDates.ListIndex) - 1).iGameNo
        llSDate = 0
        llEDate = llSDate + 6
    End If
    grdSpots.Redraw = False
    mClearSpotGrid
    grdSpots.RowHeight(0) = fgBoxGridH + 15
    llTitleRow = grdSpots.FixedRows - 1
    For ilPass = 1 To 6 Step 1
        ilTitleShown(ilPass - 1) = False
        ilDay = 0
        ilSpotCount = 0
        For llDate = llSDate To llEDate Step 1
            If ilGameNo <> 0 Then
                If lbcSpotDates.ListIndex + llDate < lbcSpotDates.ListCount Then
                    ilGameNo = tmGsfInfo(lbcSpotDates.ItemData(lbcSpotDates.ListIndex + llDate) - 1).iGameNo
                Else
                    Exit For
                End If
            End If
            For ilSdf = 0 To UBound(tmCntrFindSpot) - 1 Step 1
                If (tmCntrFindSpot(ilSdf).iLineNo = ilLineNo) Then
                    ilIncludeSpot = False
                    Select Case ilPass
                        Case 1
                            If tmCntrFindSpot(ilSdf).sStatus = "S" Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                        Case 2
                            If (tmCntrFindSpot(ilSdf).sStatus = "G") Or (tmCntrFindSpot(ilSdf).sStatus = "O") Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iMissedGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sMissedDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                        Case 3
                            If tmCntrFindSpot(ilSdf).sStatus = "M" Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                        Case 4
                            If tmCntrFindSpot(ilSdf).sStatus = "X" Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                        Case 5
                            If tmCntrFindSpot(ilSdf).sStatus = "C" Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                        Case 6
                            If tmCntrFindSpot(ilSdf).sStatus = "H" Then
                                If ilGameNo <> 0 Then
                                    If (tmCntrFindSpot(ilSdf).iGameNo = ilGameNo) Then
                                        ilIncludeSpot = True
                                    End If
                                Else
                                    If gDateValue(tmCntrFindSpot(ilSdf).sAirDate) = llDate Then
                                        ilIncludeSpot = True
                                    End If
                                End If
                            End If
                    End Select
                    If ilIncludeSpot Then
                        ilSpotCount = ilSpotCount + 1
                        If ilTitleShown(ilPass - 1) = False Then
                            'Find first empty row to place header
                            llTitleRow = -1
                            For llRow = grdSpots.FixedRows To grdSpots.Rows - 1 Step 1
                                If grdSpots.TextMatrix(llRow, DATAEXISTINDEX) = "" Then
                                    llTitleRow = llRow
                                    Exit For
                                End If
                            Next llRow
                            If llTitleRow = -1 Then
                                llTitleRow = grdSpots.Rows
                                If llTitleRow >= grdSpots.Rows Then
                                    grdSpots.AddItem ""
                                End If
                            End If
                            grdSpots.RowHeight(llTitleRow) = fgBoxGridH + 15
                            For llCol = MOSPOTINDEX To SUSPOTINDEX Step 1
                                grdSpots.Row = llTitleRow
                                grdSpots.Col = llCol
                                If llCol <> THSPOTINDEX Then
                                    grdSpots.CellBackColor = vbWhite
                                    grdSpots.CellForeColor = vbBlue
                                Else
                                    grdSpots.CellBackColor = LIGHTYELLOW
                                    grdSpots.CellForeColor = vbBlack
                                End If
                            Next llCol
                            Select Case ilPass
                                Case 1
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "Scheduled"
                                Case 2
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "MG/Outside"
                                Case 3
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "Missed"
                                Case 4
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "Fill"
                                Case 5
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "Cancelled"
                                Case 6
                                    grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = "Hidden"
                            End Select
                            grdSpots.TextMatrix(llTitleRow, DATAEXISTINDEX) = "Y"
                            ilTitleShown(ilPass - 1) = True
                        End If
                        llSpotRow = -1
                        For llRow = llTitleRow + 1 To grdSpots.Rows - 1 Step 1
                            If grdSpots.TextMatrix(llRow, MOSPOTINDEX + ilDay) = "" Then
                                llSpotRow = llRow
                                Exit For
                            End If
                        Next llRow
                        If llSpotRow = -1 Then
                            llSpotRow = grdSpots.Rows
                            If llSpotRow >= grdSpots.Rows Then
                                grdSpots.AddItem ""
                            End If
                        End If
                        grdSpots.RowHeight(llSpotRow) = fgBoxGridH + 15
                        For llCol = MOSPOTINDEX To SUSPOTINDEX Step 1
                            grdSpots.Row = llSpotRow
                            grdSpots.Col = llCol
                            grdSpots.CellBackColor = vbWhite
                            grdSpots.CellForeColor = vbBlue
                        Next llCol
                        If ckcShowISCI.Value = vbChecked Then
                            If tmCntrFindSpot(ilSdf).sPtType <> "I" Then
                                tmCntrFindSpot(ilSdf).sISCI = mObtainISCI(tmCntrFindSpot(ilSdf).sPtType, tmCntrFindSpot(ilSdf).lCopyCode)
                                tmCntrFindSpot(ilSdf).sPtType = "I"
                            End If
                            slISCI = Trim$(tmCntrFindSpot(ilSdf).sISCI)
                            If slISCI <> "" Then
                                slISCI = " " & slISCI
                            End If
                        Else
                            slISCI = ""
                        End If
                        Select Case ilPass
                            Case 1
                                grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = tmCntrFindSpot(ilSdf).sAirTime & slISCI
                            Case 2
                                ilVef = gBinarySearchVef(tmCntrFindSpot(ilSdf).iAirVefCode)
                                If (ilVef <> -1) Then
                                    If ilGameNo > 0 Then
                                        grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = "#" & Trim$(str(tmCntrFindSpot(ilSdf).iGameNo)) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirDate) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirTime) & " " & Trim$(tgMVef(ilVef).sName) & slISCI
                                    Else
                                        grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = Trim$(tmCntrFindSpot(ilSdf).sAirDate) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirTime) & " " & Trim$(tgMVef(ilVef).sName) & slISCI
                                    End If
                                Else
                                    If ilGameNo > 0 Then
                                        grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = "#" & Trim$(str(tmCntrFindSpot(ilSdf).iGameNo)) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirDate) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirTime) & slISCI
                                    Else
                                        grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = Trim$(tmCntrFindSpot(ilSdf).sAirDate) & " " & Trim$(tmCntrFindSpot(ilSdf).sAirTime) & slISCI
                                    End If
                                End If
                            Case 3
                                grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = tmCntrFindSpot(ilSdf).sAirTime & slISCI
                            Case 4
                                grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = tmCntrFindSpot(ilSdf).sAirTime & slISCI
                            Case 5
                                grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = tmCntrFindSpot(ilSdf).sAirTime
                            Case 6
                                grdSpots.TextMatrix(llSpotRow, MOSPOTINDEX + ilDay) = tmCntrFindSpot(ilSdf).sAirTime
                        End Select
                        grdSpots.TextMatrix(llSpotRow, DATAEXISTINDEX) = "Y"
                    End If
                End If
            Next ilSdf
            ilDay = ilDay + 1
        Next llDate
        If ilSpotCount > 0 Then
            grdSpots.TextMatrix(llTitleRow, THSPOTINDEX) = Trim$(str$(ilSpotCount)) & " " & grdSpots.TextMatrix(llTitleRow, THSPOTINDEX)
        End If
    Next ilPass
    cmcCancel.Caption = "&Done"
    grdSpots.Redraw = True
    gSetMousePointer grdSpots, grdSpots, vbDefault
    gSetMousePointer grdAdvtCntr, grdLines, vbDefault
    Screen.MousePointer = vbDefault
End Sub

Private Sub pbcPrinting_Paint()
    pbcPrinting.CurrentX = (pbcPrinting.Width - pbcPrinting.TextWidth("Printing Log Check Information....")) / 2
    pbcPrinting.CurrentY = (pbcPrinting.Height - pbcPrinting.TextHeight("Printing Log Check Information....")) / 2 - 30
    pbcPrinting.Print "Printing Contract Information...."
End Sub

Private Sub rbcShow_Click(Index As Integer)
    If Index = 0 Then
        hmSdfPsf = hmSdf
    Else
        hmSdfPsf = hmPsf
    End If
    mClearGrid
    imSetAll = False
    ckcVehicleNonSports.Value = vbUnchecked
    ckcVehicleSports.Value = vbUnchecked
    imSetAll = True
    mVehPop
    If Index = 0 Then
        If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) Then
            ckcVehicleSports.Visible = True
            ckcVehicleNonSports.Caption = "All Non-Sports"
        Else
            ckcVehicleSports.Visible = False
            ckcVehicleNonSports.Caption = "All Vehicles"
        End If
    Else
        ckcVehicleSports.Visible = False
        ckcVehicleNonSports.Caption = "All Vehicles"
    End If
    mSetCommands
End Sub

Private Sub rbcShow_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub

Private Function mObtainISCI(slPtType As String, llCopyCode As Long) As String
    Dim ilRet As Integer
    
    mObtainISCI = ""
    If slPtType = "1" Then                 'copy instr
        tmCifSrchKey0.lCode = llCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tmCpfSrchKey0.lCode = tmCif.lcpfCode     'product/isci
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) Then
                mObtainISCI = Trim$(tmCpf.sISCI)
            End If
        End If
    End If
End Function

Private Sub mSavePrtIndex(ilSdf As Integer, ilPrt() As Integer, ilMaxSpots)
    ilPrt(UBound(ilPrt)) = ilSdf
    ReDim Preserve ilPrt(0 To UBound(ilPrt) + 1) As Integer
    If UBound(ilPrt) > ilMaxSpots Then
        ilMaxSpots = UBound(ilPrt)
    End If
End Sub

Private Sub mSetPrtIndex(ilPrt() As Integer, ilMaxSpots)
    Dim ilLoop As Integer
    Dim ilCurrentUBound As Integer
    
    ilCurrentUBound = UBound(ilPrt)
    ReDim Preserve ilPrt(0 To ilMaxSpots) As Integer
    For ilLoop = ilCurrentUBound To ilMaxSpots Step 1
        ilPrt(ilLoop) = -1
    Next ilLoop
End Sub

Private Sub mDateSpotPop()
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llTime As Long
    Dim slSortDate As String
    Dim slSortTime As String
    Dim slSortGameNo As String
    Dim ilUpper As Integer
    Dim ilTeam As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    If lmLineHighlightRow < grdLines.FixedRows Then
        Exit Sub
    End If
    ilVefCode = Val(grdLines.TextMatrix(lmLineHighlightRow, VEFCODEINDEX))
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType <> "G" Then
            If lbcSpotDates.ListCount > 0 Then
                If lbcSpotDates.ItemData(0) = 0 Then
                    Exit Sub
                End If
            End If
            lbcSpotDates.Clear
            For ilLoop = 0 To 6 Step 1
                grdSpots.TextMatrix(0, ilLoop) = ""
            Next ilLoop
            For ilLoop = 0 To lbcLineDates.ListCount - 1 Step 1
                lbcSpotDates.AddItem lbcLineDates.List(ilLoop)
            Next ilLoop
            lbcSpotDates.ListIndex = lbcLineDates.ListIndex
        Else
            ReDim Preserve tmGsfInfo(0 To 0) As GSFINFO
            slStartDate = edcFDDropDown(0).Text
            llStartDate = gDateValue(gObtainPrevMonday(slStartDate))
            slEndDate = edcFDDropDown(1).Text
            llEndDate = gDateValue(gObtainNextSunday(slEndDate))
            tmGsfSrchKey3.iVefCode = ilVefCode
            tmGsfSrchKey3.iGameNo = 0
            ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode)
                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                    ilUpper = UBound(tmGsfInfo)
                    slSortDate = Trim$(str$(llDate))
                    Do While Len(slSortDate) < 5
                        slSortDate = "0" & slSortDate
                    Loop
                    gUnpackTimeLong tmGsf.iAirTime(0), tmGsf.iAirTime(1), False, llTime
                    slSortTime = Trim$(str(llTime))
                    Do While Len(slSortTime) < 5
                        slSortTime = "0" & slSortTime
                    Loop
                    slSortGameNo = Trim$(str(tmGsf.iGameNo))
                    tmGsfInfo(ilUpper).sKey = slSortDate & slSortTime & slSortGameNo
                    tmGsfInfo(ilUpper).lGameDate = llDate
                    tmGsfInfo(ilUpper).iGameNo = tmGsf.iGameNo
                    'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                    For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                        If tmGsf.iVisitMnfCode = tmTeam(ilTeam).iCode Then
                            tmGsfInfo(ilUpper).sVisitName = tmTeam(ilTeam).sName
                            If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                                tmGsfInfo(ilUpper).sVisitName = tmTeam(ilTeam).sUnitType
                            End If
                            Exit For
                        End If
                    Next ilTeam
                    'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                    For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                        If tmGsf.iHomeMnfCode = tmTeam(ilTeam).iCode Then
                            tmGsfInfo(ilUpper).sHomeName = tmTeam(ilTeam).sName
                            If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                                tmGsfInfo(ilUpper).sHomeName = tmTeam(ilTeam).sUnitType
                            End If
                            Exit For
                        End If
                    Next ilTeam
                    tmGsfInfo(ilUpper).sGameStatus = tmGsf.sGameStatus
                    tmGsfInfo(ilUpper).sAffPost = ""
                    ReDim Preserve tmGsfInfo(0 To ilUpper + 1) As GSFINFO
                End If
                ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If UBound(tmGsfInfo) - 1 > 0 Then
                ArraySortTyp fnAV(tmGsfInfo(), 0), UBound(tmGsfInfo), 0, LenB(tmGsfInfo(0)), 0, LenB(tmGsfInfo(0).sKey), 0
            End If
            If lbcLineDates.ListIndex >= 0 Then
                llDate = gDateValue(lbcLineDates.List(lbcLineDates.ListIndex))
            Else
                llDate = -1
            End If
            ilIndex = -1
            lbcSpotDates.Clear
            For ilLoop = 0 To 6 Step 1
                grdSpots.TextMatrix(0, ilLoop) = ""
            Next ilLoop
            For ilLoop = 0 To UBound(tmGsfInfo) - 1 Step 1
                lbcSpotDates.AddItem tmGsfInfo(ilLoop).iGameNo & " " & Format(tmGsfInfo(ilLoop).lGameDate, "m/d/yy")
                lbcSpotDates.ItemData(lbcSpotDates.NewIndex) = ilLoop + 1
                If ilIndex = -1 Then
                    If tmGsfInfo(ilLoop).lGameDate = llDate Then
                        ilIndex = ilLoop
                    End If
                End If
            Next ilLoop
            lbcSpotDates.ListIndex = ilIndex
        End If
    End If
End Sub

Private Sub mSetCommands()
    Dim ilIndex As Integer
    If (imMatchIndex = -1) Or (lbcVehicles.SelCount <= 0) Then
        cmcAdvtCntr.Enabled = False
    Else
        ilIndex = cbcMatch.ItemData(imMatchIndex)
        If (ilIndex <> 7) And (ilIndex <> 9) And (ilIndex <> 10) Then
            If imSearchIndex = -1 Then
                cmcAdvtCntr.Enabled = False
            Else
                cmcAdvtCntr.Enabled = True
            End If
        ElseIf (ilIndex = 7) Then
            If (Trim$(edcFrom.Text) = "") Or (Trim$(edcTo.Text) = "") Then
                cmcAdvtCntr.Enabled = False
            Else
                cmcAdvtCntr.Enabled = True
            End If
        Else
            If (Trim$(edcSearch.Text) = "") Then
                cmcAdvtCntr.Enabled = False
            Else
                cmcAdvtCntr.Enabled = True
            End If
        End If
    End If
    If (lmLineHighlightRow < grdLines.FixedRows) Or (lbcSpotDates.ListIndex = -1) Then
        cmcRefresh.Enabled = False
        imcPrt.Visible = False
    Else
        cmcRefresh.Enabled = True
        imcPrt.Visible = True
        imcPrt.Enabled = True
    End If
    cmcCancel.Caption = "&Cancel"
End Sub


Private Sub mHeading1(ilRet As Integer, slHeading As String, ilCurrentLineNo As Integer)
    On Error GoTo mHeading1Err:
    Printer.Print slHeading
    If ilRet <> 0 Then
        'Return
        Exit Sub
    End If
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
mHeading1Err:
    ilRet = Err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub

Private Sub mLineOutput(ilRet As Integer, slHeading As String, ilCurrentLineNo As Integer, slRecord As String, ilLinesPerPage As Integer)
    On Error GoTo mLineOutputErr:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
        ilCurrentLineNo = 0
        '6/6/16: Replaced GoSub
        'GoSub mHeading1
        mHeading1 ilRet, slHeading, ilCurrentLineNo
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
mLineOutputErr:
    ilRet = Err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub


Private Sub mPartialLineOutput(ilRet As Integer, slHeading As String, ilCurrentLineNo As Integer, slRecord As String, ilLinesPerPage As Integer, ilDay As Integer, ilCurrentX() As Integer, ilMaxPrtWidth As Integer)
    On Error GoTo mPartialLineOutputErr:
    If (ilCurrentLineNo >= ilLinesPerPage) And (ilDay = 0) Then
        Printer.NewPage
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
        ilCurrentLineNo = 0
        '6/6/16: Replaced GoSub
        'GoSub mHeading1
        mHeading1 ilRet, slHeading, ilCurrentLineNo
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
    End If
    If ilCurrentX(ilDay) = -1 Then
        If ilDay <> 6 Then
            Printer.Print slRecord; Spc(5),  'Tab,
        Else
            Printer.Print slRecord
            ilCurrentLineNo = ilCurrentLineNo + 1
        End If
    Else
        Do While Printer.TextWidth(slRecord) > ilMaxPrtWidth
            slRecord = Left(slRecord, Len(slRecord) - 1)
        Loop
        Printer.CurrentX = ilCurrentX(ilDay)
        If ilDay <> 6 Then
            Printer.Print slRecord,
        Else
            Printer.Print slRecord
        End If
    End If
    Exit Sub
mPartialLineOutputErr:
    ilRet = Err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub
