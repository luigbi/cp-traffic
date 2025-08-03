VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpEfficioRev 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   10020
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
   ScaleHeight     =   4125
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   780
      Top             =   3435
   End
   Begin VB.Frame frcMonthBy 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Calendar"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   28
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton rbcMonthBy 
         Caption         =   "Standard"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.ListBox lbcVehicle 
      Height          =   645
      ItemData        =   "ExpEfficioRev.frx":0000
      Left            =   6120
      List            =   "ExpEfficioRev.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      Height          =   195
      Left            =   6120
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PlcNetBy 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   3135
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3135
      Begin VB.OptionButton rbcNetBy 
         Caption         =   "Net"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   23
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton rbcNetBy 
         Caption         =   "T-Net"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   22
         Top             =   0
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   15
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   3120
   End
   Begin VB.CheckBox ckcNTR 
      Caption         =   "Include NTR Revenue"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox edcNoMonths 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox edcYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1080
      Width           =   840
   End
   Begin VB.TextBox edcMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   960
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1080
      Width           =   600
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   9000
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   2595
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2595
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3030
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7995
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3030
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8265
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcTo 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   7920
      TabIndex        =   18
      Top             =   2280
      Width           =   1485
   End
   Begin VB.PictureBox plcTo 
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   6285
      TabIndex        =   2
      Top             =   2280
      Width           =   6345
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   6225
      End
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
      Left            =   3600
      TabIndex        =   19
      Top             =   3720
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
      Left            =   5160
      TabIndex        =   20
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lacContract 
      Appearance      =   0  'Flat
      Caption         =   "Contract #"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label lacNoMonths 
      Appearance      =   0  'Flat
      Caption         =   "# months"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3600
      TabIndex        =   8
      Top             =   1140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label lacMonth 
      Appearance      =   0  'Flat
      Caption         =   "Month"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label lacStartYear 
      Appearance      =   0  'Flat
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1800
      TabIndex        =   6
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label lacSaveIn 
      Appearance      =   0  'Flat
      Caption         =   "Save In"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   810
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "ExpEfficioRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
'       Efficio Export for Revenue and Projections
' File Name: ExpEfficioRev.Frm- Export 1 month of receivables/history (user selectable or default most current billed month.
'            Export up to 36 months of projections from contracts.
'            Triple Net amounts are calculated, including Air Time/Ntr/Merchandising/Promotions & Acquistion costs
'
'
' Release: 7.0
'
'
Option Explicit
Option Compare Text
Private Const MAXMONTHS = 36
Dim hmMsg As Integer
Dim hmEfficio As Integer
Dim smExportCaption As String
Dim bmStdExport As Boolean
Dim smExportMesg As String

Dim smExportName As String
Dim imFirstActivate As Integer
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
Dim smYear As String    'Default year
Dim smMonth As String   'default month
Dim imMonth As Integer  'default month
Dim imYear As Integer   'default year
Dim lmLastBilled As Long        'last billed date from site
Dim smLastBilled As String

Dim imFirstTime As Integer
Dim tmChfAdvtExt() As CHFADVTEXT

Dim lmProject(0 To MAXMONTHS) As Long          'projection $, max 2 years
'2-28-14 implement tnet option
Dim lmAcquisition(0 To MAXMONTHS) As Long      'Acquisition $, max 2 years

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF


Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim tmAgf As AGF

Dim tmRvf As RVF

Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof() As SOF

Dim hmSlf As Integer            'Salesperson file handle

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tmMnfSS() As MNF                    'array of Sales Sources MNF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer

Dim hmVff As Integer            'Vfhicle features file handle
Dim tmVff As VFF                'VfF record image
Dim imVffRecLen As Integer        'VfF record length
Dim tmSrchVffKey As INTKEY0

'spots for calendar export
Dim hmSdf As Integer        'SDF file handle
Dim tmSdf As SDF            'SDF record image
Dim imSdfRecLen As Integer  'SDF record length

Dim hmSmf As Integer

Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT

Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim imSbfRecLen As Integer  'SBF record length

'Product
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long

Dim lmSlfSplit() As Long           '9-30-02 make 4 tables common: slsp slsp share %
Dim imSlfCode() As Integer
Dim imslfcomm() As Integer             'slsp under comm %, unused in this rept but reqd for common subroutine
Dim imslfremnant() As Integer          'slsp under remnant %, unused in this rept but reqd for common subroutine
Dim lmSlfSplitRev() As Long           '1-31-04 Rev % split for B & B Sales comm. if 0% and Slsp comm % in chf is non zero, calc
Dim lmTempGross As Long
Dim lmTempNet As Long
Dim lmTempPct As Long
Dim lmTempAcquisition As Long       '2-28-14
Dim smClientName As String
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcVehicle)
Dim imIncludeCodes As Integer
Dim imUseCodes() As Integer
''
''                      Calculate Gross & Net $, and Split Cash/Trade $ from a schedule line
''                      mCalcMonthAmt - Loop and calculate the gross and net values for up to 36 months
''
''                       <input> llTempGross - 36 months of projected $ (from contract line)
''                               ilLastBilledInx - index to last month invoiced.
''                               ilCorT - 1 = Cash , 2 = Trade processing
''                               slPctTrade  - % of trade
''                               slCashAgyComm - agency comm %
''                       <output> llTempGross - altered if split cash/trade calculation
''                               llTempNet - 36months projected net $
''
'Private Sub mCalcMonthAmt(llTempGross() As Long, llTempNet() As Long, llTempAcquisition() As Long, ilLastBilledInx As Integer, ilCorT As Integer, slPctTrade As String, slCashAgyComm As String)
'Dim ilTemp As Integer
'Dim slAmount As String
'Dim slSharePct As String
'Dim slStr As String
'Dim slCode As String
'Dim slDollar As String
'Dim slNet As String
'Dim slAcquisition As String
'Dim slAcqAmount As String
'Dim slAcqShare As String
'
'    For ilTemp = ilLastBilledInx To igPeriods              'loop on # buckets to process.
'        slAmount = gLongToStrDec(llTempGross(ilTemp), 2)
'        slSharePct = gLongToStrDec(10000, 2)
'        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'        slStr = gRoundStr(slStr, "1", 0)
'        'calc the acquisition share for split cash/trade
'        slAcqAmount = gLongToStrDec(llTempAcquisition(ilTemp), 2)
'        slAcqShare = gMulStr(slSharePct, slAcqAmount)
'        slAcqShare = gRoundStr(slAcqShare, "1", 0)
'        If ilCorT = 1 Then                 'all cash commissionable
'            slCode = gSubStr("100.", slPctTrade)                'get the cash % (100-trade%)
'            slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
'            slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
'            slAcquisition = gDivStr(gMulStr(slAcqShare, slCode), "100")              'Acquisition is always net (same as gross)
'        Else
'            If ilCorT = 2 Then                'at least cash is commissionable
'                slCode = gIntToStrDec(tgChfCT.iPctTrade, 0)
'                slDollar = gDivStr(gMulStr(slStr, slCode), "100")
'                slAcquisition = gDivStr(gMulStr(slAcqShare, slCode), "100")
'
'                If tgChfCT.iAgfCode > 0 And tgChfCT.sAgyCTrade = "Y" Then
'                    slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), "1", 0)
'                Else
'                    slNet = slDollar    'no commission , net is same as gross
'                End If
'            End If
'        End If
'        llTempGross(ilTemp) = Val(slDollar)
'        llTempNet(ilTemp) = Val(slNet)
'        llTempAcquisition(ilTemp) = Val(slAcquisition)
'    Next ilTemp
'    Exit Sub
'
'End Sub

'
'
'
'           mCloseEfficiofiles - Close all applicable files
'
Sub mCloseEfficioFiles()
Dim ilRet As Integer
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmVff)

    btrDestroy hmAgf
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSbf
    btrDestroy hmPrf
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmVff
    
End Sub

'
'
'           mOpenEfficioFiles - open files applicable to Efficio Export
'                           The Export takes all Receivables/History for 1 month, defaulting to the
'                           current month billed
'
'
Function mOpenEfficioFiles() As Integer
Dim ilRet As Integer
Dim ilTemp As Integer
Dim ilError As Integer
Dim slStamp As String
Dim slRegionStamp As String
Dim tlSof As SOF

    ilError = False

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen AGF)", ExpEfficioRev
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)

    hmSbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen SBF)", ExpEfficioRev
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen CHF)", ExpEfficioRev
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen VEF)", ExpEfficioRev
    On Error GoTo 0
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen MNF)", ExpEfficioRev
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen SOF)", ExpEfficioRev
    On Error GoTo 0
    imSofRecLen = Len(tlSof)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen CLF)", ExpEfficioRev
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen CFF)", ExpEfficioRev
    On Error GoTo 0
    imCffRecLen = Len(tmCff)

    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen PRF)", ExpEfficioRev
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen SDF)", ExpEfficioRev
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenEfficioFilesErr
    gBtrvErrorMsg ilRet, "gOpenEfficioFiles (btrOpen SMF)", ExpEfficioRev
    On Error GoTo 0
    
    'build array of sales offices
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tlSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSof(0 To ilTemp) As SOF
        tmSof(ilTemp) = tlSof
        ilRet = btrGetNext(hmSof, tlSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    ilRet = gObtainMnfForType("S", slStamp, tmMnfSS())        'build sales sources
    
    imFirstTime = True              'set to create the header record in text file only once
    mOpenEfficioFiles = ilError
    Exit Function

mOpenEfficioFilesErr:
    ilError = True
    Return
End Function

'
'
'           mWriteExportRec - gather all the information for a month and write
'           a record to the export .csv file
'
'           <input> tlEfficioInfo - structure containing all the info required to write up month of data from
'                                   either the receivables
'           Return - true if error, otherwise false'
Private Function mWriteExportRec(tlEfficioINfo As MATRIXINFO) As Integer
Dim ilLoop As Integer
Dim slStr As String
Dim ilIndex As Integer
Dim ilOfficeInx As Integer
Dim ilSSInx As Integer
Dim slVehicle As String
Dim slSS As String
Dim slOffice As String
Dim slSlsp As String
Dim slAdvt As String
Dim slAgency As String
Dim ilError As Integer
Dim slStripCents As String
Dim ilRemainder As Integer
Dim slPrimComp As String
Dim slSecComp As String
Dim ilRet As Integer
Dim llTNet As Long
Dim llAdjPromoMerch As Long
Dim slOrigin As String

    ilError = False
    If imFirstTime Then         'create the header record
        slStr = "Order#,Vehicle Name,Cash/Trade,Air Time/NTR,Local/Regional/National,Sales Source,Sales Office,Salesperson,Agency,Advertiser,Product,Order Type,Primary Competitive,Secondary Competitive,Invoice #,Year,Month,Gross Amount,Gross Split Amount,T-Net"
        
        On Error GoTo mWriteExportRecErr
        Print #hmEfficio, slStr     'write header description
        On Error GoTo 0

        slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
        slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")

        On Error GoTo mWriteExportRecErr
        Print #hmEfficio, slStr        'write header description
        On Error GoTo 0
        imFirstTime = False         'do the heading and time stamp only once
    End If

    'format the month info for a contract/vehicle
    slVehicle = ""
    slPrimComp = ""
    slSecComp = ""

    ilIndex = gBinarySearchVef(tlEfficioINfo.iVefCode)
    If ilIndex > 0 Then
        slVehicle = Trim$(tgMVef(ilIndex).sName)
    Else
        slVehicle = "Unknown vehicle-ID" & Trim$(Val(tlEfficioINfo.iVefCode))
    End If
        
    '1-22-12 obtain primary and secondary competitive codes to place into export
    tmMnfSrchKey.iCode = tlEfficioINfo.iMnfComp1
    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmMnf.sName = ""
    End If
    slPrimComp = Trim$(tmMnf.sName)
    
    If tlEfficioINfo.iMnfComp2 > 0 Then
        tmMnfSrchKey.iCode = tlEfficioINfo.iMnfComp2
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.sName = ""
        End If
        slSecComp = Trim$(tmMnf.sName)
    End If

    slSlsp = ""
    slOffice = ""
    slSS = ""
    slOrigin = ""
    ilIndex = gBinarySearchSlf(tlEfficioINfo.iSlfCode)
    If ilIndex <> -1 Then
        slSlsp = Trim$(tgMSlf(ilIndex).sFirstName) & " " & Trim$(tgMSlf(ilIndex).sLastName)
        For ilOfficeInx = LBound(tmSof) To UBound(tmSof)
            If tmSof(ilOfficeInx).iCode = tgMSlf(ilIndex).iSofCode Then
                slOffice = Trim$(tmSof(ilOfficeInx).sName)
                'now detrmine sales source from office
                For ilSSInx = LBound(tmMnfSS) To UBound(tmMnfSS) - 1
                    If tmMnfSS(ilSSInx).iCode = tmSof(ilOfficeInx).iMnfSSCode Then
                        slSS = tmMnfSS(ilSSInx).sName
                        If tmMnfSS(ilSSInx).iGroupNo = 1 Then
                            slOrigin = "L"
                        ElseIf tmMnfSS(ilSSInx).iGroupNo = 2 Then
                            slOrigin = "R"
                        ElseIf tmMnfSS(ilSSInx).iGroupNo = 3 Then
                            slOrigin = "N"
                        End If
                            
                        Exit For
                    End If
                Next ilSSInx
       
            End If
        Next ilOfficeInx
    End If

    slAgency = ""
    If tlEfficioINfo.iAgfCode = 0 Then       'Direct
        slAgency = "Direct"
    Else
        'do the binary search because if coming from past the agency wont be in memory
        ilIndex = gBinarySearchAgf(tlEfficioINfo.iAgfCode)
        If ilIndex <> -1 Then
            slAgency = Trim$(tgCommAgf(ilIndex).sName)
        End If
    End If

    slAdvt = ""     'Advertiser Name
    ilIndex = gBinarySearchAdf(tlEfficioINfo.iAdfCode)
    If ilIndex <> -1 Then
        slAdvt = Trim$(tgCommAdf(ilIndex).sName)
    End If

    'product, cash/trade, airtime/ntr are already strings

    For ilLoop = 1 To igPeriods       '36 months max
        If tlEfficioINfo.lNet(ilLoop) <> 0 Or tlEfficioINfo.lAcquisition(ilLoop) <> 0 Then      'do not create $0
            slStr = tlEfficioINfo.lCntrNo & ","
            slStr = slStr & """" & Trim$(slVehicle) & """" & ","
            slStr = slStr & """" & tlEfficioINfo.sCashTrade & """" & ","
            slStr = slStr & """" & tlEfficioINfo.sAirNTR & """" & ","
            slStr = slStr & """" & Trim$(slOrigin) & """" & ","
            slStr = slStr & """" & Trim$(slSS) & """" & ","
            slStr = slStr & """" & Trim$(slOffice) & """" & ","
            slStr = slStr & """" & Trim$(slSlsp) & """" & ","
            slStr = slStr & """" & Trim$(slAgency) & """" & ","
            slStr = slStr & """" & Trim$(slAdvt) & """" & ","
            slStr = slStr & """" & Trim$(tlEfficioINfo.sProduct) & """" & ","
            slStr = slStr & """" & tlEfficioINfo.sOrderType & """" & ","             '4-3-13
            slStr = slStr & """" & slPrimComp & """" & ","                  '1-22-12
            slStr = slStr & """" & slSecComp & """" & ","
            slStr = slStr & tlEfficioINfo.lInvoice & ","
            slStr = slStr & Trim$(str$(tlEfficioINfo.iYear(ilLoop))) & ","
            slStr = slStr & Trim$(str$(tlEfficioINfo.iMonth(ilLoop))) & ","
           
             't-net values
             If ExpEfficioRev!rbcNetBy(0).Value Then                 'use net vs tnet
                '1-22-12 the whole amt goes into the first slsp, as well as its split amt
                ilRemainder = tlEfficioINfo.lDirect(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlEfficioINfo.lDirect(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlEfficioINfo.lDirect(ilLoop), 2)) & ","
                End If
                'gross split
                ilRemainder = tlEfficioINfo.lGross(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlEfficioINfo.lGross(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlEfficioINfo.lGross(ilLoop), 2)) & ","
                End If
                'net split
                ilRemainder = tlEfficioINfo.lNet(ilLoop) Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(tlEfficioINfo.lNet(ilLoop), 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
                Else
                    slStr = slStr & Trim$(gLongToStrDec(tlEfficioINfo.lNet(ilLoop), 2))
                End If
            Else
                 llAdjPromoMerch = 0
                 If tlEfficioINfo.sCashTrade = "P" Or tlEfficioINfo.sCashTrade = "M" Then      'promo or merch, do not show under gross split column
                     llAdjPromoMerch = tlEfficioINfo.lGross(ilLoop)       'save gross split in case not a promo/merch amt, need to show it in gross split column.  Otherwise, blank the gross split column and
                                                                     'calc the Tnet by subtracting out Promo/merch amt
                     'tlEfficioInfo.lGross(ilLoop) & .lDirect should show zero for a merch or promo in the Gross split column, only need to subtract the amt for Tnet
                     tlEfficioINfo.lDirect(ilLoop) = 0
                     tlEfficioINfo.lGross(ilLoop) = 0
                     tlEfficioINfo.lNet(ilLoop) = -tlEfficioINfo.lNet(ilLoop)      'promo & merch must be subtracted
                 Else
                     llAdjPromoMerch = llAdjPromoMerch
                 End If
    
                 '1-22-12 the whole amt goes into the first slsp, as well as its split amt
                 'Gross Direct (entire amount of month for all slsp splits; could be same as net field if no splits)
                 ilRemainder = tlEfficioINfo.lDirect(ilLoop) Mod 100
                 If ilRemainder = 0 Then         'strip off the pennies if whole number
                     slStripCents = Trim$(gLongToStrDec(tlEfficioINfo.lDirect(ilLoop), 2))
                     slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                 Else
                     slStr = slStr & Trim$(gLongToStrDec(tlEfficioINfo.lDirect(ilLoop), 2)) & ","
                 End If
                 
                 ilRemainder = tlEfficioINfo.lGross(ilLoop) Mod 100           'gross split
                 If ilRemainder = 0 Then         'strip off the pennies if whole number
                     slStripCents = Trim$(gLongToStrDec(tlEfficioINfo.lGross(ilLoop), 2))
                     slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3)) & ","
                 Else
                     slStr = slStr & Trim$(gLongToStrDec(tlEfficioINfo.lGross(ilLoop), 2)) & ","
                 End If
           
                'compute the final TNet value:  gross minus comm minus acquisition minus promo/merch
                llTNet = tlEfficioINfo.lNet(ilLoop) - tlEfficioINfo.lAcquisition(ilLoop)     'net minus acquisition costs
                'get the triple net
                ilRemainder = llTNet Mod 100
                If ilRemainder = 0 Then         'strip off the pennies if whole number
                    slStripCents = Trim$(gLongToStrDec(llTNet, 2))
                    slStr = slStr & Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
                Else
                    slStr = slStr & Trim$(gLongToStrDec(llTNet, 2))
                End If
            End If
            
            On Error GoTo mWriteExportRecErr
            Print #hmEfficio, slStr
            On Error GoTo 0
        End If
    Next ilLoop

    For ilLoop = 1 To MAXMONTHS            'init the monthly info for next one
        tlEfficioINfo.iYear(ilLoop) = 0
        tlEfficioINfo.iMonth(ilLoop) = 0
        tlEfficioINfo.lGross(ilLoop) = 0
        tlEfficioINfo.lNet(ilLoop) = 0
        tlEfficioINfo.lDirect(ilLoop) = 0            '1-22-12
        tlEfficioINfo.lAcquisition(ilLoop) = 0
    Next ilLoop

    mWriteExportRec = ilError
    Exit Function

mWriteExportRecErr:
    ilError = True
    Resume Next

End Function

'*****************************************************************************************
'
'                   mCreateEfficioRev - Efficio export of monthly revenue from past for a single month
'
'                   <Input>  llStdStartDates - array of up to 25 start dates, denoting
'                                              start date of each period to gather
'                            llLastBilled - Date of last invoice period
'                            ilLastbilledInx - Index into llStdStartDates of period last
'                                           invoiced
'                   <Return> 0 = OK, <> 0= error

Function mCreateEfficioRev(llStdStartDates() As Long, llLastBilled As Long, ilLastBilledInx As Integer) As Integer
    Dim ilCurrentRecd As Integer            'index processing from tlRvf array
    Dim ilRet As Integer
    Dim slCode As String
    Dim slStr As String
    Dim llNet As Long
    Dim llGross As Long
    Dim llAcquisition As Long
    Dim llDate As Long
    Dim tlTranType As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim ilLoopRecv As Integer                   '9-16-01
    Dim tlEfficioINfo As MATRIXINFO
    Dim ilError As Integer
    Dim ilValidTran As Integer
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilLoopOnSlsp As Integer
    Dim llTempGross As Long
    Dim llTempNet As Long
    Dim llTempAcquisition As Long
    Dim llSplitGrossAmt As Long
    Dim llSplitNetAmt As Long
    Dim llSplitAcquisitionAmt As Long
    Dim llTempPct As Long
    Dim slGrossAmount As String
    Dim slNetAmount As String
    Dim slAcqAmount As String
    Dim slSharePct As String
    Dim ilReverseSign As Integer
    Dim slSplitGross As String
    Dim slSplitNet As String
    Dim ilUseSlsComm As Integer
    Dim ilMnfSubCo As Integer
    Dim ilLoop As Integer
    Dim ilSaveLastBilledInx As Integer
    Dim blTestOnlyPromoMerch As Boolean
    Dim blValidVehicle As Boolean
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean

    
    ilError = False

    ilUseSlsComm = False                'used for subroutine parameter
    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True

    tlTranType.iMerch = False               'default to exclude Merchandise & promotions unless its tnet
    tlTranType.iPromo = False
    If ExpEfficioRev!ckcNTR.Value = vbChecked Then                       'include NTR?
        tlTranType.iNTR = True
    Else
        tlTranType.iNTR = False
    End If
    
    If ExpEfficioRev!rbcNetBy(1).Value Then         'tNet?  if so, need to obtain the promo & merchandising transactions all the way to the end of requested period
        tlTranType.iMerch = True
        tlTranType.iPromo = True
        ilSaveLastBilledInx = ilLastBilledInx
        ilLastBilledInx = igPeriods             'tnet needs to get future Merch & Promo transactions from receivables
    End If
        
    For ilLoopRecv = 1 To ilLastBilledInx        'loop on # months to process for phf & rvf by contract # & tran date
        If (llStdStartDates(ilLoopRecv + 1) - 1) > llLastBilled Then
            tlTranType.iAdj = False              'ignore AN for tran dates in the future
            tlTranType.iNTR = False
            tlTranType.iHardCost = False
        End If

        slStr = Format$(llStdStartDates(ilLoopRecv), "m/d/yy")
        slCode = Format$(llStdStartDates(ilLoopRecv + 1) - 1, "m/d/yy")
        ilRet = gObtainPhfRvfbyCntr(ExpEfficioRev, 0, slStr, slCode, tlTranType, tlRvf())
        If ilRet = 0 Then
            'gLogMsg "** Error in reading History or Receivables- export aborted **", "EfficioExport.txt", False
            'Print #hmMsg, "** Error in reading History or Receivables- export aborted **"
            gAutomationAlertAndLogHandler "** Error in reading History or Receivables- export aborted **"
            ilError = True
            mCreateEfficioRev = ilError
            Exit Function
        End If

        For ilCurrentRecd = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
            tmRvf = tlRvf(ilCurrentRecd)

            ilValidTran = False
            'test for inclusion/exclusion of ntrs & air time transactions
            If (tlTranType.iNTR = True And tmRvf.iMnfItem > 0) Or (tmRvf.iMnfItem = 0) Then
                ilValidTran = True
            End If
            
            blValidVehicle = True
'            If Not gFilterLists(tmRvf.iAirVefCode, imIncludeCodes, imUseCodes()) Then
'                blValidVehicle = False
'                ilValidTran = False           'not a selected vehicle; bypass
'            End If
            
            If tmRvf.sCashTrade <> "P" And tmRvf.sCashTrade <> "M" Then
                'cash/trade IN or AN, ignore in future for standard, ok to include for calendar
                'IN/AN for airtime or NTR cant be in the future
                If (llStdStartDates(ilLoopRecv + 1) - 1) > llLastBilled Then
                    ilValidTran = False
                End If
            End If

            gPDNToLong tmRvf.sNet, llNet
            gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
            slCode = Format$(llDate, "m/d/yy")
            slCode = gObtainEndStd(slCode)          '11-5-13 cannot assume the month is the proper month that should be in the exported.
                                                    'ie NTR may be billed at the end of a cal month, but within the start of the std bdcst

            'Setup month and year to store in export
            gObtainYearMonthDayStr slCode, True, slYear, slMonth, slDay
 
            'PHFRVF routine has filtered only "I" & "HI" and "AN", along with the trans dates
            'see if selective contract for debugging
            'ignore Installment types of "I", which is billing, not revenue
            If (llNet <> 0 And ilValidTran) And (lmCntrNo = 0 Or (lmCntrNo <> 0 And lmCntrNo = tmRvf.lCntrNo)) And (tmRvf.sType <> "I") Then               'dont write out zero records
                'get contract from history or rec file if different than previous read
                'If tmRvf.lCntrNo <> tmChf.lCntrNo Then
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                         ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop

                    gFakeChf tmRvf, tmChf
                    ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
                    ReDim imSlfCode(0 To 9) As Integer             '4-20-00
                    ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
                    ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
                    ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

                    ilMnfSubCo = gGetSubCmpy(tmChf, imSlfCode(), lmSlfSplit(), tmRvf.iAirVefCode, ilUseSlsComm, lmSlfSplitRev())

                    'slsp, agency & advt & vehicles are in memory
                'End If

                If ExpEfficioRev!rbcNetBy(0).Value Then             'net (vs tnet)
                    gPDNToLong tmRvf.sGross, llGross
                    gPDNToLong tmRvf.sNet, llNet
                Else        'tnet

                    gPDNToLong tmRvf.sGross, llGross
                    gPDNToLong tmRvf.sNet, llNet
                    llAcquisition = tmRvf.lAcquisitionCost
'                    If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then
'                        llGross = -llGross
'                        llNet = -llNet
'                    End If

                    '7/31/15 implement acq commission  if applicable
                    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                        ilAcqCommPct = 0
                        blAcqOK = gGetAcqCommInfoByVehicle(tmRvf.iBillVefCode, ilAcqLoInx, ilAcqHiInx)
                        ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)
                        gCalcAcqComm ilAcqCommPct, llAcquisition, llAcqNet, llAcqComm
                        llAcquisition = llAcqNet
                    End If
                End If

                tlEfficioINfo.lCntrNo = tmChf.lCntrNo
                tlEfficioINfo.sOrderType = tmChf.sType
                tlEfficioINfo.sAirNTR = "A"          'assume Air time
                'if NTR, get that commission instead
                If tmRvf.iMnfItem > 0 Then          'this indicates NTR
                    tlEfficioINfo.sAirNTR = "N"
                    'dont need any info from the NTR record at this time
                    'retrieve the associated NTR record from SBF
                    'tmSbfSrchKey.lCode = tmRvf.lSbfCode  '12-16-02
                    'ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Print #hmMsg, "** Error in NTR transaction, NTR code " & Str$(tmRvf.lSbfCode) & " is invalid"
                    '    ilError = True
                    '    mCrMatrixPast = ilError
                    '    Exit Function
                    'End If
                End If                                                'associated sales source
                
                '1-22-12 obtain primary and secondary competitive codes
                tlEfficioINfo.iMnfComp1 = tmChf.iMnfComp(0)
                tlEfficioINfo.iMnfComp2 = tmChf.iMnfComp(1)
                '4-3-13 Order type :   Standard, PI, DR, Reservation, PSA, Promo, etc
                tlEfficioINfo.sOrderType = tmChf.sType
                
                lmTempGross = llGross
                lmTempNet = llNet
                lmTempAcquisition = llAcquisition
                lmTempPct = 0
                'determine amt of revenue sharing; could exceed 100%
                For ilLoopOnSlsp = 0 To 9
                    lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                Next ilLoopOnSlsp
                ilReverseSign = False
                If llNet < 0 Then
                    ilReverseSign = True            'always work with positive #s
                    lmTempGross = -lmTempGross
                    lmTempNet = -lmTempNet
                    lmTempAcquisition = -lmTempAcquisition
                    '3-18-14 invoice that was UNDO didnt show correctly
                    llGross = -llGross
                    llNet = -llNet
                    llAcquisition = -llAcquisition
                End If
    
                For ilLoopOnSlsp = 0 To 9
                    If lmSlfSplit(ilLoopOnSlsp) > 0 Then
                        tlEfficioINfo.iVefCode = tmRvf.iAirVefCode
                        'tlEfficioInfo.iSlfCode = tmRvf.iSlfCode
                        tlEfficioINfo.iSlfCode = imSlfCode(ilLoopOnSlsp)
                        tlEfficioINfo.iAgfCode = tmRvf.iAgfCode
                        tlEfficioINfo.iAdfCode = tmRvf.iAdfCode
                        tlEfficioINfo.sCashTrade = tmRvf.sCashTrade
                        If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                            If ilReverseSign Then
                                tlEfficioINfo.lDirect(ilLoopRecv) = -llGross     'working in all position #s, need to negate it if it was negative trans amt
                            Else
                                tlEfficioINfo.lDirect(ilLoopRecv) = llGross
                            End If
                        End If

                        'tlEfficioInfo.lGross(1) = llGross
                        'tlEfficioInfo.lNet(1) = llNet
                        
                        mObtainSlsRevenueShare llGross, llNet, llAcquisition, ilLoopOnSlsp, tlEfficioINfo, ilLoopRecv, ilReverseSign

                        tmPrfSrchKey.lCode = tmRvf.lPrfCode     'Product
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmPrf.sName = ""
                        End If
                        tlEfficioINfo.sProduct = tmPrf.sName
                        If Trim$(tlEfficioINfo.sProduct) = "" Then
                            tlEfficioINfo.sProduct = Trim$(tmChf.sProduct)
                        End If
                        tlEfficioINfo.iYear(ilLoopRecv) = Val(slYear)
                        tlEfficioINfo.iMonth(ilLoopRecv) = Val(slMonth)
                        tlEfficioINfo.lInvoice = tmRvf.lInvNo
 
                        ilRet = mWriteExportRec(tlEfficioINfo)
                        If ilRet <> 0 Then   'error
                            'gLogMsg "** Error writing export record for contract # " & str$(tmRvf.lCntrNo) & " from Receivables/History", "EfficioExport.txt", False
                            'Print #hmMsg, "** Error writing export record for contract # " & str$(tmRvf.lCntrNo) & " from Receivables/History"
                            gAutomationAlertAndLogHandler "** Error writing export record for contract # " & str$(tmRvf.lCntrNo) & " from Receivables/History"
                            ilError = True
                            mCreateEfficioRev = ilError
                            Exit Function
                        End If
                    End If
                    
                Next ilLoopOnSlsp       'for illooponslsp = 0 to 9
            End If
        Next ilCurrentRecd
    Next ilLoopRecv

    If ExpEfficioRev!rbcNetBy(1).Value Then         'tnet
        ilLastBilledInx = ilSaveLastBilledInx
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slNTR  As String
    Dim slCntr As String
    Dim slMonthBy As String

    ilRet = 0
    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgDBPath & "\Messages\" & "EfficioExport.txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    

    If ckcNTR.Value = vbChecked Then
        slNTR = "Include NTR"
    Else
       slNTR = "Exclude NTR"
    End If
    If edcContract.Text = "" Then
        slCntr = "All contracts"
    Else
        slCntr = "Cntr # " & edcContract.Text
    End If
    
    If rbcMonthBy(0).Value = True Then
        slMonthBy = "Std"
    Else
        slMonthBy = "Cal"
    End If
    smExportMesg = Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & slMonthBy & " " & edcMonth.Text & " " & edcYear.Text & " " & edcNoMonths.Text & " months, " & slNTR & ", " & slCntr & " **"

    If igRptCallType = EXP_EFFICIOREV Then
        'Print #hmMsg, "** Export Efficio Revenue: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & slMonthBy & " " & edcMonth.Text & " " & edcYear.Text & " " & edcNoMonths.Text & " months, " & slNTR & ", " & slCntr & " **"
        'Print #hmMsg, "** Export Efficio Revenue: " & Trim$(smExportMesg)
        gAutomationAlertAndLogHandler "** Export Efficio Revenue: " & Trim$(smExportMesg)
    Else
        'Print #hmMsg, "** Export Efficio Projections: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " for " & slMonthBy & " " & edcMonth.Text & " " & edcYear.Text & " " & edcNoMonths.Text & " months, " & slNTR & ", " & slCntr & " **"
        'Print #hmMsg, "** Export Efficio Projections: " & Trim$(smExportMesg)
        gAutomationAlertAndLogHandler "** Export Efficio Projections: " & Trim$(smExportMesg)
    End If
    

    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Sub ckcAll_Click()
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    'mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'           Efficio export is an export to extract History/Receivables data only, for 1 designated month.
'           It is to be run on-demand, by the Menu Export list.\
'           A comma delimited file is created which is stored in the Export folder defined in Traffic.ini
'
'           The Efficio module has been copied from Matrix module.  Some options remain left in for possible
'           future use.  Hidden until further notice.  Some options remain hidden and defaulted for parameter usage.
'
Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Dim slMonthHdr As String * 36
    Dim ilSaveMonth As Integer
    Dim ilYear As Integer
    'Dim llStdStartDates(1 To 37) As Long   '3 years standard start dates
    Dim llStdStartDates(0 To 37) As Long   '3 years standard start dates, ignore index zero
    'Dim llStartDates(1 To 37) As Long       'max 3 years
    Dim llStartDates(0 To 37) As Long       'max 3 years, ignore index zero
    Dim llLastBilled As Long
    Dim ilLastBilledInx As Integer
    Dim slStart As String
    Dim slTimeStamp As String
    Dim ilHowManyDefined As Integer
    Dim ilHowMany As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim llSelectedRevMonth As Long
    Dim ilError As Integer
    Dim llTempDate As Long

    lacInfo(0).Visible = False
    lacInfo(1).Visible = False

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = ExpEfficioRev!edcMonth.Text             'month in text form (jan..dec, or 1-12
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then

            ExpEfficioRev!edcMonth.SetFocus                 'invalid month
            ''MsgBox "Invalid Month", vbOkOnly + vbApplicationModal, " Month"
            gAutomationAlertAndLogHandler "Invalid Month", vbOkOnly + vbApplicationModal, "Month"
            Exit Sub
        End If
    End If


    slStr = ExpEfficioRev!edcYear.Text
    ilYear = gVerifyYear(slStr)
    If ilYear = 0 Then

        ExpEfficioRev!edcYear.SetFocus                 'invalid year
        ''MsgBox "Invalid Year", vbOkOnly + vbApplicationModal, " Year"
        gAutomationAlertAndLogHandler "Invalid Year", vbOkOnly + vbApplicationModal, "Year"
        Exit Sub
    End If
    
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
    llLastBilled = gDateValue(slStr)            'convert last month billed to long

    bmStdExport = True                          'assume standard exporting
    If rbcMonthBy(1).Value Then
        bmStdExport = False
    End If

    If igRptCallType = EXP_EFFICIOREV Then  'user is allowed to change the month in the past for revenue
        If bmStdExport Then
            'the month/year entered cannot be in the future
            slStr = Trim$(Val(ilSaveMonth)) & "/15/" & Trim$(Val(ilYear))
            slStart = gObtainStartStd(slStr)
            llSelectedRevMonth = gDateValue(slStart)
            If llSelectedRevMonth > llLastBilled Then
                ExpEfficioRev!edcMonth.SetFocus                 'invalid
                ''MsgBox "Revenue Month/Year selected has not been invoiced", vbOkOnly + vbApplicationModal, "Revenue Month/Year"
                gAutomationAlertAndLogHandler "Revenue Month/Year selected has not been invoiced", vbOkOnly + vbApplicationModal, "Revenue Month/Year"
                Exit Sub
            End If
        Else            ' Calendar
            slStart = str$(ilSaveMonth) & "/01/" & Trim$(str(ilYear))
        
        End If
    Else                  'projections user is not allowed to change projection start date
        If bmStdExport Then
            slStart = Format(llLastBilled + 1, "m/d/yy")
        Else
            slStart = str$(ilSaveMonth) & "/01/" & Trim$(str(ilYear))       'get the current std month to increment to next month for calendar
            
        End If
        
    End If
    
    'always 1 for revenue, or 36 for projections
    slStr = ExpEfficioRev!edcNoMonths.Text            '#periods
    igPeriods = Val(slStr)

    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = ExpEfficioRev!edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    'smExportFile contains the name to use which has been moved to edcTo.Text
    smExportName = Trim$(edcTo.Text)
    If Len(smExportName) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If
    
    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then
        smExportName = sgExportPath & smExportName
    End If

    ilRet = 0
    'On Error GoTo cmcExportErr:
    'slDateTime = FileDateTime(smExportName)
    ilRet = gFileExist(smExportName)
    If ilRet = 0 Then
        'file already exists, do not overwrite
        ''MsgBox "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        gAutomationAlertAndLogHandler "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        Exit Sub
        'Kill smExportName
    End If

    If Not mOpenMsgFile() Then          'open message file
         cmcCancel.SetFocus
         Exit Sub
    End If
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmEfficio = FreeFile
    'Open smExportName For Output As hmEfficio
    ilRet = gFileOpen(smExportName, "Output", hmEfficio)
    If ilRet <> 0 Then
        'Print #hmMsg, "** Terminated **"
        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        Close #hmEfficio
        imExporting = False
        Screen.MousePointer = vbDefault
        'TTP 10011 - Error.Numner prevents MsgBox.  Additionally the Error # is stored in ilRet.
        'MsgBox "Open Error #" & str$(Error.Numner) & smExportName, vbOkOnly, "Open Error"
        ''MsgBox "Open Error #" & str$(ilRet) & " - " & smExportName, vbOkOnly, "Open Error"
        gAutomationAlertAndLogHandler "Open Error #" & str$(ilRet) & " - " & smExportName, vbOkOnly, "Open Error"
        Exit Sub
    End If
    'Print #hmMsg, "** Storing Output into " & smExportName & " **"
    gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
    If rbcMonthBy(0).Value = True Then
        gAutomationAlertAndLogHandler "* Calendar = Standard"
    Else
        gAutomationAlertAndLogHandler "* Calendar = Monthly"
    End If
    gAutomationAlertAndLogHandler "* Month = " & edcMonth.Text
    gAutomationAlertAndLogHandler "* Year = " & edcYear.Text
    gAutomationAlertAndLogHandler "* # months = " & edcNoMonths.Text
    If ckcNTR.Value = vbChecked Then
        gAutomationAlertAndLogHandler "* Include NTR Revenue = True"
    Else
        gAutomationAlertAndLogHandler "* Include NTR Revenue = False"
    End If
    If rbcNetBy(0).Value = True Then
        gAutomationAlertAndLogHandler "* Dollars = Net"
    Else
        gAutomationAlertAndLogHandler "* Dollars = T-Net"
    End If
    If ckcAll.Value = vbChecked Then
        gAutomationAlertAndLogHandler "* Include All Vehicles = True"
    Else
        gAutomationAlertAndLogHandler "* Include All Vehicles = False"
    End If
    gAutomationAlertAndLogHandler "* Contract = " & edcContract.Text
    
    Screen.MousePointer = vbHourglass
    imExporting = True
   
    If mOpenEfficioFiles() = 0 Then
        'vehicle selection not an option at this time, get the all from list box
        ReDim imUseCodes(0 To 1) As Integer 'zero index is ignored and in gFilterLists
        ilHowManyDefined = lbcVehicle.ListCount
        ilHowMany = lbcVehicle.SelCount
        If ilHowMany > (ilHowManyDefined / 2) + 1 Then   'more than half selected
            imIncludeCodes = False
        Else
            imIncludeCodes = True
        End If
    
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = tgUserVehicle(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If lbcVehicle.Selected(ilLoop) And imIncludeCodes Then               'selected ?
                imUseCodes(UBound(imUseCodes)) = Val(slCode)
                'ReDim Preserve imUseCodes(1 To UBound(imUseCodes) + 1)
                ReDim Preserve imUseCodes(0 To UBound(imUseCodes) + 1)  'Index zero ignored
            Else        'exclude these
                If (Not lbcVehicle.Selected(ilLoop)) And (Not imIncludeCodes) Then
                    imUseCodes(UBound(imUseCodes)) = Val(slCode)
                    'ReDim Preserve imUseCodes(1 To UBound(imUseCodes) + 1)
                    ReDim Preserve imUseCodes(0 To UBound(imUseCodes) + 1)  'Index zero ignored
                End If
            End If
        Next ilLoop
        
        If bmStdExport Then              'std
            gBuildStartDates slStart, 1, igPeriods + 1, llStdStartDates()  ' llLastBilled, ilLastBilledInx  'build array of std start & end dates
            'determine what month index the actual is (versus the future dates)
            'assume everything in the past if by std
    
            If llLastBilled >= llStdStartDates(igPeriods + 1) Then  'all in past
               ilLastBilledInx = igPeriods
            End If
    
            For ilLoop = 1 To igPeriods Step 1
               If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then
                   ilLastBilledInx = ilLoop
                   Exit For
               End If
            Next ilLoop
        
            ilError = mCreateEfficioRev(llStdStartDates(), llLastBilled, ilLastBilledInx)       'get past (receivables)
            If ilError = True Then
                Close #hmEfficio
                mCloseEfficioFiles
                Erase tmSof, tmMnfSS
                Erase lmSlfSplit, imSlfCode, imslfcomm, imslfremnant, lmSlfSplitRev
            Else
                If igRptCallType = EXP_EFFICIOPROJ Then
                    ilRet = mEfficioProj(llStdStartDates(), llLastBilled)
                End If
                
                Close #hmEfficio
                mCloseEfficioFiles
                Erase llStdStartDates
                Erase tmSof, tmMnfSS
                Erase lmSlfSplit, imSlfCode, imslfcomm, imslfremnant, lmSlfSplitRev
        
                Screen.MousePointer = vbDefault
            End If
        Else
            gBuildStartDates slStart, 4, igPeriods + 1, llStdStartDates()  ' llLastBilled, ilLastBilledInx  'build array of std start & end dates
            mCrCalendarEfficio llStdStartDates()
        End If
'        'determine what month index the actual is (versus the future dates)
'        'assume everything in the past if by std
'
'        If llLastBilled >= llStdStartDates(igPeriods + 1) Then  'all in past
'           ilLastBilledInx = igPeriods
'        End If
'
'        For ilLoop = 1 To igPeriods Step 1
'           If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then
'               ilLastBilledInx = ilLoop
'               Exit For
'           End If
'        Next ilLoop
'
'        ilError = mCreateEfficioRev(llStdStartDates(), llLastBilled, ilLastBilledInx)
'        If ilError = True Then
'            Close #hmEfficio
'            mCloseEfficioFiles
'            Erase tmSof, tmMnfSS
'            Erase lmSlfSplit, imSlfCode, imslfcomm, imslfremnant, lmSlfSplitRev
'        Else
'            If igRptCallType = EXP_EFFICIOPROJ Then
'                ilRet = mEfficioProj(llStdStartDates(), llLastBilled)
'            End If
'
'            Close #hmEfficio
'            mCloseEfficioFiles
'            Erase llStdStartDates
'            Erase tmSof, tmMnfSS
'            Erase lmSlfSplit, imSlfCode, imslfcomm, imslfremnant, lmSlfSplitRev
'
'            Screen.MousePointer = vbDefault
'        End If
    Else
        lacInfo(0).Caption = "Open Error: Export Failed"
        'Print #hmMsg, "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        gAutomationAlertAndLogHandler "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    End If

    If ilError = 0 Then
    
        If igRptCallType = EXP_EFFICIOREV Then
            lacInfo(0).Caption = "Export Efficio Revenue Successfully Completed"
            'Print #hmMsg, "** Export Efficio Revenue Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            'Print #hmMsg, "** Export Efficio Revenue Successfully completed: " & Trim$(smExportMesg)
            gAutomationAlertAndLogHandler "** Export Efficio Revenue Successfully completed: " & Trim$(smExportMesg)
        Else
            lacInfo(0).Caption = "Export Efficio Projections Successfully Completed"
            'Print #hmMsg, "** Export Efficio Projections Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            'Print #hmMsg, "** Export Efficio Projections Successfully completed: " & Trim$(smExportMesg)
            gAutomationAlertAndLogHandler "** Export Efficio Projections Successfully completed: " & Trim$(smExportMesg)
        End If
    Else
        lacInfo(0).Caption = "Export Failed"
        'Print #hmMsg, "** Export Failed **"
        gAutomationAlertAndLogHandler "** Export Failed **"
    End If
    'lacInfo(1).Caption = "Export File: " & smExportName
    lacInfo(0).Visible = True
    'lacInfo(1).Visible = True
    Close #hmMsg
    cmcCancel.Caption = "&Done"
    If igExportType <= 1 Then       'ok to set focus if manual mode
        cmcCancel.SetFocus
    End If
    'cmcExport.Enabled = False
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub

Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If edcTo.Text = "" Then
        edcTo.Text = smExportName
    End If
End Sub
Private Sub edcContract_GotFocus()
    gCtrlGotFocus edcContract
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcMonth_Change()
    tmcClick_Timer
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub
Private Sub edcNoMonths_GotFocus()
    gCtrlGotFocus edcNoMonths
End Sub
Private Sub edcTo_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If

    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcYear_Change()
    tmcClick_Timer
End Sub

Private Sub edcYear_GotFocus()
    gCtrlGotFocus edcYear
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
    
    'retain option for Net or t-net.  Currently on t-net option default used
    PlcNetBy.Visible = False
    frcMonthBy.Visible = False
    DoEvents
    frcMonthBy.Visible = True
    
    '6/9/15: Replaced Acquisition with Barter
    '6-4-14 TNet always assumed for initial creation
    'If Not (Asc(tgSpf.sOverrideOptions) And SPACQUISITION) = SPACQUISITION Then
    If Not ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
        rbcNetBy(0).Value = True
        PlcNetBy.Visible = False
    Else
        'show Tnet, and default to TNet is acquisition used
        If rbcNetBy(1).Value Then
            rbcNetBy_Click 1
        Else
            rbcNetBy(1).Value = True
        End If
    End If
        

    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
        mInit
        'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix
        If igExportType <= 1 Then                       'manual from exports or manual from traffic
            Me.WindowState = vbNormal
            If (Asc(tgSaf(0).sFeatures1) And EFFICIOEXPORT) <> EFFICIOEXPORT Then
                'Print #hmMsg, "** Efficio Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                gLogMsg "** Efficio Export Disabled  ", "EfficioExport.txt", False
                If igExportType <= 1 Then           'manual mode, show disallowed on screen
                    lacInfo(0).AddItem "Efficio Export Disabled"
                End If
                imTerminate = True
                Exit Sub
            Else
                cmcExport.Enabled = True
            End If
        Else
            Me.WindowState = vbMinimized
            If (Asc(tgSaf(0).sFeatures1) And EFFICIOEXPORT) <> EFFICIOEXPORT Then
                gLogMsg "** Efficio Export Disabled  ", "EfficioExport.txt", False
               ' Print #hmMsg, "** Efficio Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                imTerminate = True
            End If
            If Not imTerminate Then
                gOpenTmf
                tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
                tmcSetTime.Enabled = True
                If igExportType = 2 Then                       'manual from exports or manual from traffic
                    gUpdateTaskMonitor 1, "EPE"
                ElseIf igExportType = 3 Then                       'manual from exports or manual from traffic
                    gUpdateTaskMonitor 1, "ERE"
                End If
                cmcExport_Click
                If igExportType = 2 Then                       'manual from exports or manual from traffic
                    gUpdateTaskMonitor 2, "EPE"
                ElseIf igExportType = 3 Then                       'manual from exports or manual from traffic
                    gUpdateTaskMonitor 2, "ERE"
                End If
            End If
            imTerminate = True
        End If
        tmcClick.Interval = 2000    '2 seconds
        tmcClick.Enabled = True
        If imTerminate Then
            cmcCancel_Click
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgAcqComm
    Erase tgAcqCommInx
    
    mCloseEfficioFiles

    If igExportType > 1 Then
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    Set ExpEfficioRev = Nothing   'Remve data segment
    
End Sub
Private Sub imcHelp_Click()

    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    Dim slDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slMonthStr As String * 36
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slStart As String
    Dim slReturn As String * 130        'this has to have a length
    Dim slCode As String
    Dim slFileName As String


    'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix

    slMonthStr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    igRptCallType = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)    'need to know if revenue or projections

    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))

    gCenterStdAlone ExpEfficioRev
 
    ilRet = gObtainAdvt()   'Build into tgCommAdf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainAgency() 'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    
   ' ilRet = gVffRead()
    
    ilRet = gObtainSalesperson() 'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If
    
    ilRet = gBuildAcqCommInfo(ExpEfficioRev)
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpEfficioRev
    imMnfRecLen = Len(tmMnf)

    hmVff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff)", ExpEfficioRev
    imVffRecLen = Len(tmVff)
    
    If igExportType >= 2 Then       'need to retrieve client namefor auto export; cannot execute this code in manual mode as a timing issue prevents the
                                    'filename from showing in the text box.
                                    'igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-Matrix

        
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
        
        On Error GoTo mObtainIniValuesErr
       'determine if coming from auto mode and to get parameters from Exports.ini
       'find exports.ini
        sgIniPath = gSetPathEndSlash(sgIniPath, True)
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "Exports.Ini"
        Else
            slFileName = CurDir$ & "\Exports.Ini"
        End If
        'look for matching sections which has the parameter options
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Calendar", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            rbcMonthBy(0).Value = True            'default to std
        Else
            If InStr(1, slReturn, "Std", vbTextCompare) > 0 Then
                rbcMonthBy(0).Value = True
            Else
                rbcMonthBy(1).Value = True
            End If
        End If
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Dollars", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            rbcNetBy(0).Value = True            'default to Net
        Else
            If InStr(1, slReturn, "TNet", vbTextCompare) > 0 Then
                rbcNetBy(1).Value = True
            Else
                rbcNetBy(0).Value = True
            End If
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Months", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            If igRptCallType = EXP_EFFICIOREV Then
                edcNoMonths.Text = 1                        'revenue is always current month invoiced
            Else                                            'projections max is 24 months
                edcNoMonths.Text = 24
            End If
        Else                                                'entry found.  Rev allowed 1 month, Projections allowed max 24
            slCode = Trim$(gStripChr0(slReturn))
            If igRptCallType = EXP_EFFICIOREV Then
                If Val(slCode) = 0 Or Val(slCode) > 1 Then
                    edcNoMonths.Text = 1
                Else
                    edcNoMonths.Text = Trim$(gStripChr0(slReturn))
                End If
            Else                                                'projections
                If Val(slCode) = 0 Or Val(slCode) > 24 Then     'max 24 months
                    edcNoMonths.Text = 24                       'invalid input , take default of 24 weeks
                Else
                    edcNoMonths.Text = Trim$(gStripChr0(slReturn))
                End If
            End If
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "NTR", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcNTR.Value = vbChecked
        Else
            If InStr(1, slReturn, "No", vbTextCompare) > 0 Then  'found No to exclude NTR
                ckcNTR.Value = vbUnchecked
            Else
                ckcNTR.Value = vbChecked
            End If
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Export", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'default to the export path
            sgExportPath = sgExportPath
        Else
            sgExportPath = Trim$(gStripChr0(slReturn))
        End If
        sgExportPath = gSetPathEndSlash(sgExportPath, True)

    End If
    
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), smLastBilled       'convert last bdcst billing date to string
    lmLastBilled = gDateValue(smLastBilled)            'convert last month billed to long
    
    slDate = Format$(smLastBilled, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If igRptCallType = EXP_EFFICIOREV Then
        'Default to last month billed
        ilMonth = Val(slMonth)
        ilYear = Val(slYear)
        edcNoMonths.Text = "1"           'only one month at a time for revenue
        smExportCaption = "Export Efficio Revenue"
    Else        'for projections, its the first unbilled date +1, but user not allowed to change it
       'hide user input parameters not relevant to Projections
        'lacMonth.Visible = False
        'edcMonth.Visible = False
        'lacStartYear.Visible = False
        'edcYear.Visible = False
        edcMonth.Enabled = False
        edcYear.Enabled = False
        lacNoMonths.Visible = False
        edcNoMonths.Visible = False
        edcNoMonths.Text = MAXMONTHS  '"36"              'do max 3 years to ensure all contracts in future processed
        slStart = Format(lmLastBilled + 1, "m/d/yy")
        'get the std month end date for filename
        slStr = gObtainEndStd(slStart)
        
        slDate = Format$(slStr, "m/d/yy")
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        ilYear = Val(slYear)
        ilMonth = Val(slMonth)
        smExportCaption = "Export Efficio Projections"
    End If

    edcMonth.Text = Mid$(slMonthStr, (ilMonth - 1) * 3 + 1, 3)
    edcYear.Text = Trim$(str$(ilYear))
    tmcClick.Enabled = True
    imSetAll = True
    
    lbcVehicle.Clear
    ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSPORT + VEHSELLING + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + VEHSPORT + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    ckcAll.Value = vbChecked
    
    smClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smClientName = Trim$(tmMnf.sName)
        End If
    End If

    
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
    
mObtainIniValuesErr:
    Resume Next

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
    Dim ilRet As Integer
    'Erase tmSOfficeCode
    'Erase tmSalesOffice

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpEfficioRev
    igManUnload = NO
End Sub

Private Sub lbcVehicle_Click()
  If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
   ' mSetCommands
End Sub


Private Sub PlcNetBy_Paint()
    PlcNetBy.CurrentX = 0
    PlcNetBy.CurrentY = 0
    PlcNetBy.Print "Export Net"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smExportCaption
End Sub

Private Sub rbcMonthBy_Click(Index As Integer)
    tmcClick_Timer
    Exit Sub
End Sub

Private Sub rbcNetBy_Click(Index As Integer)
    tmcClick_Timer
    Exit Sub
End Sub

Private Sub tmcClick_Timer()
Dim slRepeat As String
Dim ilRet As Integer
Dim slDateTime As String
Dim slMonthHdr As String * 36
Dim slStr As String
Dim ilYear As Integer
Dim slMonthBy As String * 3


    tmcClick.Enabled = False
    'Determine name of export (.txt file)
    slRepeat = "A"
    'month and year has been validated
    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    smMonth = ExpEfficioRev!edcMonth.Text             'month in text form (jan..dec, or 1-12
    gGetMonthNoFromString smMonth, imMonth          'if string month input, determine month #
    If imMonth = 0 Then                           'input isn't text month name, try month #
        imMonth = Val(smMonth)
        If imMonth = 0 Then
            Exit Sub
        End If
        smMonth = Mid$(slMonthHdr, (imMonth - 1) * 3 + 1, 3)
    End If

    smYear = ExpEfficioRev!edcYear.Text
    imYear = Val(smYear)
    slMonthBy = "Std"
    If rbcMonthBy(1).Value = True Then          'Calendar selected?
        slMonthBy = "Cal"
    End If
    
    Do
        ilRet = 0
        'On Error GoTo cmcExportDupNameErr:
        If igRptCallType = EXP_EFFICIOREV Then
            smExportName = sgExportPath & "Efficio Rev " & slMonthBy & " " & Trim$(smMonth) & Trim$(smYear)
        Else
            smExportName = sgExportPath & "Efficio Proj " & slMonthBy & " " & Trim$(smMonth) & Trim$(smYear)
        End If
        smExportName = smExportName & slRepeat & " " & gFileNameFilter(Trim$(smClientName)) & ".csv"            '2-27-14
        'slDateTime = FileDateTime(smExportName)
        ilRet = gFileExist(smExportName)
        If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
            slRepeat = Chr(Asc(slRepeat) + 1)
        End If
    Loop While ilRet = 0
    edcTo.Text = smExportName
    edcTo.Visible = True
    Exit Sub
'cmcExportDupNameErr:
'    ilRet = 1
'    Resume Next
End Sub
Private Sub mObtainSlsRevenueShare(llGross As Long, llNet As Long, llAcquisition As Long, ilLoopOnSlsp As Integer, tlEfficioINfo As MATRIXINFO, ilMonthInx As Integer, ilReverseSign As Integer)
Dim slStr As String
Dim slSharePct As String
Dim slGrossAmount As String
Dim slNetAmount As String
Dim llSplitNetAmt As Long
Dim llSplitGrossAmt As Long
Dim llSplitAcquisitionAmt As Long
Dim slAcqAmount As String

            
            If lmSlfSplit(ilLoopOnSlsp) = 0 Then
                tlEfficioINfo.lGross(ilMonthInx) = 0
                tlEfficioINfo.lNet(ilMonthInx) = 0
                tlEfficioINfo.lAcquisition(ilMonthInx) = 0
                Exit Sub
            End If
            lmTempPct = lmTempPct - lmSlfSplit(ilLoopOnSlsp)            'orignally the total % of all slsp (could exceed 100%)
                                                                        'as each slsp is processed, that % is subt from original.
                                                                        'the last gets the extra pennies
            slSharePct = gLongToStrDec(lmSlfSplit(ilLoopOnSlsp), 4)       'slsp share

            If lmTempPct <= 0 Then           '100 exhausted, last slsp gets extra pennies
                slSharePct = "100.0000"
                llSplitNetAmt = lmTempNet           'remainder of $ left to split, last one gets extra pennies
                llSplitGrossAmt = lmTempGross
                llSplitAcquisitionAmt = lmTempAcquisition
            Else
                slGrossAmount = gLongToStrDec(llGross, 2)
                slStr = gMulStr(slSharePct, slGrossAmount)                       ' gross portion of possible split
                llSplitGrossAmt = Val(gRoundStr(slStr, "1", 0))
                lmTempGross = lmTempGross - llSplitGrossAmt
                
                slNetAmount = gLongToStrDec(llNet, 2)
                slStr = gMulStr(slSharePct, slNetAmount)                       ' net portion of possible split
                llSplitNetAmt = Val(gRoundStr(slStr, "1", 0))
                lmTempNet = lmTempNet - llSplitNetAmt
                
                slAcqAmount = gLongToStrDec(llAcquisition, 2)
                slStr = gMulStr(slSharePct, slAcqAmount)                       ' acquisition portion of possible split
                llSplitAcquisitionAmt = Val(gRoundStr(slStr, "1", 0))
                lmTempAcquisition = lmTempAcquisition - llSplitAcquisitionAmt
                
            End If
            
            If ilReverseSign Then
                tlEfficioINfo.lGross(ilMonthInx) = -llSplitGrossAmt         '9-23-11 put $ in the month they belong in
                tlEfficioINfo.lNet(ilMonthInx) = -llSplitNetAmt
                tlEfficioINfo.lAcquisition(ilMonthInx) = -llSplitAcquisitionAmt
            Else
                tlEfficioINfo.lGross(ilMonthInx) = llSplitGrossAmt           '9-23-11 put $ in the month they belong in
                tlEfficioINfo.lNet(ilMonthInx) = llSplitNetAmt
                tlEfficioINfo.lAcquisition(ilMonthInx) = llSplitAcquisitionAmt
            End If
            

End Sub
'
'                   mSplitAndCreate - obtain all $ obtained from spots, create export records for each
'                   split salesperson
'
'                   Contract header and lines are in memory
Private Function mSplitAndCreate(llStartDates() As Long, tlEfficioINfo As MATRIXINFO, ilFirstProjInx As Integer, slCashAgyComm As String, ilVefCode As Integer) As Integer
Dim ilMnfSubCo As Integer
Dim ilCorT As Integer
Dim ilStartCorT As Integer
Dim ilEndCorT As Integer
Dim ilTemp As Integer
Dim llTempGross(0 To MAXMONTHS) As Long 'index zero ignored
Dim llTempNet(0 To MAXMONTHS) As Long   'index zero ignored
Dim llTempAcquisition(0 To MAXMONTHS) As Long   'index zero ignored
Dim slPctTrade As String
Dim ilLoopOnMonth As Integer
Dim ilLoopOnSlsp As Integer
Dim ilReverseSign As Integer
Dim slStr As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
Dim ilRet As Integer
Dim ilError As Integer
Dim blGotRevenue As Boolean

             ilError = False                        'error return from function
             slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0)
             
             If tgChfCT.iPctTrade = 0 Then                     'setup loop to do cash & trade
                 ilStartCorT = 1
                 ilEndCorT = 1
             ElseIf tgChfCT.iPctTrade = 100 Then
                 ilStartCorT = 2
                 ilEndCorT = 2
             Else
                 ilStartCorT = 1     'split cash/trade
                 ilEndCorT = 2
             End If

            'create an output line
            ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
            ReDim imSlfCode(0 To 9) As Integer             '4-20-00
            ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
            ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
            ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

            ilMnfSubCo = gGetSubCmpy(tgChfCT, imSlfCode(), lmSlfSplit(), ilVefCode, False, lmSlfSplitRev())

            For ilCorT = ilStartCorT To ilEndCorT Step 1        '2 passes if split cash/trade
                'move the projected values into temporary work area, as they might be modified
                'due to split cash/trade calculations
                blGotRevenue = False
                For ilTemp = 1 To MAXMONTHS
                    llTempGross(ilTemp) = lmProject(ilTemp)
                    llTempAcquisition(ilTemp) = lmAcquisition(ilTemp)
                    If lmProject(ilTemp) <> 0 Or lmAcquisition(ilTemp) <> 0 Then
                        blGotRevenue = True
                    End If
                Next ilTemp
                If blGotRevenue Then
'                    mCalcMonthAmt llTempGross(), llTempNet(), llTempAcquisition(), ilFirstProjInx, ilCorT, slPctTrade, slCashAgyComm
                    gCalcMonthAmt llTempGross(), llTempNet(), llTempAcquisition(), ilFirstProjInx, igPeriods, ilCorT, slCashAgyComm, tgChfCT
               
                    'Create the export records.  The
                    'Setup month and year to store in export
                    'use next months start date minus 1 to get the end date of the current month
                    For ilLoopOnMonth = ilFirstProjInx To igPeriods
                        If llTempNet(ilLoopOnMonth) <> 0 Or llTempAcquisition(ilLoopOnMonth) <> 0 Then           'bypass $0
                            lmTempGross = llTempGross(ilLoopOnMonth)
                            lmTempNet = llTempNet(ilLoopOnMonth)
                            lmTempAcquisition = llTempAcquisition(ilLoopOnMonth)
                            lmTempPct = 0
                            'determine amt of revenue sharing; could exceed 100%
                            For ilLoopOnSlsp = 0 To 9
                                lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                            Next ilLoopOnSlsp
                            ilReverseSign = False
                            'all schedule line $ are positive, negative from receivables only
    '                                If llNet < 0 Then
    '                                    ilReverseSign = True            'always work with positive #s
    '                                    lmTempGross = -lmTempGross
    '                                    lmTempNet = -lmTempNet
    '                                End If
                            
                            For ilLoopOnSlsp = 0 To 9
                                slStr = Format$(llStartDates(ilLoopOnMonth + 1) - 1, "m/d/yy")
                                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                                tlEfficioINfo.iYear(ilLoopOnMonth) = Val(slYear)
                                tlEfficioINfo.iMonth(ilLoopOnMonth) = Val(slMonth)
                                tlEfficioINfo.sAirNTR = "A"          'assume Air time
                                If ilCorT = 1 Then          'cash
                                    tlEfficioINfo.sCashTrade = "C"
                                Else
                                    tlEfficioINfo.sCashTrade = "T"
                                End If
                                
                                                                    
                                If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                                    tlEfficioINfo.lDirect(ilLoopOnMonth) = lmTempGross         'llGross
                                End If
    
                                mObtainSlsRevenueShare llTempGross(ilLoopOnMonth), llTempNet(ilLoopOnMonth), llTempAcquisition(ilLoopOnMonth), ilLoopOnSlsp, tlEfficioINfo, ilLoopOnMonth, False
                                tlEfficioINfo.iSlfCode = imSlfCode(ilLoopOnSlsp)
                                tlEfficioINfo.iVefCode = ilVefCode
                                ilRet = mWriteExportRec(tlEfficioINfo)
                                If ilRet <> 0 Then   'error
                                    'gLogMsg "Error writing export record for contract # " & str$(tgChfCT.lCntrNo) & ", Line # " & str$(tmClf.iLine), "EFFICIOEXPORT.txt", False
                                    'Print #hmMsg, "Error writing export record for contract # " & str$(tgChfCT.lCntrNo) & ", Line # " & str$(tmClf.iLine)
                                    gAutomationAlertAndLogHandler "Error writing export record for contract # " & str$(tgChfCT.lCntrNo) & ", Line # " & str$(tmClf.iLine)
                                    ilError = True
                                    mSplitAndCreate = ilError
                                    Exit Function
                                End If
                            Next ilLoopOnSlsp
                        End If              'llTempNet(ilLoopOnMonth) <> 0
                    Next ilLoopOnMonth
                End If                          'blGotRevenue
            Next ilCorT
            
            mSplitAndCreate = ilError
        Exit Function
End Function
'
'                       mEfficioNTR - include NTR into Efficio export only for a billed month
'
Private Function mEfficioNTR(tlSBFTypes As SBFTypes, llStartDates() As Long, tlEfficioINfo As MATRIXINFO, ilFirstProjInx As Integer, slCashAgyComm As String) As Integer
    Dim ilRet As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilFoundMonth As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilLoopOnMonth As Integer
    Dim slPctTrade As String
    Dim ilMnfSubCo As Integer
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim llAmt As Long
    Dim llAcquisitionAmt As Long
    Dim slGross As String
    Dim slGrossPct As String
    Dim slNet As String
    Dim slAcquisition As String
    Dim llNet As Long
    Dim llGross As Long
    Dim llAcquisition As Long
    Dim ilLoopOnSlsp As Integer
    Dim ilReverseSign As Integer
    Dim slStr As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilError As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim slCashAgyCommPct As String
    Dim blValidVehicle As Boolean
    ReDim tlSbf(0 To 0) As SBF

                ilError = False                         'error return
                
                slStart = Format$(llStartDates(1), "m/d/yy")
                slEnd = Format$(llStartDates(igPeriods + 1) - 1, "m/d/yy")
                If tgChfCT.iPctTrade = 0 Then                     'setup loop to do cash & trade
                    ilStartCorT = 1
                    ilEndCorT = 1
                ElseIf tgChfCT.iPctTrade = 100 Then
                    ilStartCorT = 2
                    ilEndCorT = 2
                Else
                    ilStartCorT = 1     'split cash/trade
                    ilEndCorT = 2
                End If
                 ilRet = gObtainSBF(ExpEfficioRev, hmSbf, tgChfCT.lCode, slStart, slEnd, tlSBFTypes, tlSbf(), 0)   '11-28-06 add last parm to indicate which key to use

                For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1
                    tmSbf = tlSbf(llSbf)
                    ilFoundMonth = False
                    gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
                    llDate = gDateValue(slDate)
                    For ilLoopOnMonth = ilFirstProjInx To igPeriods Step 1       'loop thru months to find the match
                        If llDate >= llStartDates(ilLoopOnMonth) And llDate < llStartDates(ilLoopOnMonth + 1) Then
                            ilFoundMonth = True
                            Exit For
                        End If
                    Next ilLoopOnMonth
                    
                    
                    If Not gFilterLists(tmSbf.iAirVefCode, imIncludeCodes, imUseCodes()) Then
                        blValidVehicle = False
                        ilFoundMonth = False            'not a selected vehicle; bypass
                    End If

                    If ilFoundMonth Then
                        If tmSbf.sAgyComm = "N" Then        'ntr comm flag overrides the contract
                            slCashAgyCommPct = ".00"
                        Else
                            slCashAgyCommPct = slCashAgyComm        'agy comm determine from the contracts agency info
                        End If

                        ReDim lmSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
                        ReDim imSlfCode(0 To 9) As Integer             '4-20-00
                        ReDim imslfcomm(0 To 9) As Integer             'slsp under comm %
                        ReDim imslfremnant(0 To 9) As Integer          'slsp under remnant %
                        ReDim lmSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)
    
                        ilMnfSubCo = gGetSubCmpy(tgChfCT, imSlfCode(), lmSlfSplit(), tmSbf.iAirVefCode, False, lmSlfSplitRev())

                        For ilCorT = ilStartCorT To ilEndCorT
                            If ilCorT = 1 Then              'cash
                                slPctTrade = gSubStr("100.", gIntToStrDec(tgChfCT.iPctTrade, 0))
                                tlEfficioINfo.sCashTrade = "C"
                            Else            'trade portion
                                slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0)
                                tlEfficioINfo.sCashTrade = "T"
                            End If
                            'convert the $ to gross & net strings
                            llAmt = tmSbf.lGross * tmSbf.iNoItems
                            slGross = gLongToStrDec(llAmt, 2)       'convert to xxxx.xx
                            slGrossPct = gSubStr("100.00", slCashAgyCommPct)        'determine  % to client (normally 85%)
                            slNet = gDivStr(gMulStr(slGrossPct, slGross), "100")    'net value
                            
                            llAcquisitionAmt = tmSbf.lAcquisitionCost * tmSbf.iNoItems
                            slAcquisition = gLongToStrDec(llAcquisitionAmt, 2)      'convert to xxxx.xx
                            
                            'calculate the new gross & net if split cash/trade
                            slNet = gDivStr(gMulStr(slNet, slPctTrade), "100")
                            llNet = gStrDecToLong(slNet, 2)
                            slGross = gDivStr(gMulStr(slGross, slPctTrade), "100")
                            llGross = gStrDecToLong(slGross, 2)
                            
                            slAcquisition = gDivStr(gMulStr(slAcquisition, slPctTrade), "100")
                            llAcquisition = gStrDecToLong(slAcquisition, 2)

                            If llNet <> 0 Or llAcquisition <> 0 Then           'bypass $0
                                lmTempGross = llGross
                                lmTempNet = llNet
                                lmTempAcquisition = llAcquisition
                                'determine amt of revenue sharing; could exceed 100%
                                For ilLoopOnSlsp = 0 To 9
                                    lmTempPct = lmTempPct + lmSlfSplit(ilLoopOnSlsp)
                                Next ilLoopOnSlsp
                                ilReverseSign = False
                                
                                For ilLoopOnSlsp = 0 To 9
                                    slStr = Format$(llStartDates(ilLoopOnMonth + 1) - 1, "m/d/yy")
                                    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                                    tlEfficioINfo.iYear(ilLoopOnMonth) = Val(slYear)
                                    tlEfficioINfo.iMonth(ilLoopOnMonth) = Val(slMonth)
                                    tlEfficioINfo.sAirNTR = "N"          'NTR flag
                                    If ilCorT = 1 Then          'cash
                                        tlEfficioINfo.sCashTrade = "C"
                                    Else
                                        tlEfficioINfo.sCashTrade = "T"
                                    End If
                                    
                                    If ilLoopOnSlsp = 0 Then            '1-22-12 1st slsp gets total gross amt as well as split in its record
                                        tlEfficioINfo.lDirect(ilLoopOnMonth) = llGross
                                    End If
                                    
                                    mObtainSlsRevenueShare llGross, llNet, llAcquisition, ilLoopOnSlsp, tlEfficioINfo, ilLoopOnMonth, False
                                    tlEfficioINfo.iSlfCode = imSlfCode(ilLoopOnSlsp)
                                    tlEfficioINfo.iVefCode = tmSbf.iBillVefCode
                                    'tlEfficioInfo.lGross(ilLoopOnMonth) = llGross
                                    'tlEfficioInfo.lNet(ilLoopOnMonth) = llNet
                                    ilRet = mWriteExportRec(tlEfficioINfo)
                                    If ilRet <> 0 Then   'error
                                        'gLogMsg "Error writing export record for NTR, contract # " & str$(tgChfCT.lCntrNo) & " Contract file", "EFFICIOEXPORT.txt", False
                                        'Print #hmMsg, "Error writing export record for NTR, contract # " & str$(tgChfCT.lCntrNo) & " Contract file"
                                        gAutomationAlertAndLogHandler "Error writing export record for NTR, contract # " & str$(tgChfCT.lCntrNo) & " Contract file"
                                        ilError = True
                                        mEfficioNTR = ilError
                                        Exit Function
                                    End If
                                Next ilLoopOnSlsp
                            End If
                        Next ilCorT
                    End If
                Next llSbf
                mEfficioNTR = ilError
            Exit Function
End Function
     '
'            Efficio Projection Export - Gather Projection data from contracts
'            Loop thru contracts within date last std bdcst billing
'            through the number of months requested (up to 36 months).
'
'                   <Input>  llStdStartDates - array of max 36 start dates, denoting
'                                              start date of each period to gather
'                            llLastBilled - Date of last invoice period
'
'********************************************************************************************
Function mEfficioProj(llStdStartDates() As Long, llLastBilled As Long) As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slStdStart As String            'start date to gather (std start)
    Dim slStdEnd As String              'end date to gather (end of std year)
    Dim slTempStart As String           'Start date of period requested (minus 1 month) to handle makegoods outside
                        'end date of contrct
    Dim slCntrStatus As String          'list of contract status to gather (working, order, hold, etc)
    Dim slCntrType As String            'list of contract types to gather (Per inq, Direct Response, Remnants, etc)
    Dim ilHOState As Integer            'which type of HO cntr states to include (whether revisions should be included)
    Dim llContrCode As Long
    Dim ilCurrentRecd As Integer
    Dim ilLoop As Integer
    Dim ilClf As Integer                'loop count for lines
    Dim ilTemp As Integer
    Dim llStdStart As Long              'requested start date to gather (serial date)
    Dim llStdEnd As Long                'requested end date to gather (serial date)
    Dim llDate As Long
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim slCashAgyComm As String
    Dim slPctTrade As String
    Dim slNet As String
    Dim llNet As Long
    Dim slGross As String
    Dim llGross As Long
    Dim slGrossPct As String
    Dim llAmt As Long
    Dim ilFoundMonth As Integer
    Dim ilFirstProjInx As Integer
    ReDim llTempGross(0 To MAXMONTHS) As Long      'max MAXMONTHS months projection, gross $, index zero ignored
    ReDim llTempNet(0 To MAXMONTHS) As Long        'max MAXMONTHS months projection, net $, index zero ignored
    Dim tlEfficioINfo As MATRIXINFO
    Dim ilLoopOnMonth As Integer
    ReDim tlSbf(0 To 0) As SBF
    Dim ilError As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim slDate As String
    Dim tlSBFTypes As SBFTypes
    Dim ilLoopOnSlsp As Integer
    Dim ilMnfSubCo As Integer
    Dim ilReverseSign As Integer
    Dim llSlsRevNetShare As Long
    Dim llSlsRevGrossShare As Long
    Dim blValidVehicle As Boolean
    Dim blFirstTime As Boolean
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean


    ilError = False             'assume everything is OK
    'determine what month index the actual is (versus the future dates)
    slStdStart = Format$(llStdStartDates(1), "m/d/yy")       'assume first date of proj is the quarter entered
    slStdEnd = Format$(llStdStartDates(igPeriods + 1), "m/d/yy")

    For ilLoop = 1 To igPeriods Step 1
        If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then
            ilFirstProjInx = ilLoop + 1
            slStdStart = Format$(llStdStartDates(ilFirstProjInx), "m/d/yy")
            Exit For
        End If
    Next ilLoop
    If ilFirstProjInx = 0 Then
        ilFirstProjInx = 1                          'all projections, no actuals
    End If
        If llLastBilled >= llStdStartDates(igPeriods + 1) Then   'all data was in the past only, dont do contracts
        Exit Function
    End If

    llStdStart = llStdStartDates(ilFirstProjInx)  'first date to project
    llStdEnd = llStdStartDates(igPeriods + 1)                'end date to project

    'setup type statement as to which type of SBF records to retrieve (only NTR)
    tlSBFTypes.iNTR = True          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing

    slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
    slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
    ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
    'build table (into tmChfAdvtExt) of all contracts that fall within the dates required

    slTempStart = Format$((gDateValue(slStdStart)), "m/d/yy")
    ilRet = gObtainCntrForDate(ExpEfficioRev, slTempStart, "", slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())

    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1                                          'loop while llCurrentRecd < llRecsRemaining

        llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())
        If lmCntrNo = 0 Or lmCntrNo <> 0 And lmCntrNo = tgChfCT.lCntrNo Then    'single contract for debugging only
 
             'obtain agency for commission, lines need them for net calculation, as well as for NTR
             If tgChfCT.iAgfCode > 0 Then
                slCashAgyComm = ".00"
                tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                End If
             Else
                 slCashAgyComm = ".00"
             End If              'iagfcode > 0
             
            'prepare common data to get exported
            tlEfficioINfo.lCntrNo = tgChfCT.lCntrNo
            'tlEfficioInfo.iSlfCode = tgChfCT.iSlfCode(0)
            tlEfficioINfo.iAgfCode = tgChfCT.iAgfCode
            tlEfficioINfo.iAdfCode = tgChfCT.iAdfCode
            tlEfficioINfo.sProduct = tgChfCT.sProduct
            '1-22-12 obtain primary and secondary competitive codes
            tlEfficioINfo.iMnfComp1 = tgChfCT.iMnfComp(0)
            tlEfficioINfo.iMnfComp2 = tgChfCT.iMnfComp(1)
            '4-3-13 Order type:  standard, psa, promo, dr, pi, etc
            tlEfficioINfo.sOrderType = tgChfCT.sType
            
            For ilLoop = 1 To igPeriods               '1-22-12
                tlEfficioINfo.lDirect(ilLoop) = 0
            Next ilLoop


            mEfficioNTR tlSBFTypes, llStdStartDates(), tlEfficioINfo, ilFirstProjInx, slCashAgyComm
            
            slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0) 're-establish if went to NTR

            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                tmClf = tgClfCT(ilClf).ClfRec
                
                If (tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E") Then
                    mBuildExportFlights ilClf, llStdStartDates(), ilFirstProjInx, igPeriods + 1
                    
                    'calc acq net if necessary
                    For ilLoopOnMonth = ilFirstProjInx To igPeriods
                        '7/31/15 implement acq commission  if applicable
                        If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE And ExpEfficioRev!rbcNetBy(1).Value Then
                            ilAcqCommPct = 0
                            blAcqOK = gGetAcqCommInfoByVehicle(tmClf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                            ilAcqCommPct = gGetEffectiveAcqComm(llStdStartDates(ilLoopOnMonth), ilAcqLoInx, ilAcqHiInx)
                            gCalcAcqComm ilAcqCommPct, lmAcquisition(ilLoopOnMonth), llAcqNet, llAcqComm
                            lmAcquisition(ilLoopOnMonth) = llAcqNet
                            
                        End If
                    Next ilLoopOnMonth
                    
                    ilRet = mSplitAndCreate(llStdStartDates(), tlEfficioINfo, ilFirstProjInx, slCashAgyComm, tmClf.iVefCode)
                End If              'tmclf.stype
                
                For ilLoop = 1 To MAXMONTHS            'init the projected gross & net values
                    llTempNet(ilLoop) = 0
                    lmProject(ilLoop) = 0
                    lmAcquisition(ilLoop) = 0
                    llTempGross(ilLoop) = 0
                Next ilLoop
            Next ilClf                          'next schedule line
        End If                                  'selective contract #
    Next ilCurrentRecd

    Erase llTempGross, llTempNet

    mEfficioProj = ilError
End Function
'
'
'                   mBuildexportFlights - Loop through the flights of the schedule line
'                                   and build the projections dollars into lmprojmonths array
'                   <input> ilclf = sched line index into tgClfCt
'                           llStdStartDates() - up to 36 std month start dates
'                           ilFirstProjInx - index of 1st month to start projecting
'                           ilHowManyPer - # entries containing a date to test in date array
'                   <output> lmProject = array of 36 months data corresponding to
'                                           36 std start months
'                           lmAcquisition - array of 36 months acquisition costs
'       2-28-14 implement tnet / net option
Sub mBuildExportFlights(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilHowManyPer As Integer)
Dim ilCff As Integer
Dim slStr As String
Dim llFltStart As Long
Dim llFltEnd As Long
Dim ilLoop As Integer
Dim llDate As Long
Dim llDate2 As Long
Dim llSpots As Long
Dim ilTemp As Integer
Dim llStdStart As Long
Dim llStdEnd As Long
Dim ilMonthInx As Integer
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilHowManyPer)
    ilCff = tgClfCT(ilClf).iFirstCff
    Do While ilCff <> -1
    tmCff = tgCffCT(ilCff).CffRec

    'first decide if its Cancel Before Start
    gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
    llFltStart = gDateValue(slStr)
    gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
    llFltEnd = gDateValue(slStr)
    If llFltEnd < llFltStart Then
        Exit Sub
    End If
    'gUnpackDate tmcff.iStartDate(0), tmcff.iStartDate(1), slStr
    'llFltStart = gDateValue(slStr)
    'backup start date to Monday
    ilLoop = gWeekDayLong(llFltStart)
    Do While ilLoop <> 0
        llFltStart = llFltStart - 1
        ilLoop = gWeekDayLong(llFltStart)
    Loop
    'gUnpackDate tmcff.iEndDate(0), tmcff.iEndDate(1), slStr
    'llFltEnd = gDateValue(slStr)
    'the flight dates must be within the start and end of the projection periods,
    'not be a CAncel before start flight, and have a cost > 0
    If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) And (tmCff.lActPrice > 0 Or tmClf.lAcquisitionCost > 0) Then
        'only retrieve for projections, anything in the past has already
        'been invoiced and has been retrieved from history or receiv files
        'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
        If llStdStart > llFltStart Then
            llFltStart = llStdStart
        End If
        'use flight end date or requsted end date, whichever is lesser
        If llStdEnd < llFltEnd Then
            llFltEnd = llStdEnd
        End If

        For llDate = llFltStart To llFltEnd Step 7
            'Loop on the number of weeks in this flight
            'calc week into of this flight to accum the spot count
            If tmCff.sDyWk = "W" Then            'weekly
                llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
            Else                                        'daily
                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                    llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                Else
                    llFltEnd = llDate + 6
                    If llDate > llFltEnd Then
                        llFltEnd = llFltEnd       'this flight isn't 7 days
                    End If
                    For llDate2 = llDate To llFltEnd Step 1
                        ilTemp = gWeekDayLong(llDate2)
                        llSpots = llSpots + tmCff.iDay(ilTemp)
                    Next llDate2
                End If
            End If
            'determine month that this week belongs in, then accumulate the gross and net $
            'currently, the projections are based on STandard bdcst
            For ilMonthInx = ilFirstProjInx To ilHowManyPer - 1 Step 1       '5-22-03 (-1 to adjust to not exceed max dates) loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    If tgClfCT(ilClf).ClfRec.sType = "E" Then            'pckage pricing where spot rate equals the entire price for the week
                        lmProject(ilMonthInx) = lmProject(ilMonthInx) + tmCff.lActPrice
                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + tmClf.lAcquisitionCost
                    Else
                        lmProject(ilMonthInx) = lmProject(ilMonthInx) + (llSpots * tmCff.lActPrice)
                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + (llSpots * tmClf.lAcquisitionCost)
                    Exit For
                End If
                End If
            Next ilMonthInx
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If                                          '
    ilCff = tgCffCT(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
    Exit Sub
End Sub
'
'           Create Calendar Month Efficio export with spot file
'
Public Sub mCrCalendarEfficio(llStartDates() As Long)
Dim ilRet As Integer
Dim slCntrTypes As String
Dim slCntrStatus As String
Dim ilHOState As Integer
Dim slStart As String
Dim slEnd As String
Dim llContrCode As Long
Dim ilCurrentRecd As Integer
Dim llLoopOnSpots As Long
Dim blFirstTime  As Boolean
Dim slPrevKey As String
Dim slCurrKey As String
Dim ilCurrLine As Integer
Dim ilPrevLine As Integer
Dim ilCurrVefCode As Integer
Dim ilPrevVefCode As Integer
Dim slTempKey As String
Dim slTempLine As String
Dim slTempVehicle As String
Dim ilMnfSubCo As Integer
Dim ilLoop As Integer
Dim ilVefCode As Integer
Dim llSpotInx As Long
Dim llAirDate As Long
Dim ilLoopOnMonth As Integer
Dim ilFoundMonth As Integer
Dim ilClf As Integer
Dim slCashAgyComm  As String
Dim tlEfficioINfo As MATRIXINFO
Dim tlSBFTypes As SBFTypes
Dim blValidVehicle As Boolean

        'setup type statement as to which type of SBF records to retrieve (only NTR)
        tlSBFTypes.iNTR = True          'include NTR billing
        tlSBFTypes.iInstallment = False      'exclude Installment billing
        tlSBFTypes.iImport = False           'exclude rep import billing


        slStart = Format$(llStartDates(1) - 30, "m/d/yy")
        slEnd = Format$(llStartDates(igPeriods + 1) - 1, "m/d/yy")           'end of last calendar month
        If lmCntrNo > 0 Then
            ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
            tmChfSrchKey1.lCntrNo = lmCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
               ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
            Else
                'setup 1 entry in the active contract array for processing single contract
                tmChfAdvtExt(0).lCntrNo = tmChf.lCntrNo
                tmChfAdvtExt(0).lCode = tmChf.lCode
                tmChfAdvtExt(0).iSlfCode(0) = tmChf.iSlfCode(0)
                tmChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
            End If
        Else
            'Gather all contracts for previous year and current year whose effective date entered
            'is prior to the effective date that affects either previous year or current year
            slCntrTypes = gBuildCntTypes()
            slCntrStatus = "HO"               'Sched Holds, orders
            
            ilHOState = 1                       'Sched holds or orders only, no unsch contracts since spots need to be retrieved
            ilRet = gObtainCntrForDate(ExpEfficioRev, slStart, slEnd, slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
        End If
        
        'readjust the start date to only pick up spots from the user requested period.  Backing it up to find the contracts was necessary in case
        'a mg/outside spot was sched after the contracts expiration date
        slStart = Format$(llStartDates(1), "m/d/yy")
        For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1     'loop on contracts


            'obtain the contract & lines and save the common header info
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())

            'obtain agency for commission, lines need them for net calculation, as well as for NTR
            If tgChfCT.iAgfCode > 0 Then
                slCashAgyComm = ".00"
                tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                End If
             Else
                 slCashAgyComm = ".00"
             End If              'iagfcode > 0
             
            'prepare common data to get exported
            tlEfficioINfo.lCntrNo = tgChfCT.lCntrNo
            'tlEfficioInfo.iSlfCode = tgChfCT.iSlfCode(0)
            tlEfficioINfo.iAgfCode = tgChfCT.iAgfCode
            tlEfficioINfo.iAdfCode = tgChfCT.iAdfCode
            tlEfficioINfo.sProduct = tgChfCT.sProduct
            '1-22-12 obtain primary and secondary competitive codes
            tlEfficioINfo.iMnfComp1 = tgChfCT.iMnfComp(0)
            tlEfficioINfo.iMnfComp2 = tgChfCT.iMnfComp(1)
            '4-3-13 Order type:  standard, psa, promo, dr, pi, etc
            tlEfficioINfo.sOrderType = tgChfCT.sType
            
            For ilLoop = 1 To igPeriods               '1-22-12
                tlEfficioINfo.lDirect(ilLoop) = 0
            Next ilLoop

            If tgChfCT.sNTRDefined = "Y" And ExpEfficioRev!ckcNTR.Value = vbChecked Then        'this has NTR billing
                mEfficioNTR tlSBFTypes, llStartDates(), tlEfficioINfo, 1, slCashAgyComm
            End If
            
            'slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0) 're-establish if went to NTR
            
            'search for all spots by for this contract
            ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
            'ReDim tmSdfExt(1 To 1) As SDFEXT
            ReDim tmSdfExt(0 To 0) As SDFEXT
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntrSpot(-1, False, llContrCode, -1, "S", slStart, slEnd, tmSdfExtSort(), tmSdfExt(), 0, False, True) 'search for spots between requested user dates
            blFirstTime = True
            'loop thru the sorted array to get spots by line, vehicle, date
            For llLoopOnSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
                llSpotInx = tmSdfExtSort(llLoopOnSpots).lSdfExtIndex
                'process only scheduled, makegood and outsides spots (ignore missed, cancelled hidden)
                
                blValidVehicle = True
'                If Not gFilterLists(tmSdfExt(llSpotInx).iVefCode, imIncludeCodes, imUseCodes()) Then      'filter vehicle if selected
'                    blValidVehicle = False
'                End If

                If ((tmSdfExt(llSpotInx).sSchStatus = "S" Or tmSdfExt(llSpotInx).sSchStatus = "G" Or tmSdfExt(llSpotInx).sSchStatus = "O") And (tmSdfExt(llSpotInx).sSpotType <> "X")) And (blValidVehicle) Then
                    If blFirstTime Then
                        slTempKey = Trim$(tmSdfExtSort(llLoopOnSpots).sKey)
                        ilRet = gParseItem(slTempKey, 1, "|", slTempLine)
                        ilPrevLine = Val(slTempLine)
                        ilRet = gParseItem(slTempKey, 2, "|", slTempVehicle)
                        slPrevKey = Trim$(slTempLine) & "|" & Trim$(slTempVehicle)
                        slCurrKey = slPrevKey
                        ilCurrLine = ilPrevLine
                        ilCurrVefCode = tmSdfExt(llSpotInx).iVefCode
                        ilPrevVefCode = ilCurrVefCode
                        blFirstTime = False
                        'get the line index
                        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                            If ilCurrLine = tgClfCT(ilClf).ClfRec.iLine Then
                                tmClf = tgClfCT(ilClf).ClfRec
                                Exit For
                            End If
                        Next ilClf
                    Else
                        slTempKey = Trim$(tmSdfExtSort(llLoopOnSpots).sKey)
                        ilRet = gParseItem(slTempKey, 1, "|", slTempLine)
                        ilCurrLine = Val(slTempLine)
                        ilRet = gParseItem(slTempKey, 2, "|", slTempVehicle)
                        slCurrKey = Trim$(slTempLine) & "|" & Trim$(slTempVehicle)
                        ilCurrVefCode = tmSdfExt(llSpotInx).iVefCode
                    End If

                    If StrComp(slPrevKey, slCurrKey, vbBinaryCompare) = 0 Then
                        'equal line & vehicle
                        'determine the month this spot goes into, and accumulate the $
                        mGetRateAndAddToArray tmSdfExtSort(llLoopOnSpots).lSdfExtIndex, llStartDates()
                    Else
                        'different line or vehicle, create an output line
                        ilRet = mSplitAndCreate(llStartDates(), tlEfficioINfo, 1, slCashAgyComm, ilPrevVefCode)
                        slPrevKey = slCurrKey
                        ilPrevLine = ilCurrLine
                        ilPrevVefCode = ilCurrVefCode
                        'initialize for next line/vehicle
                        For ilLoop = 1 To MAXMONTHS
                            lmProject(ilLoop) = 0
                            lmAcquisition(ilLoop) = 0
                        Next ilLoop
                        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                            If ilCurrLine = tgClfCT(ilClf).ClfRec.iLine Then
                                tmClf = tgClfCT(ilClf).ClfRec
                                mGetRateAndAddToArray tmSdfExtSort(llLoopOnSpots).lSdfExtIndex, llStartDates()
                                Exit For
                            End If
                        Next ilClf
                    End If
                End If                      'endif SchStatus
            Next llLoopOnSpots
            ilRet = mSplitAndCreate(llStartDates(), tlEfficioINfo, 1, slCashAgyComm, ilPrevVefCode)
            For ilLoop = 1 To MAXMONTHS
                lmProject(ilLoop) = 0
                lmAcquisition(ilLoop) = 0
            Next ilLoop
            Next ilCurrentRecd
  

        Erase tmSdfExtSort, tmSdfExt
        Erase tmChfAdvtExt
    Exit Sub
End Sub
'
'               gGetRateAndAddToArray
'               <input> llSpotInx - index from the sorted array of spots (tmSdfExtSort), that points to the actual spot data array (tmSdfExt)
'               pass the index from the sorted spot array of the spot information
'               to be able to retrieve the spot cost from the flight
'
'
Public Sub mGetRateAndAddToArray(llSpotInx As Long, llStartDates() As Long)
Dim ilLoopOnMonth As Integer
Dim llAirDate As Long
Dim ilRet As Integer
Dim slSpotRate As String
Dim llSpotRate As Long
Dim ilAcqCommPct As Integer
Dim ilAcqLoInx As Integer
Dim ilAcqHiInx As Integer
Dim llAcqNet As Long
Dim llAcqComm As Long
Dim blAcqOK As Boolean

            gUnpackDateLong tmSdfExt(llSpotInx).iDate(0), tmSdfExt(llSpotInx).iDate(1), llAirDate
            If tmSdfExt(llSpotInx).sSpotType = "A" Or tmSdfExt(llSpotInx).sSpotType = "T" Or tmSdfExt(llSpotInx).sSpotType = "Q" Then   'include regular sched spots , PI, DR
                                                                'ignore fills, psa, promos, billboards
                'make up spot info for spot price routine
                tmSdf.iAdfCode = tmSdfExt(llSpotInx).iAdfCode
                tmSdf.iDate(0) = tmSdfExt(llSpotInx).iDate(0)
                tmSdf.iDate(1) = tmSdfExt(llSpotInx).iDate(1)
                tmSdf.iVefCode = tmSdfExt(llSpotInx).iVefCode
                tmSdf.lChfCode = tmSdfExt(llSpotInx).lChfCode
                tmSdf.lCode = tmSdfExt(llSpotInx).lCode
                tmSdf.sSchStatus = tmSdfExt(llSpotInx).sSchStatus
                tmSdf.sSpotType = tmSdfExt(llSpotInx).sSpotType
                tmSdf.sPriceType = tmSdfExt(llSpotInx).sPriceType
                
                For ilLoopOnMonth = 1 To igPeriods Step 1       'loop thru months to find the match
                    If llAirDate >= llStartDates(ilLoopOnMonth) And llAirDate < llStartDates(ilLoopOnMonth + 1) Then
                        'process the spot $
                        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotRate)
                        If (InStr(slSpotRate, ".") <> 0) Then        'found spot cost
                            'is it a .00?
                            If gCompNumberStr(slSpotRate, "0.00") = 0 Then       'its a .00 spot
                                llSpotRate = 0
                            Else
                                llSpotRate = gStrDecToLong(slSpotRate, 2)
                            End If
                        Else
                            'its a bonus, recap, n/c, etc. which is still $0
                            llSpotRate = 0
                        End If
                        lmProject(ilLoopOnMonth) = lmProject(ilLoopOnMonth) + llSpotRate
                        
                        '7/31/15 implement acq commission  if applicable
                        If ExpEfficioRev!rbcNetBy(1).Value Then     't-net?
                            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then    'acq comm are applicable
                                ilAcqCommPct = 0
                                blAcqOK = gGetAcqCommInfoByVehicle(tmSdf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                ilAcqCommPct = gGetEffectiveAcqComm(llAirDate, ilAcqLoInx, ilAcqHiInx)
                                gCalcAcqComm ilAcqCommPct, tmClf.lAcquisitionCost, llAcqNet, llAcqComm
                                lmAcquisition(ilLoopOnMonth) = lmAcquisition(ilLoopOnMonth) + llAcqNet
                            Else                        'no acq comm; both net and gross are the same
                                lmAcquisition(ilLoopOnMonth) = lmAcquisition(ilLoopOnMonth) + tmClf.lAcquisitionCost
                            End If
                        End If
                        'lmAcquisition(ilLoopOnMonth) = lmAcquisition(ilLoopOnMonth) + tmClf.lAcquisitionCost
                        Exit For
                    End If
                If ilLoopOnMonth = 36 Then
                igPeriods = igPeriods
                End If
                Next ilLoopOnMonth
            End If
        Exit Sub
End Sub

Private Sub tmcSetTime_Timer()
    If igExportType = 2 Then                       'manual from exports or manual from traffic
        gUpdateTaskMonitor 0, "EPE"
    ElseIf igExportType = 3 Then                       'manual from exports or manual from traffic
        gUpdateTaskMonitor 0, "ERE"
    End If
End Sub
