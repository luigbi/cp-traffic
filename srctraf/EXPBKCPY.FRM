VERSION 5.00
Begin VB.Form ExpBkCpy 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   240
   ClientTop       =   1140
   ClientWidth     =   9405
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
   ScaleHeight     =   5970
   ScaleWidth      =   9405
   Begin VB.PictureBox plcRotInfo 
      BackColor       =   &H00FFFF80&
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
      Height          =   1230
      Left            =   1455
      ScaleHeight     =   1170
      ScaleWidth      =   6465
      TabIndex        =   24
      Top             =   4455
      Visible         =   0   'False
      Width           =   6525
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Rotation: #  by user name"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Top             =   45
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Entered Date xx/xx/xx    Version Date xx/xx/xx   Modified xx Times"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   28
         Top             =   270
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Date Range Assigned To:  Earliest xx/xx/xx   Latest xx/xx/xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   27
         Top             =   495
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Bulk Feed: Send on xx/xx/xx"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   25
         Top             =   945
         Width           =   6315
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Last Assignment Done: Date xx/xx/xx   Time xx:xx:xxam"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   26
         Top             =   720
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo Export..."
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
      Left            =   7440
      TabIndex        =   32
      Top             =   5640
      Width           =   1260
   End
   Begin VB.Timer tmcRot 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8940
      Top             =   5145
   End
   Begin VB.Timer tmcDDE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8925
      Top             =   4680
   End
   Begin VB.CommandButton cmcReSend 
      Appearance      =   0  'Flat
      Caption         =   "Resen&d Rotation..."
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
      Left            =   5790
      TabIndex        =   12
      Top             =   5640
      Width           =   1605
   End
   Begin VB.CommandButton cmcSuppress 
      Appearance      =   0  'Flat
      Caption         =   "&Suppress Rotation"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   5640
      Width           =   1665
   End
   Begin VB.CommandButton cmcAffCopy 
      Appearance      =   0  'Flat
      Caption         =   "&Affiliate Copy..."
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
      Left            =   2685
      TabIndex        =   10
      Top             =   5640
      Width           =   1350
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1845
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   14
      Top             =   570
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
         TabIndex        =   15
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "EXPBKCPY.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   19
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   21
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.ListBox lbcVehicleCode 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   7815
      Sorted          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8805
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4605
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8790
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4755
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4935
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
      Left            =   495
      TabIndex        =   8
      Top             =   5640
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
      Left            =   1590
      TabIndex        =   9
      Top             =   5640
      Width           =   1050
   End
   Begin VB.PictureBox plcDates 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4200
      Left            =   15
      ScaleHeight     =   4170
      ScaleWidth      =   9330
      TabIndex        =   1
      Top             =   270
      Width           =   9360
      Begin VB.PictureBox pbcLbcRot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   2100
         Left            =   90
         ScaleHeight     =   2100
         ScaleWidth      =   8940
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1830
         Width           =   8940
      End
      Begin VB.ListBox lbcVeh 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   5925
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   60
         Width           =   3360
      End
      Begin VB.TextBox edcTranDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Top             =   90
         Width           =   930
      End
      Begin VB.ListBox lbcRot 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   75
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   1815
         Width           =   8970
      End
      Begin VB.VScrollBar vbcRot 
         Height          =   2145
         LargeChange     =   9
         Left            =   9045
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1800
         Width           =   270
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Rotations"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1695
      End
      Begin VB.CommandButton cmcTranDate 
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
         Left            =   2775
         Picture         =   "EXPBKCPY.frx":2E1A
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   195
      End
      Begin VB.Label lacProcessing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   1140
         Width           =   5715
      End
      Begin VB.Label lacTranDate 
         Appearance      =   0  'Flat
         Caption         =   "Transmission Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   1725
      End
   End
   Begin VB.PictureBox pbcLbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   825
      Left            =   660
      ScaleHeight     =   825
      ScaleWidth      =   7800
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4545
      Width           =   7800
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   645
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   4515
      Width           =   8130
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   4545
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ExpBkCpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPBKCPY.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpBkCpy.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the agency conversion input screen code
'
'Note:
'   tmVef and tmVpfInfo are matched and contain the following vehicles:
'   Conventional without Bulk Groups
'   Conventional with Bulk Groups (first vehicle of group alphabetically, remaining vehicles in group are in tmLkVehInfo)
'   Airing without Bulk Groups
'   Airing with Bulk Groups (first vehicle of group alphabetically, remaining vehicles in group are in tmLkVehInfo)
'   Selling
'
'   Only conventional and airing vehicles have groups.
'   Bulk Groups obtained from vehicle options interface area
'   tmVef and tmVpfInfo build in mVehPop
'   Populate lbcVehicle from Airing and Conventional from tmVef
'   Populate lbvVeh by call to gPopUserVehicleBox (with conventional; selling and airing vehicles) in mInit
'
Option Explicit
Option Compare Text
Dim imListFieldVeh(0 To 3) As Integer
Dim imListFieldRot(0 To 12) As Integer
Dim imFirstActivate As Integer
Dim tmVehCode() As SORTCODE
Dim smVehCodeTag As String
Dim imVefCode() As Integer
Dim imDVefCode() As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmExport As Integer   'file hanle
Dim hmMsg As Integer        'Message File Handle
Dim smFileNames() As String 'Use if abort issued
'Contract header
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
'Rotation header
Dim tmCrfSrchKey As LONGKEY0  'CRF key record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
Dim tmSTCrf() As CRF        'Supersede test
Dim imSelectPrevState() As Integer
Dim lmRotCodeBuild() As Long 'Rotation codes build for export- aviod sending same twicw
'Short Title
Dim tmSif As SIF            'SIF record image
Dim hmSif As Integer        'SIF Handle
Dim imSifRecLen As Integer      'SIF record length
'Short Title via Contract
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0  'VSF key record image
Dim hmVsf As Integer        'VSF Handle
Dim imVsfRecLen As Integer      'VSF record length
'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length
Dim tmCnfRot() As CNF
'Media Code
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0  'MCF key record image
Dim hmMcf As Integer        'MCF Handle
Dim imMcfRecLen As Integer      'MCF record length
'Inventory
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0  'CIF key record image
Dim hmCif As Integer        'CIF Handle
Dim imCifRecLen As Integer      'CIF record length
'Copy Feed
Dim tmCyf As CYF            'CYF record image
Dim tmCyfSrchKey As CYFKEY0  'CYF key record image
Dim hmCyf As Integer        'CYF Handle
Dim imCyfRecLen As Integer      'CYF record length
Dim tmCyfTest() As CYFTEST
'Product
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim hmCpf As Integer        'CPF Handle
Dim imCpfRecLen As Integer      'CPF record length
'Comment
Dim tmCsf As CSF            'CSF record image
Dim tmCsfSrchKey As LONGKEY0  'CSF key record image
Dim hmCsf As Integer        'CSF Handle
Dim imCsfRecLen As Integer      'CSF record length
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim imAdfRecLen As Integer  'ADF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ADF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Vehicle
Dim hmVef As Integer
Dim tmVef() As VEF
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim imNoRotations As Integer
Dim tmVpfInfo() As VPFINFO
Dim tmLkVehInfo() As LKVEHINFO
Dim tmSALink() As LKVEHINFO
'Vehicle link
Dim tmVlf As VLF            'VLF record image
Dim hmVlf As Integer        'VLF Handle
Dim imVlfRecLen As Integer      'VLF record length
'Dim tmRec As LPOPREC
'Export copy as ShortTitle(ISCI).mp2
Dim lmSentCifCode() As Long
Dim hmNewCIF As Integer 'Handle for New only
Dim hmAllCIF As Integer 'Handle for All

Dim imIgnoreCkcAll As Integer
Dim imLastIndex As Integer
Dim imCurrentIndex As Integer
Dim imShiftKey As Integer   'Bit 0=Shift; 1=Ctrl; 2=Alt
Dim imButton As Integer
Dim imButtonIndex As Integer
Dim imIgnoreRightMove As Integer
Dim imIgnoreVbcChg As Integer
Dim imTypeIndex As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smTodaysDate As String  'mm/dd/yy"
Dim tmVehTimes() As VEHTIMES
Dim tmRotInfo() As SENDROTINFO
Dim tmAddCyf() As SENDCOPYINFO
'Btrieve wait
Dim imWaitCount As Integer
Dim imTimeDelay As Integer
Dim imLockValue As Integer
Dim imTranLog As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=End Date; 3=Transmission Date
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slTime As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    If imIgnoreCkcAll Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    If lbcRot.ListCount <= 0 Then
        If Value Then
            'ckcAll.Value = False
            ckcAll.Value = vbUnchecked
        End If
        imIgnoreVbcChg = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = Value
    If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
        llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
    Else
        llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
    End If
    llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
    DoEvents
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        tgSortCrf(ilLoop).iSelected = Value
    Next ilLoop
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = lbcVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "|", slTime)
        slTime = "0:0"
        lbcVehicle.List(ilLoop) = slName & "|" & slTime
    Next ilLoop
    If Value Then
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tmCrf = tgSortCrf(ilLoop).tCrf
            mComputeTime True
            ilIndex = tgSortCrf(ilLoop).iCombineIndex
            Do While ilIndex >= 0
                tmCrf = tgCombineCrf(ilIndex).tCrf
                mComputeTime True
                ilIndex = tgCombineCrf(ilIndex).iCombineIndex
            Loop
            ilIndex = tgSortCrf(ilLoop).iDuplIndex
            Do While ilIndex >= 0
                tmCrf = tgDuplCrf(ilIndex).tCrf
                mComputeTime True
                ilIndex = tgDuplCrf(ilIndex).iDuplIndex
            Loop
            'DoEvents
        Next ilLoop
    End If
    pbcLbcRot_Paint
    pbclbcVehicle_Paint
    mSetCommands
    Screen.MousePointer = vbDefault
    imIgnoreVbcChg = False
End Sub
Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcAffCopy_Click()
    AffCopy.Show vbModal
End Sub
Private Sub cmcCalDn_Click()
    If imDateBox = 1 Then
        'imCalMonth = imCalMonth - 1
        'If imCalMonth <= 0 Then
        '    imCalMonth = 12
        '    imCalYear = imCalYear - 1
        'End If
        'pbcCalendar_Paint
        'edcStartDate.SelStart = 0
        'edcStartDate.SelLength = Len(edcStartDate.Text)
        'edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        'imCalMonth = imCalMonth - 1
        'If imCalMonth <= 0 Then
        '    imCalMonth = 12
        '    imCalYear = imCalYear - 1
        'End If
        'pbcCalendar_Paint
        'edcEndDate.SelStart = 0
        'edcEndDate.SelLength = Len(edcEndDate.Text)
        'edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
        edcTranDate.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    If imDateBox = 1 Then
        'imCalMonth = imCalMonth + 1
        'If imCalMonth > 12 Then
        '    imCalMonth = 1
        '    imCalYear = imCalYear + 1
        'End If
        'pbcCalendar_Paint
        'edcStartDate.SelStart = 0
        'edcStartDate.SelLength = Len(edcStartDate.Text)
        'edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        'imCalMonth = imCalMonth + 1
        'If imCalMonth > 12 Then
        '    imCalMonth = 1
        '    imCalYear = imCalYear + 1
        'End If
        'pbcCalendar_Paint
        'edcEndDate.SelStart = 0
        'edcEndDate.SelLength = Len(edcEndDate.Text)
        'edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
        edcTranDate.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_Click()
    Dim slExportFile As String
    Dim slMsgFile As String
    Dim slMsgFileName As String
    Dim ilMsgType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slTranDate As String
    Dim slDate As String
    Dim slTime As String
    Dim slNameTime As String
    Dim slVehName As String
    Dim slStr As String
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVefSelected As Integer
    Dim ilPos As Integer
    Dim ilPageNo As Integer
    Dim ilOldPageNo As Integer
    Dim ilLineNo As Integer
    Dim ilCopyNo As Integer
    Dim slRotStartDate As String
    Dim slRotEndDate As String
    Dim slRecord As String
    Dim ilTotalTime As Integer
    Dim slSec As String
    Dim ilUpper As Integer
    Dim ilDay As Integer
    Dim ilDayReq As Integer
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim llDate As Long
    Dim ilNewInv As Integer
    Dim ilCheckInv As Integer
    Dim slComment As String
    Dim slKey As String
    Dim slProduct As String
    Dim slCreative As String
    Dim slISCI As String
    Dim slProdISCITitle As String
    Dim slPrevProdISCITitle As String
    Dim slPrevVehName As String
    Dim ilShowName As Integer
    Dim slNowDate As String
    Dim slStnCode As String
    Dim slFeedNo As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slLine As String
    Dim slName As String
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilTotalNoInstr As Integer
    Dim ilLen As Integer
    Dim slPrevNewInv As String
    Dim slBlank As String
    Dim slCopyHeader As String
    Dim ilNoCopyLines As Integer
    Dim ilAnyTransmitted As Integer
    Dim ilCombineIndex As Integer
    Dim ilDuplIndex As Integer
    Dim ilCrfIndex As Integer
    Dim ilLoop1 As Integer
    Dim ilVIndex As Integer
    Dim ilLastInvIndex As Integer
    Dim ilShowInv As Integer
    Dim slCart As String
    Dim ilShowCart As Integer
    Dim ilCombineOk As Integer
    Dim ilCheck As Integer
    Dim ilVpfIndex As Integer
    Dim tlCyf As CYF
    ReDim ilDayOn(0 To 6) As Integer
    Dim ilBeginMsgSent As Integer
    Dim slMsgLine As String
    Dim slTimeStamp As String
    Dim ilTranDate0 As Integer
    Dim ilTranDate1 As Integer
    Dim slXRefLetter As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVsf As Integer
    Dim llSifCode As Long
    Dim ilSAGroupNo As Integer
    Dim slShortTitle As String
    Dim slDateTime As String

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    slStr = edcTranDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcTranDate.SetFocus
        Exit Sub
    End If
    gPackDate slStr, ilTranDate0, ilTranDate1
    Screen.MousePointer = vbHourglass
    
    sgMessageFile = sgDBPath & "Messages\" & "ExpBulkCopy.Txt"
    
    gAutomationAlertAndLogHandler "** Export Bulk Copy **"
    gAutomationAlertAndLogHandler "* TransDate = " & edcTranDate.Text
    If ckcAll.Value = vbChecked Then
        gAutomationAlertAndLogHandler "* All Rotations = True"
    Else
        gAutomationAlertAndLogHandler "* All Rotations = False"
    End If
    
    gAutomationAlertAndLogHandler "Building export table..."
    mBuildExpTable
    ilUpper = UBound(tmAddCyf)
    If ilUpper > 0 Then
        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, LenB, 0
        ArraySortTyp fnAV(tmAddCyf(), 0), ilUpper, 0, LenB(tmAddCyf(0)), 0, LenB(tmAddCyf(0).sKey), 0
    End If
    ilUpper = UBound(tmRotInfo)
    If ilUpper > 0 Then
        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, LenB, 0
        ArraySortTyp fnAV(tmRotInfo(), 0), ilUpper, 0, LenB(tmRotInfo(0)), 0, LenB(tmRotInfo(0).sKey), 0
    End If
    slBlank = ""
    slStr = edcTranDate.Text
    slTranDate = Format$(gDateValue(slStr), "mm/dd/yyyy")
    slNowDate = Format$(Now, "m/d/yy")
    'gObtainYearMonthDayStr slNowDate, True, slFYear, slFMonth, slFDay
    gObtainYearMonthDayStr slTranDate, True, slFYear, slFMonth, slFDay
    slFeedNo = right$(slFYear, 2) & slFMonth & slFDay
    'Determine if file exist- if so don't allow export
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameTime = lbcVehicle.List(ilVeh) 'Airing and conventional vehicles (with and without bulk groups)
        ilPos = InStr(slNameTime, "|")
        slVehName = Left$(slNameTime, ilPos - 1)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            ilLen = Len(Trim$(tmVef(ilLoop).sName))
            If (Trim$(tmVef(ilLoop).sName) = Left$(slVehName, ilLen)) Then
                ilVehIndex = ilLoop
                ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tmVef(ilLoop).iCode)).iSAGroupNo
                Exit For
            End If
        Next ilLoop
        'Test if vehicle has a rotation to be transmitted
        ilVefSelected = False
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            If lbcVeh.Selected(ilLoop) Then
                slNameCode = tmVehCode(ilLoop).sKey    'Selling and conventional vehicles 'lbcVehCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                If tmVef(ilVehIndex).sType = "C" Then
                    If tmVef(ilVehIndex).iCode = Val(slCode) Then
                        ilVefSelected = True
                        Exit For
                    End If
                    ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            If tmLkVehInfo(ilVIndex).iVefCode = tmVef(ilVehIndex).iCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                Else
                    'ilVpfIndex = gVpfFind(ExpBkCpy, Val(slCode))
                    'If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
                    '    ilVefSelected = True
                    '    Exit For
                    'End If
                    ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                    'If ilVIndex >= 0 Then
                    '    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    '    Do While ilVIndex > 0
                    '        ilVpfIndex = gVpfFind(ExpBkCpy, tmLkVehInfo(ilVIndex).iVefCode)
                    '        If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
                    '            ilVefSelected = True
                    '            Exit Do
                    '        End If
                    '        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    '    Loop
                    'End If
                    'For ilLoop2 = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
                    '    If tgVpf(ilVpfIndex).iGLink(ilLoop2) > 0 Then
                    '        If tgVpf(ilVpfIndex).iGLink(ilLoop2) = tmVef(ilVehIndex).iCode Then
                    '            ilVefSelected = True
                    '            Exit For
                    '        End If
                    '    End If
                    'Next ilLoop2
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                        Do While ilVIndex >= 0
                            If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop
        If ilVefSelected Then
            slStnCode = Trim$(tmVef(ilVehIndex).sCodeStn)
            slExportFile = sgExportPath & slStnCode & slFeedNo & ".msg"
            ilRet = 0
            'On Error GoTo cmcExportErr:
            'slTimeStamp = FileDateTime(slExportFile)
            ilRet = gFileExist(slExportFile)
            If ilRet = 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Bulk Feed already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                gAutomationAlertAndLogHandler "Bulk Feed already generated for this date, Export terminated", vbOkOnly + vbCritical + vbApplicationModal, "Export"
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    Next ilVeh
    
    gAutomationAlertAndLogHandler "Exporting files.."
    slStnCode = "X"
    slXRefLetter = "A"
    Do
        slExportFile = sgExportPath & slStnCode & slFeedNo & slXRefLetter & ".asc"
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slTimeStamp = FileDateTime(slExportFile)
        ilRet = gFileExist(slExportFile)
        If ilRet = 0 Then
            slXRefLetter = Chr$(Asc(slXRefLetter) + 1)
        End If
    Loop While ilRet = 0    'equal zero if file exist
    Screen.MousePointer = vbHourglass
    ReDim smFileNames(0 To 0) As String 'Use if abort issued
    For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameTime = lbcVehicle.List(ilVeh)
        ilPos = InStr(slNameTime, "|")
        slVehName = Left$(slNameTime, ilPos - 1)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            ilLen = Len(Trim$(tmVef(ilLoop).sName))
            If (Trim$(tmVef(ilLoop).sName) = Left$(slVehName, ilLen)) Then
                ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tmVef(ilLoop).iCode)).iSAGroupNo
                ilVehIndex = ilLoop
                Exit For
            End If
        Next ilLoop
        'Test if vehicle has a rotation to be transmitted
        ilVefSelected = False
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            If lbcVeh.Selected(ilLoop) Then
                slNameCode = tmVehCode(ilLoop).sKey    'lbcVehCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                ilVpfIndex = gVpfFind(ExpBkCpy, Val(slCode))
                'If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
                '    ilShowCart = True
                'Else
                '    ilShowCart = False
                'End If
                If tmVef(ilVehIndex).sType = "C" Then
                    If tmVef(ilVehIndex).iCode = Val(slCode) Then
                        ilVefSelected = True
                        Exit For
                    End If
                    ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            If tmLkVehInfo(ilVIndex).iVefCode = tmVef(ilVehIndex).iCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
               Else
                    'ilVpfIndex = gVpfFind(ExpBkCpy, Val(slCode))
                    'If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
                    '    ilVefSelected = True
                    '    Exit For
                    'End If
                    ilVIndex = mFindVpfIndex(Val(slCode)) 'tmVef(ilVehIndex).iCode)
                    'If ilVIndex >= 0 Then
                    '    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    '    Do While ilVIndex > 0
                    '        ilVpfIndex = gVpfFind(ExpBkCpy, tmLkVehInfo(ilVIndex).iVefCode)
                    '        If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
                    '            ilVefSelected = True
                    '            Exit Do
                    '        End If
                    '        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    '    Loop
                    'End If
                    'For ilLoop2 = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
                    '    If tgVpf(ilVpfIndex).iGLink(ilLoop2) > 0 Then
                    '        If tgVpf(ilVpfIndex).iGLink(ilLoop2) = tmVef(ilVehIndex).iCode Then
                    '            ilVefSelected = True
                    '            Exit For
                    '        End If
                    '    End If
                    'Next ilLoop2
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                        Do While ilVIndex >= 0
                            If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
                                ilVefSelected = True
                                Exit Do
                            End If
                            ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop
        If ilVefSelected Then
            ReDim lmSentCifCode(0 To 0) As Long
            ilVpfIndex = gVpfFind(ExpBkCpy, tmVef(ilVehIndex).iCode)
            If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
                ilShowCart = True
            Else
                ilShowCart = False
                If (tmVef(ilVehIndex).sType = "A") Or (tmVef(ilVehIndex).sType = "C") Then
                    ilVIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    If ilVIndex >= 0 Then
                        'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                        '    If tmVpfInfo(ilVIndex).iVefLink(ilLoop2) > 0 Then
                        '        ilVpfIndex = gVpfFind(ExpBkCpy, tmVpfInfo(ilVIndex).iVefLink(ilLoop2))
                        '        If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
                        '            ilShowCart = True
                        '            Exit For
                        '        End If
                        '    End If
                        'Next ilLoop2
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            ilVpfIndex = gVpfFind(ExpBkCpy, tmLkVehInfo(ilVIndex).iVefCode)
                            If tgVpf(ilVpfIndex).sExpBkCpyCart = "Y" Then
                                ilShowCart = True
                                Exit Do
                            End If
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
            gAutomationAlertAndLogHandler "Processing: " & slVehName
            lacProcessing.Caption = "Processing: " & slVehName
            DoEvents
            'slFileName = ""
            'For ilLoop = 1 To Len(slVehName) Step 1
            '    slChar = Mid$(slVehName, ilLoop, 1)
            '    If slChar <> " " Then
            '        slFileName = slFileName & slChar
            '        If Len(slFileName) = 8 Then
            '            Exit For
            '        End If
            '    End If
            'Next ilLoop
            slStnCode = Trim$(tmVef(ilVehIndex).sCodeStn)
            slExportFile = sgExportPath & slStnCode & slFeedNo & ".msg"
            ilRet = 0
            'On Error GoTo cmcExportErr:
            'hmExport = FreeFile
            ''Create file name based on vehicle name
            'Open slExportFile For Output As hmExport
            ilRet = gFileOpen(slExportFile, "Output", hmExport)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
            smFileNames(UBound(smFileNames)) = slExportFile
            ReDim Preserve smFileNames(0 To UBound(smFileNames) + 1) As String 'Use if abort issued

            ''Create eBulk file names, ABC request 11/17/03
            'On Error GoTo cmcExportErr
            ilRet = 0
            slStr = sgExportPath & slStnCode & slFeedNo & "_New.bul"
            'slDateTime = FileDateTime(slStr)
            ilRet = gFileExist(slStr)
            If ilRet = 0 Then
                Kill slStr
            End If
            ilRet = 0
            'hmNewCIF = FreeFile
            'Open slStr For Output As hmNewCIF
            ilRet = gFileOpen(slStr, "Output", hmNewCIF)
            If ilRet <> 0 Then
                Close #hmNewCIF
                hmNewCIF = -1
            End If
            ilRet = 0
            slStr = sgExportPath & slStnCode & slFeedNo & "_All.bul"
            'slDateTime = FileDateTime(slStr)
            ilRet = gFileExist(slStr)
            If ilRet = 0 Then
                Kill slStr
            End If
            ilRet = 0
            'hmAllCIF = FreeFile
            'Open slStr For Output As hmAllCIF
            ilRet = gFileOpen(slStr, "Output", hmAllCIF)
            If ilRet <> 0 Then
                Close #hmAllCIF
                hmAllCIF = -1
            End If

            'Output new inventory
            ilPageNo = 0
            ilCopyNo = 1
            ilLineNo = 52
            ilTotalTime = 0
            slPrevNewInv = ""
            ilLastInvIndex = -1
            ilAnyTransmitted = False
            ilBeginMsgSent = False
            For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
                If tmAddCyf(ilLoop).tCyf.iVefCode = tmVef(ilVehIndex).iCode Then
                    If Not ilBeginMsgSent Then    'Add Start Page Note-
                        ilNoCopyLines = 1
                        '6/2/16: Replaced GoSub
                        'GoSub cmcExportCopyHeader
                        mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                        ilMsgType = 0
                        slMsgFileName = slStnCode & "000000.Beg"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                        slMsgFileName = slStnCode & slFeedNo & ".Beg"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                        slMsgFileName = "00" & slFeedNo & ".Beg"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                        slMsgFileName = "00000000.Beg"
                        '6/3/16: Replaced GoSub
                        'GoSub cmcExportSendMsg
                        mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                        ilBeginMsgSent = True
                    End If
                    'Bypass any inventory that has been shown previously but is
                    'only for a different zone
                    If ilLastInvIndex <> -1 Then
                        If (tmAddCyf(ilLastInvIndex).tCyf.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode) And (tmAddCyf(ilLastInvIndex).tCyf.iVefCode = tmAddCyf(ilLoop).tCyf.iVefCode) Then
                            ilShowInv = False
                        Else
                            ilShowInv = True
                        End If
                    Else
                        ilShowInv = True
                    End If
                    If ilShowInv Then
                        ilLastInvIndex = ilLoop
                        tmCifSrchKey.lCode = tmAddCyf(ilLoop).tCyf.lCifCode
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmCpf.sName = ""
                                tmCpf.sISCI = ""
                                tmCpf.sCreative = ""
                            End If
                            ilAnyTransmitted = True
                            slRotStartDate = Format$(tmAddCyf(ilLoop).lRotStartDate, "mm/dd/yyyy")
                            slRotEndDate = Format$(tmAddCyf(ilLoop).lRotEndDate, "mm/dd/yyyy")
                            slRecord = "     "
                            'If (Trim$(tmCpf.sName) <> "") Then
                            If (Trim$(tmCpf.sName) <> "") And (tgSpf.sUseProdSptScr <> "P") Then
                                slRecord = slRecord & UCase(Trim$(tmCpf.sName))
                            Else
                                slRecord = slRecord & UCase(Trim$(tmAddCyf(ilLoop).sChfProduct))
                            End If
                            slRecord = slRecord & "    "
                            slRecord = slRecord & "(St. " & slRotStartDate & " - " & slRotEndDate & ")"
                            If StrComp(Trim$(slRecord), Trim$(slPrevNewInv), 1) <> 0 Then
                                slPrevNewInv = slRecord
                                ilNoCopyLines = 5   'Make sure enough room for advertiser and inventory
                                '6/2/16: Replaced GoSub
                                'GoSub cmcExportCopyHeader
                                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                                ilNoCopyLines = 1
                                If Len(slPrevNewInv) <> 0 Then
                                    'slBlank = ""  'Blank line
                                    If Not mExportLine(slBlank, ilLineNo) Then
                                        Exit Sub
                                    End If
                                    '6/2/16: Replaced GoSub
                                    'GoSub cmcExportCopyHeader
                                    mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                                    'slBlank = ""  'Blank line
                                    If Not mExportLine(slBlank, ilLineNo) Then
                                        Exit Sub
                                    End If
                                End If
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replaced GoSub
                                'GoSub cmcExportCopyHeader
                                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                                slRecord = "     "
                                slRecord = slRecord & "------------------------------"
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replaced GoSub
                                'GoSub cmcExportCopyHeader
                                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                            Else
                                ilNoCopyLines = 1   'Make sure enough room for inventory
                                '6/2/16: Replaced GoSub
                                'GoSub cmcExportCopyHeader
                                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                            End If
                            ilNoCopyLines = 1   'Make sure enough room for inventory
                            slRecord = Trim$(str$(ilCopyNo))
                            Do While Len(slRecord) < 2
                                slRecord = " " & slRecord
                            Loop
                            ilCopyNo = ilCopyNo + 1
                            Do While Len(slRecord) < 5
                                slRecord = slRecord & " "
                            Loop
                            slRecord = slRecord & Trim$(tmCpf.sISCI)
                            Do While Len(slRecord) < 27
                                slRecord = slRecord & " "
                            Loop
                            slRecord = slRecord & UCase(Trim$(tmCpf.sCreative))
                            Do While Len(slRecord) < 71
                                slRecord = slRecord & " "
                            Loop
                            ilTotalTime = ilTotalTime + tmCif.iLen
                            slRecord = slRecord & Trim$(str$(tmCif.iLen))
                            If Not mExportLine(slRecord, ilLineNo) Then
                                Exit Sub
                            End If
                            'GoSub cmcExportCopyHeader
                            'slRecord = ""  'Blank line
                            'If Not mExportLine(slRecord, ilLineNo) Then
                            '    Exit Sub
                            'End If
                            'GoSub cmcExportCopyHeader
                            'slRecord = ""  'Blank line
                            'If Not mExportLine(slRecord, ilLineNo) Then
                            '    Exit Sub
                            'End If
                        End If
                    End If
                End If
            Next ilLoop
            slPrevNewInv = ""
            If Not ilAnyTransmitted Then
                If Not ilBeginMsgSent Then    'Add Start Page Note-
                    ilNoCopyLines = 1
                    '6/2/16: Replaced GoSub
                    'GoSub cmcExportCopyHeader
                    mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                    ilMsgType = 0
                    slMsgFileName = slStnCode & "000000.Beg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                    slMsgFileName = slStnCode & slFeedNo & ".Beg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                    slMsgFileName = "00" & slFeedNo & ".Beg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                    slMsgFileName = "00000000.Beg"
                    '6/3/16: Replaced GoSub
                    'GoSub cmcExportSendMsg
                    mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
                    ilBeginMsgSent = True
                End If
                ilNoCopyLines = 1   'Make sure enough room for advertiser and inventory
                '6/2/16: Replaced GoSub
                'GoSub cmcExportCopyHeader
                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
                slRecord = "*** No Commercials Being Fed Today ***"
                If Not mExportLine(slRecord, ilLineNo) Then
                    Exit Sub
                End If
                If Not mExportLine(slBlank, ilLineNo) Then
                    Exit Sub
                End If
            End If
            '6/2/16: Replaced GoSub
            'GoSub cmcExportCopyHeader
            mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            'slBlank = ""  'Blank line
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replaced GoSub
            'GoSub cmcExportCopyHeader
            mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replaced GoSub
            'GoSub cmcExportCopyHeader
            mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replaced GoSub
            'GoSub cmcExportCopyHeader
            mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            'slBlank = ""  'Blank line
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replaced GoSub
            'GoSub cmcExportCopyHeader
            mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            slRecord = " "
            Do While Len(slRecord) < 37
                slRecord = slRecord & " "
            Loop
            slRecord = slRecord & "TOTAL RUN TIME IN MINUTES:"
            slSec = Trim$(str$((100 * (ilTotalTime Mod 60)) / 60))
            Do While Len(slSec) < 2
                slSec = slSec & "0"
            Loop
            slStr = Trim$(str$(ilTotalTime \ 60)) & "." & slSec
            Do While Len(slStr) < 10
                slStr = " " & slStr
            Loop
            slRecord = slRecord & slStr
            If Not mExportLine(slRecord, ilLineNo) Then
                Exit Sub
            End If
            'Output rotation instructions
            ilLineNo = 52   'Force new page
            For ilLoop = 0 To UBound(tmRotInfo) - 1 Step 1
                If tmRotInfo(ilLoop).iStatus = 1 Then
                    If tmRotInfo(ilLoop).iVefCode = tmVef(ilVehIndex).iCode Then
                        tmCrfSrchKey.lCode = tmRotInfo(ilLoop).lCrfCode
                        ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilAnyTransmitted = True
                            tmChfSrchKey.lCode = tmCrf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                ilLineNo = 52   'Force new page with each rotation
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                slRecord = " "
                                Do While Len(slRecord) < 24
                                    slRecord = slRecord & " "
                                Loop
                                If tmAdf.iCode <> tmChf.iAdfCode Then
                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                End If
                                llSifCode = 0
                                If tmChf.lVefCode < 0 Then
                                    tmVsfSrchKey.lCode = -tmChf.lVefCode
                                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Do While ilRet = BTRV_ERR_NONE
                                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                            If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                                                If tmVsf.lFSComm(ilVsf) > 0 Then
                                                    llSifCode = tmVsf.lFSComm(ilVsf)
                                                End If
                                                Exit For
                                            End If
                                        Next ilVsf
                                        If llSifCode <> 0 Then
                                            Exit Do
                                        End If
                                        If tmVsf.lLkVsfCode <= 0 Then
                                            Exit Do
                                        End If
                                        tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                End If
                                slShortTitle = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
                                slRecord = slRecord & slShortTitle    'Trim$(tmChf.sProduct)
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""  'Blank line
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                slRecord = " "
                                Do While Len(slRecord) < 24
                                    slRecord = slRecord & " "
                                Loop
                                gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                                llRotStartDate = gDateValue(slDate)
                                slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                slRecord = slRecord & "Start Date: " & slDate
                                Do While Len(slRecord) < 50
                                    slRecord = slRecord & " "
                                Loop
                                gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                                llRotEndDate = gDateValue(slDate)
                                slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                slRecord = slRecord & "End Date: " & slDate
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                slRecord = "Please air using the flight data and rotation indicated below:"
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                If tmRotInfo(ilLoop).iRevised Then
                                    slRecord = "        REVISED INSTRUCTIONS!!"
                                    If Not mExportLine(slRecord, ilLineNo) Then
                                        Exit Sub
                                    End If
                                    '6/2/16: Replace GoSub
                                    'GoSub cmcExportRotHeader
                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                End If
                                If (tmCrf.iStartTime(0) <> 0) Or (tmCrf.iStartTime(1) <> 0) Or (tmCrf.iEndTime(0) <> 0) Or (tmCrf.iEndTime(1) <> 0) Then
                                    slRecord = "      AIR THIS COPY BETWEEN "
                                    gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
                                    slTime = UCase(slTime)
                                    If slTime = "12AM" Then
                                        slTime = "12M"
                                    ElseIf slTime = "12PM" Then
                                        slTime = "12N"
                                    End If
                                    'If slTime = "12M" Then
                                    '    slTime = "12AM"
                                    'ElseIf slTime = "12N" Then
                                    '    slTime = "12PM"
                                    'End If
                                    slRecord = slRecord & slTime & " AND "
                                    gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
                                    slTime = UCase(slTime)
                                    If slTime = "12AM" Then
                                        slTime = "12M"
                                    ElseIf slTime = "12PM" Then
                                        slTime = "12N"
                                    End If
                                    'If slTime = "12M" Then
                                    '    slTime = "12AM"
                                    'ElseIf slTime = "12N" Then
                                    '    slTime = "12PM"
                                    'End If
                                    slRecord = slRecord & slTime
                                    If Not mExportLine(slRecord, ilLineNo) Then
                                        Exit Sub
                                    End If
                                    '6/2/16: Replace GoSub
                                    'GoSub cmcExportRotHeader
                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                End If
                                ilDayReq = False
                                For ilDay = 0 To 6 Step 1
                                    ilDayOn(ilDay) = False
                                Next ilDay
                                If llRotEndDate - llRotStartDate >= 6 Then
                                    For ilDay = 0 To 6 Step 1
                                        If tmCrf.sDay(ilDay) <> "Y" Then
                                            ilDayReq = True
                                            ilDayOn(ilDay) = False
                                        Else
                                            ilDayOn(ilDay) = True
                                        End If
                                    Next ilDay
                                Else
                                    For llDate = llRotStartDate To llRotEndDate Step 1
                                        ilDay = gWeekDayLong(llDate)
                                        If tmCrf.sDay(ilDay) <> "Y" Then
                                            'ilDayReq = True
                                            ilDayOn(ilDay) = False
                                        Else
                                            ilDayOn(ilDay) = True
                                        End If
                                    Next llDate
                                    ilDayReq = True
                                End If
                                If ilDayReq Then
                                    slStr = ""
                                    If (ilDayOn(0) = True) And (ilDayOn(1) = True) And (ilDayOn(2) = True) And (ilDayOn(3) = True) And (ilDayOn(4) = True) And (ilDayOn(5) = False) And (ilDayOn(6) = False) Then
                                        slStr = " Mo Thru Fr"
                                    ElseIf (ilDayOn(0) = False) And (ilDayOn(1) = False) And (ilDayOn(2) = False) And (ilDayOn(3) = False) And (ilDayOn(4) = False) And (ilDayOn(5) = True) And (ilDayOn(6) = True) Then
                                        slStr = " THE WEEKENDS"
                                    Else
                                        For ilDay = 0 To 6 Step 1
                                            If ilDayOn(ilDay) = True Then
                                                Select Case ilDay
                                                    Case 0
                                                        slStr = slStr & " Mo"
                                                    Case 1
                                                        slStr = slStr & " Tu"
                                                    Case 2
                                                        slStr = slStr & " We"
                                                    Case 3
                                                        slStr = slStr & " Th"
                                                    Case 4
                                                        slStr = slStr & " Fr"
                                                    Case 5
                                                        slStr = slStr & " Sa"
                                                    Case 6
                                                        slStr = slStr & " Su"
                                                End Select
                                            End If
                                        Next ilDay
                                    End If
                                    slRecord = "      AIR THIS COPY ON" & slStr
                                    If Not mExportLine(slRecord, ilLineNo) Then
                                        Exit Sub
                                    End If
                                    '6/2/16: Replace GoSub
                                    'GoSub cmcExportRotHeader
                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                End If
                                If Trim$(tmCrf.sZone) <> "" Then
                                    Select Case tmCrf.sZone
                                        Case "EST"
                                            slStr = "EASTERN TIME ZONE"
                                        Case "CST"
                                            slStr = "CENTRAL TIME ZONE"
                                        Case "MST"
                                            slStr = "MOUNTAIN TIME ZONE"
                                        Case "PST"
                                            slStr = "PACIFIC TIME ZONE"
                                    End Select
                                    slRecord = "      AIR THIS COPY ONLY IF YOU ARE IN THE " & slStr
                                    If Not mExportLine(slRecord, ilLineNo) Then
                                        Exit Sub
                                    End If
                                    '6/2/16: Replace GoSub
                                    'GoSub cmcExportRotHeader
                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                End If
                                If tmCrf.lCsfCode <> 0 Then
                                    tmCsfSrchKey.lCode = tmCrf.lCsfCode
                                    tmCsf.sComment = ""
                                    imCsfRecLen = Len(tmCsf) '5011
                                    ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        'Output 70 characters per line
                                        slComment = gStripChr0(tmCsf.sComment)
                                        'If tmCsf.iStrLen > 0 Then
                                        If slComment <> "" Then
                                            'slComment = Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                                            slRecord = "     "
                                            Do While Len(slComment) > 0
                                                ilPos = InStr(slComment, " ")
                                                If ilPos > 0 Then
                                                    If Len(slRecord) + ilPos - 1 > 70 Then
                                                        If Not mExportLine(slRecord, ilLineNo) Then
                                                            Exit Sub
                                                        End If
                                                        '6/2/16: Replace GoSub
                                                        'GoSub cmcExportRotHeader
                                                        mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                        slRecord = "     "
                                                    End If
                                                    slRecord = slRecord & Left$(slComment, ilPos)
                                                    slComment = right$(slComment, Len(slComment) - ilPos)
                                                Else
                                                    If Len(slRecord) + Len(slComment) > 70 Then
                                                        If Not mExportLine(slRecord, ilLineNo) Then
                                                            Exit Sub
                                                        End If
                                                        '6/2/16: Replace GoSub
                                                        'GoSub cmcExportRotHeader
                                                        mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                        slRecord = "     "
                                                    End If
                                                    slRecord = slRecord & slComment
                                                    If Not mExportLine(slRecord, ilLineNo) Then
                                                        Exit Sub
                                                    End If
                                                    '6/2/16: Replace GoSub
                                                    'GoSub cmcExportRotHeader
                                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                    slComment = ""
                                                    Exit Do
                                                End If
                                            Loop
                                        End If
                                    End If
                                End If
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                slRecord = "COMM. NUMBER         COMMERCIAL NAME               LENGTH ROTATION" ' PRODUCT"
                                If Not mExportLine(slRecord, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                'slRecord = ""
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                If Not mExportLine(slBlank, ilLineNo) Then
                                    Exit Sub
                                End If
                                '6/2/16: Replace GoSub
                                'GoSub cmcExportRotHeader
                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                ReDim tmCnfRot(0 To 0) As CNF
                                ilCombineIndex = -1
                                Do
                                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                                    tmCnfSrchKey.iInstrNo = 0
                                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                                        ilFound = False
                                        For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                            If tmCnf.lCifCode = tmCnfRot(ilTest).lCifCode Then
                                                ilFound = True
                                                tmCnfRot(ilTest).iInstrNo = tmCnfRot(ilTest).iInstrNo + 1
                                                Exit For
                                            End If
                                        Next ilTest
                                        If Not ilFound Then
                                            tmCnfRot(UBound(tmCnfRot)) = tmCnf
                                            tmCnfRot(UBound(tmCnfRot)).iInstrNo = 1
                                            ReDim Preserve tmCnfRot(0 To UBound(tmCnfRot) + 1) As CNF
                                        End If
                                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                    ilCombineOk = False
                                    Do
                                        If ilCombineIndex = -1 Then
                                            ilCrfIndex = tmRotInfo(ilLoop).iSortCrfIndex
                                            If ilCrfIndex >= 0 Then
                                                ilCombineIndex = tgSortCrf(ilCrfIndex).iCombineIndex
                                            End If
                                        Else
                                            ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
                                        End If
                                        If ilCombineIndex >= 0 Then
                                            'Check if this rotation if sent to this vehicle
                                            ilVpfIndex = tgCombineCrf(ilCombineIndex).iVpfIndex
                                            'For ilCheck = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                                            '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilCheck) <> 0 Then
                                            '        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilCheck) = tmVef(ilVehIndex).iCode Then
                                            '            tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                                            '            ilCombineOk = True
                                            '            Exit For
                                            '        End If
                                            '    End If
                                            'Next ilCheck
                                            'ilVIndex = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                                            'Do While ilVIndex > 0
                                            '    ilVpfIndex = gVpfFind(ExpBkCpy, tmLkVehInfo(ilVIndex).iVefCode)
                                            '    If (tgVpf(ilVpfIndex).iSAGroupNo = ilSAGroupNo) And (ilSAGroupNo <> 0) Then
                                            '        ilVefSelected = True
                                            '        Exit Do
                                            '    End If
                                            '    ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                                            'Loop
                                            'If conventional- must use SAGroupNo- Code Later
                                            If ilVpfIndex >= 0 Then
                                                ilVIndex = tmVpfInfo(ilVpfIndex).iFirstSALink
                                                Do While ilVIndex >= 0
                                                    If tmVef(ilVehIndex).iCode = tmSALink(ilVIndex).iVefCode Then
                                                        'tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                                                        'ilCombineOk = True
                                                        'Exit Do
                                                        If tmRotInfo(ilLoop).lCrfCode <> tgCombineCrf(ilCombineIndex).tCrf.lCode Then
                                                            tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                                                            ilCombineOk = True
                                                            Exit Do
                                                        End If
                                                    End If
                                                    ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                                                Loop
                                            End If
                                        Else
                                            ilCombineOk = True
                                        End If
                                    Loop While Not ilCombineOk
                                Loop While ilCombineIndex >= 0
                                ilTotalNoInstr = 0
                                For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                    ilTotalNoInstr = ilTotalNoInstr + tmCnfRot(ilTest).iInstrNo
                                Next ilTest
                                For ilTest = 0 To UBound(tmCnfRot) - 1 Step 1
                                    tmCnf = tmCnfRot(ilTest)
                                    tmCifSrchKey.lCode = tmCnf.lCifCode
                                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tmCpf.sName = ""
                                            tmCpf.sISCI = ""
                                            tmCpf.sCreative = ""
                                        End If
                                        '6/2/16: Replace GoSub
                                        'GoSub cmcExportRotHeader
                                        mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                        slISCI = Trim$(tmCpf.sISCI)
                                        slRecord = slISCI
                                        Do While Len(slRecord) < 21
                                            slRecord = slRecord & " "
                                        Loop
                                        slRecord = slRecord & UCase(Trim$(tmCpf.sCreative))
                                        Do While Len(slRecord) < 54
                                            slRecord = slRecord & " "
                                        Loop
                                        slRecord = slRecord & Trim$(str$(tmCif.iLen))
                                        Do While Len(slRecord) < 58
                                            slRecord = slRecord & " "
                                        Loop
                                        'Rotation %
                                        slRecord = slRecord & gDivStr(gMulStr(Trim$(str$(tmCnf.iInstrNo)), "100"), str$(ilTotalNoInstr)) & "%"
                                        'Product Name
                                        Do While Len(slRecord) < 67
                                            slRecord = slRecord & " "
                                        Loop
                                        'If (Trim$(tmCpf.sName) <> "") And (StrComp(Trim$(tmChf.sProduct), Trim$(tmCpf.sName), 1) <> 0) Then
                                        '    slRecord = slRecord & Trim$(tmCpf.sName)
                                        'End If
                                        If Not mExportLine(slRecord, ilLineNo) Then
                                            Exit Sub
                                        End If
                                        '6/2/16: Replace GoSub
                                        'GoSub cmcExportRotHeader
                                        mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                        'Output cart
                                        'ilShowCart = False
                                        'If (InStr(UCase(slVehName), "SATELLITE MUSIC") > 0) Or (InStr(UCase(slVehName), "24-HOUR FORMATS") > 0) Then
                                        '    ilShowCart = True
                                        'End If
                                        'If (InStr(UCase(slVehName), "SMN MIX") > 0) Or (InStr(UCase(slVehName), "MIX 24-HOUR FORMATS") > 0) Then
                                        '    ilShowCart = True
                                        'End If
                                        If ilShowCart Then
                                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmMcf.sName = ""
                                                End If
                                            End If
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                                slCart = slCart & "-" & tmCif.sCut
                                            End If
                                            slRecord = " "
                                            Do While Len(slRecord) < 21
                                                slRecord = slRecord & " "
                                            Loop
                                            slRecord = slRecord & slCart
                                            If Not mExportLine(slRecord, ilLineNo) Then
                                                Exit Sub
                                            End If
                                            '6/2/16: Replace GoSub
                                            'GoSub cmcExportRotHeader
                                            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                        End If
                                        'Output comment if defined
                                        If tmCif.lCsfCode > 0 Then
                                            tmCsfSrchKey.lCode = tmCif.lCsfCode
                                            tmCsf.sComment = ""
                                            imCsfRecLen = Len(tmCsf) '5011
                                            ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                'Output 70 characters per line
                                                slComment = gStripChr0(tmCsf.sComment)
                                                'If tmCsf.iStrLen > 0 Then
                                                If slComment <> "" Then
                                                    'slComment = Trim$(Left$(tmCsf.sComment, tmCsf.iStrLen))
                                                    slRecord = " "
                                                    Do While Len(slRecord) < 21
                                                        slRecord = slRecord & " "
                                                    Loop
                                                    Do While Len(slComment) > 0
                                                        ilPos = InStr(slComment, " ")
                                                        If ilPos > 0 Then
                                                            If Len(slRecord) + ilPos - 1 > 70 Then
                                                                If Not mExportLine(slRecord, ilLineNo) Then
                                                                    Exit Sub
                                                                End If
                                                                '6/2/16: Replace GoSub
                                                                'GoSub cmcExportRotHeader
                                                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                                slRecord = " "
                                                                Do While Len(slRecord) < 21
                                                                    slRecord = slRecord & " "
                                                                Loop
                                                            End If
                                                            slRecord = slRecord & Left$(slComment, ilPos)
                                                            slComment = right$(slComment, Len(slComment) - ilPos)
                                                        Else
                                                            If Len(slRecord) + Len(slComment) > 70 Then
                                                                If Not mExportLine(slRecord, ilLineNo) Then
                                                                    Exit Sub
                                                                End If
                                                                '6/2/16: Replace GoSub
                                                                'GoSub cmcExportRotHeader
                                                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                                slRecord = " "
                                                                Do While Len(slRecord) < 21
                                                                    slRecord = slRecord & " "
                                                                Loop
                                                            End If
                                                            slRecord = slRecord & slComment
                                                            If Not mExportLine(slRecord, ilLineNo) Then
                                                                Exit Sub
                                                            End If
                                                            '6/2/16: Replace GoSub
                                                            'GoSub cmcExportRotHeader
                                                            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                            slComment = ""
                                                            Exit Do
                                                        End If
                                                    Loop
                                                End If
                                            End If
                                        End If
                                        slRecord = " "
                                        Do While Len(slRecord) < 21
                                            slRecord = slRecord & " "
                                        Loop
                                        ilNewInv = False
                                        For ilCheckInv = 0 To UBound(tmAddCyf) - 1 Step 1
                                            If (tmAddCyf(ilCheckInv).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilCheckInv).tCyf.iVefCode = tmVef(ilVehIndex).iCode) And (tmAddCyf(ilCheckInv).tCyf.sTimeZone = tmCrf.sZone) Then
                                                ilNewInv = True
                                                Exit For
                                            End If
                                        Next ilCheckInv
                                        If ilNewInv Then
                                            slRecord = slRecord & "Date fed (sent): " & slTranDate
                                            If Not mExportLine(slRecord, ilLineNo) Then
                                                Exit Sub
                                            End If
                                            '6/2/16: Replace GoSub
                                            'GoSub cmcExportRotHeader
                                            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                        Else
                                            'Find last sent date
                                            tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                            tmCyfSrchKey.iVefCode = tmVef(ilVehIndex).iCode
                                            tmCyfSrchKey.sSource = " "
                                            tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                            tmCyfSrchKey.lRafCode = 0
                                            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Then
                                                slRecord = slRecord & "Date fed (sent): " & slTranDate
                                                If Not mExportLine(slRecord, ilLineNo) Then
                                                    Exit Sub
                                                End If
                                                If ilTest <> UBound(tmCnfRot) - 1 Then
                                                    '6/2/16: Replace GoSub
                                                    'GoSub cmcExportRotHeader
                                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                End If
                                            Else
                                                gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slDate
                                                slDate = Format$(gDateValue(slDate), "mm/dd/yyyy")
                                                slRecord = slRecord & "Date fed (sent): " & slDate
                                                If Not mExportLine(slRecord, ilLineNo) Then
                                                    Exit Sub
                                                End If
                                                '6/2/16: Replace GoSub
                                                'GoSub cmcExportRotHeader
                                                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                slRecord = " "
                                                Do While Len(slRecord) < 19
                                                    slRecord = slRecord & " "
                                                Loop
                                                slRecord = slRecord & "(FED (SENT) PREVIOUSLY, NOT TODAY!)"
                                                If Not mExportLine(slRecord, ilLineNo) Then
                                                    Exit Sub
                                                End If
                                                If ilTest <> UBound(tmCnfRot) - 1 Then
                                                    '6/2/16: Replace GoSub
                                                    'GoSub cmcExportRotHeader
                                                    mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                                End If
                                            End If
                                        End If
                                        If ilTest <> UBound(tmCnfRot) - 1 Then
                                            'slRecord = ""
                                            If Not mExportLine(slBlank, ilLineNo) Then
                                                Exit Sub
                                            End If
                                            '6/2/16: Replace GoSub
                                            'GoSub cmcExportRotHeader
                                            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                                            'slRecord = ""
                                            If Not mExportLine(slBlank, ilLineNo) Then
                                                Exit Sub
                                            End If
                                        End If
                                        If (ilNewInv) And (hmNewCIF <> -1) Then
                                            ilFound = False
                                            For ilCheck = 0 To UBound(lmSentCifCode) - 1 Step 1
                                                If tmCif.lCode = lmSentCifCode(ilCheck) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilCheck
                                            If Not ilFound Then
                                                '7496
                                                'slStr = gFileNameFilter(slShortTitle) & "(" & gFileNameFilter(slISCI) & ").mp2"
                                                slStr = gFileNameFilter(slShortTitle) & "(" & gFileNameFilter(slISCI) & ")" & sgAudioExtension
                                                Print #hmNewCIF, slStr
                                            End If
                                        End If
                                        If (hmAllCIF <> -1) Then
                                            ilFound = False
                                            For ilCheck = 0 To UBound(lmSentCifCode) - 1 Step 1
                                                If tmCif.lCode = lmSentCifCode(ilCheck) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilCheck
                                            If Not ilFound Then
                                                '7496
                                                'slStr = gFileNameFilter(slShortTitle) & "(" & gFileNameFilter(slISCI) & ").mp2"
                                                slStr = gFileNameFilter(slShortTitle) & "(" & gFileNameFilter(slISCI) & ")" & sgAudioExtension
                                                Print #hmAllCIF, slStr
                                            End If
                                        End If
                                        ilFound = False
                                        For ilCheck = 0 To UBound(lmSentCifCode) - 1 Step 1
                                            If tmCif.lCode = lmSentCifCode(ilCheck) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        Next ilCheck
                                        If Not ilFound Then
                                            lmSentCifCode(UBound(lmSentCifCode)) = tmCif.lCode
                                            ReDim Preserve lmSentCifCode(0 To UBound(lmSentCifCode) + 1) As Long
                                        End If
                                    End If
                                Next ilTest
                            End If
                        End If
                    End If
                End If
            Next ilLoop
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replace GoSub
            'GoSub cmcExportRotHeader
            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
            '6/2/16: Replace GoSub
            'GoSub cmcExportRotHeader
            mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
            slRecord = "TOTAL NUMBER OF PAGES FOR " & slVehName & ":"
            slStr = Trim$(str$(ilPageNo))
            Do While Len(slStr) < 5
                slStr = " " & slStr
            Loop
            slRecord = slRecord & slStr
            If Not mExportLine(slRecord, ilLineNo) Then
                Exit Sub
            End If
            If ilLineNo < 50 Then
                '6/2/16: Replace GoSub
                'GoSub cmcExportRotHeader
                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
                'Add Messages
                If Not mExportLine(slBlank, ilLineNo) Then
                    Exit Sub
                End If
                '6/2/16: Replace GoSub
                'GoSub cmcExportRotHeader
                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
            End If
            ilMsgType = 1
            slMsgFileName = slStnCode & "000000.End"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
            slMsgFileName = slStnCode & slFeedNo & ".End"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
            slMsgFileName = "00" & slFeedNo & ".End"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
            slMsgFileName = "00000000.End"
            '6/3/16: Replaced GoSub
            'GoSub cmcExportSendMsg
            mExportSendMsg slMsgFileName, slMsgFile, slMsgLine, ilMsgType, ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv, slRecord
            Close hmExport
            Close #hmNewCIF
            Close #hmAllCIF
        End If
    Next ilVeh
    'Update files
    ilRet = btrBeginTrans(hmCrf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("File in Use [Re-press Export], BeginTran" & str(ilRet), vbOkOnly + vbExclamation, "Export")
        Exit Sub
    End If
    gAutomationAlertAndLogHandler "Updating copy inventory.."
    
    lacProcessing.Caption = "Updating Copy Inventory"
    DoEvents
    For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
        'If transmitted- remove old record, then insert instead of updating
        Do
            tmCyfSrchKey.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
            tmCyfSrchKey.iVefCode = tmAddCyf(ilLoop).tCyf.iVefCode
            tmCyfSrchKey.sSource = " "
            tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
            tmCyfSrchKey.lRafCode = 0
            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                ilRet = btrDelete(hmCyf)
            Else
                ilRet = BTRV_ERR_NONE
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Sub
        End If
        ilRet = btrInsert(hmCyf, tmAddCyf(ilLoop).tCyf, imCyfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Sub
        End If
        'Test if airing and group vehicles defined- if so insert for other
        'vehicles
        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
            If tmVef(ilLoop1).iCode = tmAddCyf(ilLoop).tCyf.iVefCode Then
                If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                    'Update cyf for all vehicles
                    ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                    'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                    If ilVIndex >= 0 Then
                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                        Do While ilVIndex >= 0
                            Do
                                tmCyfSrchKey.lCifCode = tmAddCyf(ilLoop).tCyf.lCifCode
                                tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                                tmCyfSrchKey.sSource = " "
                                tmCyfSrchKey.sTimeZone = tmAddCyf(ilLoop).tCyf.sTimeZone
                                tmCyfSrchKey.lRafCode = 0
                                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                    ilRet = btrDelete(hmCyf)
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("File in Use [Re-press Export], Delete Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                                Exit Sub
                            End If
                            tlCyf = tmAddCyf(ilLoop).tCyf
                            tlCyf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                            ilRet = btrInsert(hmCyf, tlCyf, imCyfRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("File in Use [Re-press Export], Insert Cyf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                                Exit Sub
                            End If
                        'Next ilLoop2
                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                        Loop
                    End If
                End If
            End If
        Next ilLoop1
    Next ilLoop
    gAutomationAlertAndLogHandler "Updating Copy Rotations..."
    lacProcessing.Caption = "Updating Copy Rotations"
    DoEvents
    For ilLoop = 0 To UBound(tmRotInfo) - 1 Step 1
        Do
            tmCrfSrchKey.lCode = tmRotInfo(ilLoop).lCrfCode
            ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("File in Use [Re-press Export], GetEqual Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                Exit Sub
            End If
            tmCrf.sFeedStatus = "S" '"S"
            tmCrf.iFeedDate(0) = ilTranDate0
            tmCrf.iFeedDate(1) = ilTranDate1
            ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("File in Use [Re-press Export], Update Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
            Exit Sub
        End If
        ilCrfIndex = tmRotInfo(ilLoop).iSortCrfIndex
        If ilCrfIndex >= 0 Then
            ilCombineIndex = tgSortCrf(ilCrfIndex).iCombineIndex
            Do While ilCombineIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgCombineCrf(ilCombineIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("File in Use [Re-press Export], GetDirect Combine Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                        Exit Sub
                    End If
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    'tmCrf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("File in Use [Re-press Export], GetByKey Combine Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
                    '    Exit Sub
                    'End If
                    tmCrf.sFeedStatus = "S" '"S"
                    tmCrf.iFeedDate(0) = ilTranDate0
                    tmCrf.iFeedDate(1) = ilTranDate1
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("File in Use [Re-press Export], Update Combine Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                    Exit Sub
                End If
                ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
            Loop
            ilDuplIndex = tgSortCrf(ilCrfIndex).iDuplIndex
            Do While ilDuplIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilDuplIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("File in Use [Re-press Export], GetDirect Dupl Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                        Exit Sub
                    End If
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    'tmCrf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    mAbortTrans     'ilCRet = btrAbortTrans(hmCrf)
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("File in Use [Re-press Export], GetByKey Dupl Crf" & Str(ilRet), vbOkOnly + vbExclamation, "Export")
                    '    Exit Sub
                    'End If
                    tmCrf.sFeedStatus = "S" '"S"
                    tmCrf.iFeedDate(0) = ilTranDate0
                    tmCrf.iFeedDate(1) = ilTranDate1
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    mAbortTrans 'ilCRet = btrAbortTrans(hmCrf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("File in Use [Re-press Export], Update Dupl Crf" & str(ilRet), vbOkOnly + vbExclamation, "Export")
                    Exit Sub
                End If
                ilDuplIndex = tgDuplCrf(ilDuplIndex).iDuplIndex
            Loop
        End If
    Next ilLoop
    ilRet = btrEndTrans(hmCrf)
    'DoEvents
    'For ilLoop = 0 To UBound(tgDuplCrf) - 1 Step 1
    '    Do
    '        ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilLoop).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    '        If ilRet = BTRV_ERR_NONE Then
    '            tmCrf.sFeedStatus = "S" '"S"
    '            tmCrf.iFeedDate(0) = ilTranDate0
    '            tmCrf.iFeedDate(1) = ilTranDate1
    '            ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
    '        End If
    '    Loop While ilRet = BTRV_ERR_CONFLICT
    'Next ilLoop
    lacProcessing.Caption = "Generating Cross Reference"
    DoEvents
    'Make key Short Title Cart # ISCI Creative Title, then vehicle name
    For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
        slKey = Trim$(tmAddCyf(ilLoop).sKey)
        slKey = tmAddCyf(ilLoop).sXFKey & "|" & slKey
        tmAddCyf(ilLoop).sKey = slKey
    Next ilLoop
    ilUpper = UBound(tmAddCyf)
    If ilUpper > 0 Then
        'ArraySortTyp fnAV(tgSort(),0), ilUpper, 0, LenB(tgSort(0)), 0, LenB, 0
        ArraySortTyp fnAV(tmAddCyf(), 0), ilUpper, 0, LenB(tmAddCyf(0)), 0, LenB(tmAddCyf(0).sKey), 0
    End If
    slStnCode = "X"
    slExportFile = sgExportPath & slStnCode & slFeedNo & slXRefLetter & ".asc"
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmExport = FreeFile
    ''Create file name based on vehicle name
    'Open slExportFile For Output As hmExport
    ilRet = gFileOpen(slExportFile, "Output", hmExport)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        ''MsgBox "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slExportFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    'Output new inventory
    ilPageNo = 0
    ilLineNo = 48
    slLine = ""
    slPrevProdISCITitle = " "
    For ilLoop = 0 To UBound(tmAddCyf) - 1 Step 1
        slKey = Trim$(tmAddCyf(ilLoop).sKey)
        ilRet = gParseItem(slKey, 1, "|", slProduct)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 2, "|", slCart)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 3, "|", slISCI)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 4, "|", slCreative)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 5, "|", slName)  'Obtain Index and code number
        ilShowName = False
        'For ilTest = 0 To UBound(tgVpf) Step 1
        '    If tmAddCyf(ilLoop).tCyf.iVefCode = tgVpf(ilTest).iVefKCode Then
            ilTest = gBinarySearchVpf(tmAddCyf(ilLoop).tCyf.iVefCode)
            If ilTest <> -1 Then
                If tgVpf(ilTest).sBulkXFer = "Y" Then
                    ilShowName = True
                End If
        '        Exit For
            End If
        'Next ilTest
        If Not ilShowName Then
            'Test if airing and group vehicles defined- if so test other
            'vehicles
            For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                If tmVef(ilLoop1).iCode = tmAddCyf(ilLoop).tCyf.iVefCode Then
                    If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                        'Update cyf for all vehicles
                        ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                        'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                        If ilVIndex >= 0 Then
                            ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                            Do While ilVIndex >= 0
                                'For ilTest = 0 To UBound(tgVpf) Step 1
                                '    'If tmVpfInfo(ilVIndex).iVefLink(ilLoop2) = tgVpf(ilTest).iVefKCode Then
                                '    If tmLkVehInfo(ilVIndex).iVefCode = tgVpf(ilTest).iVefKCode Then
                                    ilTest = gBinarySearchVpf(tmLkVehInfo(ilVIndex).iVefCode)
                                    If ilTest <> -1 Then
                                        If tgVpf(ilTest).sBulkXFer = "Y" Then
                                            ilShowName = True
                                        End If
                                '        Exit For
                                    End If
                                'Next ilTest
                                If ilShowName Then
                                    Exit For
                                End If
                            'Next ilLoop2
                                ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                            Loop
                        End If
                        If ilShowName Then
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop1
        End If
        'If InStr(UCase(slName), "DIRECTION") > 0 Then
        '    ilShowName = True
        'End If
        'If InStr(UCase(slName), "INFORMATION") > 0 Then
        '    ilShowName = True
        'End If
        'If InStr(UCase(slName), "ENTERTAINMENT") > 0 Then
        '    ilShowName = True
        'End If
        'If InStr(UCase(slName), "CONTEMPORARY") > 0 Then
        '    ilShowName = True
        'End If
        'If InStr(UCase(slName), "FM") > 0 Then
        '    ilShowName = True
        'End If
        'If InStr(UCase(slName), "ROCK") > 0 Then
        '    ilShowName = True
        'End If
        'If (InStr(UCase(slName), "SATELLITE MUSIC") > 0) Or (InStr(UCase(slName), "24-HOUR FORMATS") > 0) Then
        '    ilShowName = True
        'End If
        'If (InStr(UCase(slName), "SMN MIX") > 0) Or (InStr(UCase(slName), "MIX 24-HOUR FORMATS") > 0) Then
        '    ilShowName = True
        'End If
        If ilShowName Then
            Do While Len(slProduct) < 20
                slProduct = slProduct & " "
            Loop
            Do While Len(slCart) < 10
                slCart = slCart & " "
            Loop
            Do While Len(slISCI) < 20
                slISCI = slISCI & " "
            Loop
            Do While Len(slCreative) < 30
                slCreative = slCreative & " "
            Loop
            slProdISCITitle = slProduct & " " & slCart & " " & slISCI & " " & slCreative
            If slPrevProdISCITitle <> slProdISCITitle Then
                'First- output blank line
                If slPrevProdISCITitle <> " " Then
                    If Not mExportLine(slBlank, ilLineNo) Then
                        Exit Sub
                    End If
                End If
                slPrevProdISCITitle = slProdISCITitle
                slLine = slProdISCITitle
                slLine = slLine & " " & slName
                slPrevVehName = ""
            Else
                slLine = " "
                Do While Len(slLine) < Len(slProdISCITitle)
                    slLine = slLine & " "
                Loop
                slLine = slLine & " " & slName
            End If
            If slPrevVehName <> slName Then
                ilOldPageNo = ilPageNo
                '6/3/16: Replaced GoSub
                'GoSub cmcExportXFerHeader
                If Not mExportXFerHeader(ilLineNo, ilPageNo, slTranDate, slRecord) Then
                    Exit Sub
                End If
                If (ilOldPageNo <> ilPageNo) And (ilOldPageNo > 0) Then
                    slLine = slProdISCITitle
                    slLine = slLine & " " & slName
                End If
                If Not mExportLine(slLine, ilLineNo) Then
                    Exit Sub
                End If
                slPrevVehName = slName
            End If
        End If
    Next ilLoop
    Close hmExport
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    mVehPop False
    'For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
    '    lbcVeh.Selected(ilLoop) = False
    '    tmcRot.Enabled = False
    'Next ilLoop
    mRotPop
    lacProcessing.Caption = ""
    
    gAutomationAlertAndLogHandler "Completed Export"
    Screen.MousePointer = vbDefault
    imExporting = False
    cmcCancel.SetFocus
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
'cmcExportCopyHeader:
'    If ilLineNo + ilNoCopyLines > 52 Then
'        If ilPageNo = 0 Then
'            'slRecord = ""
'            If Not mExportLine(slBlank, ilLineNo) Then
'                Exit Sub
'            End If
'        Else
'            slCopyHeader = Chr(12)  'Form Feed
'            If Not mExportLine(slCopyHeader, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        slCopyHeader = " "
'        Do While Len(slCopyHeader) < 35
'            slCopyHeader = slCopyHeader & " "
'        Loop
'        slCopyHeader = slCopyHeader & Trim$(tgSpf.sGClient)
'        If Not mExportLine(slCopyHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slCopyHeader = " "
'        Do While Len(slCopyHeader) < 35
'            slCopyHeader = slCopyHeader & " "
'        Loop
'        slCopyHeader = slCopyHeader & slVehName
'        If Not mExportLine(slCopyHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slCopyHeader = " "
'        Do While Len(slCopyHeader) < 35
'            slCopyHeader = slCopyHeader & " "
'        Loop
'        slCopyHeader = slCopyHeader & "Commercial Feed # " & slStnCode & "-" & slFeedNo
'        If Not mExportLine(slCopyHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slCopyHeader = " "
'        Do While Len(slCopyHeader) < 35
'            slCopyHeader = slCopyHeader & " "
'        Loop
'        slCopyHeader = slCopyHeader & slTranDate & "  "
'        slCopyHeader = slCopyHeader & "Page:"
'        slStr = Trim$(str$(ilPageNo))
'        Do While Len(slStr) < 5
'            slStr = " " & slStr
'        Loop
'        slCopyHeader = slCopyHeader & slStr
'        If Not mExportLine(slCopyHeader, ilLineNo) Then
'            Exit Sub
'        End If
'        slCopyHeader = ""
'        If Not mExportLine(slBlank, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slBlank, ilLineNo) Then
'            Exit Sub
'        End If
'        If (slPrevNewInv <> "") And (ilNoCopyLines = 1) Then
'            If Not mExportLine(slPrevNewInv, ilLineNo) Then
'                Exit Sub
'            End If
'            slCopyHeader = "     "
'            slCopyHeader = slCopyHeader & "------------------------------"
'            If Not mExportLine(slCopyHeader, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'    End If
'    Return
'cmcExportRotHeader:
'    If ilLineNo >= 52 Then
'        If ilPageNo > 0 Then
'            slRecord = Chr(12)  'Form Feed
'            If Not mExportLine(slRecord, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        slRecord = "-"
'        Do While Len(slRecord) < 78
'            slRecord = slRecord & "-"
'        Loop
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 68
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & "Page:"
'        slStr = Trim$(str$(ilPageNo))
'        Do While Len(slStr) < 5
'            slStr = " " & slStr
'        Loop
'        slRecord = slRecord & slStr
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = ""
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = Trim$(tgSpf.sGClient) & " " & slVehName & " Network Feed Instructions " & slTranDate
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = ""
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'    End If
'    Return
'cmcExportXFerHeader:
'    If ilLineNo >= 48 Then
'        If ilPageNo = 0 Then
'            slRecord = ""
'            If Not mExportLine(slRecord, ilLineNo) Then
'                Exit Sub
'            End If
'        Else
'            slRecord = Chr(12)  'Form Feed
'            If Not mExportLine(slRecord, ilLineNo) Then
'                Exit Sub
'            End If
'        End If
'        ilPageNo = ilPageNo + 1
'        ilLineNo = 0
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & Trim$(tgSpf.sGClient)
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & "Cross Reference"
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = " "
'        Do While Len(slRecord) < 35
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & slTranDate & "  "
'        slRecord = slRecord & "Page:"
'        slStr = Trim$(str$(ilPageNo))
'        Do While Len(slStr) < 5
'            slStr = " " & slStr
'        Loop
'        slRecord = slRecord & slStr
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = ""
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'        slRecord = "Short Title"
'        Do While Len(slRecord) < 20
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Cart"
'        Do While Len(slRecord) < 31
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " ISCI"
'        Do While Len(slRecord) < 52
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Creative Title"
'        Do While Len(slRecord) < 83
'            slRecord = slRecord & " "
'        Loop
'        slRecord = slRecord & " Vehicle"
'        If Not mExportLine(slRecord, ilLineNo) Then
'            Exit Sub
'        End If
'    End If
'    Return
'cmcExportSendMsg:
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    hmMsg = FreeFile
'    slMsgFile = sgExportPath & slMsgFileName
'    Open slMsgFile For Input Access Read As hmMsg
'    If ilRet = 0 Then
'        Do
'            On Error GoTo cmcExportErr:
'            Line Input #hmMsg, slMsgLine
'            On Error GoTo 0
'            If (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            End If
'            If Len(slMsgLine) > 0 Then
'                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
'                    Exit Do
'                End If
'                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
'                If ilPos > 0 Then
'                    Mid$(slMsgLine, ilPos) = slTranDate
'                End If
'            End If
'            If Not mExportLine(slMsgLine, ilLineNo) Then
'                Exit Sub
'            End If
'            If ilMsgType = 0 Then
'                '6/2/16: Replaced GoSub
'                'GoSub cmcExportCopyHeader
'                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
'            ElseIf ilMsgType = 1 Then
'                '6/2/16: Replace GoSub
'                'GoSub cmcExportRotHeader
'                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
'            End If
'        Loop
'        Close hmMsg
'    End If
'    Return
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    

End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_LostFocus()
    lacProcessing.Caption = ""
End Sub

Private Sub cmcReSend_Click()
    igBFCall = 1
    BulkFeed.Show vbModal
    If igBFReturn = 1 Then
        Screen.MousePointer = vbHourglass
        ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
        mVehPop False
        mRotPop
        Screen.MousePointer = vbDefault
        mSetCommands
    End If
End Sub
Private Sub cmcReSend_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcSuppress_Click()
    Dim ilCrf As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer

    If lbcRot.ListCount <= 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
        If tgSortCrf(ilCrf).iSelected Then
            ilIndex = tgSortCrf(ilCrf).iCombineIndex
            Do While ilIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgCombineCrf(ilIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        'tmRec = tmCrf
                        'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                        'tmCrf = tmRec
                        tmCrf.sFeedStatus = "P" '"S"
                        ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilIndex = tgCombineCrf(ilIndex).iCombineIndex
            Loop
            ilIndex = tgSortCrf(ilCrf).iDuplIndex
            Do While ilIndex >= 0
                Do
                    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgDuplCrf(ilIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        'tmRec = tmCrf
                        'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                        'tmCrf = tmRec
                        tmCrf.sFeedStatus = "P" '"S"
                        ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilIndex = tgDuplCrf(ilIndex).iDuplIndex
            Loop
            Do
                ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tgSortCrf(ilCrf).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet = BTRV_ERR_NONE Then
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    'tmCrf = tmRec
                    tmCrf.sFeedStatus = "P" 'Suppress
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    Next ilCrf
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    mVehPop False
    mRotPop
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub
Private Sub cmcSuppress_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcTranDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcTranDate_GotFocus()
    Dim slStr As String
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    '    'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 3 Then
        plcCalendar.Visible = False
        slStr = edcTranDate.Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = 3
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_Click()
    UndoBkFd.Show vbModal
    If igBFReturn = 1 Then
        ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
        Screen.MousePointer = vbHourglass
        mVehPop False
        mRotPop
        Screen.MousePointer = vbDefault
        mSetCommands
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcTranDate_Change()
    Dim slStr As String
    slStr = edcTranDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcTranDate_GotFocus()
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
    If imDateBox <> 3 Then
        plcCalendar.Visible = False
    End If
    imDateBox = 3
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTranDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTranDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTranDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcTranDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
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
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
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
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmSTCrf
    Erase tgSortCrf
    Erase imSelectPrevState
    Erase lmRotCodeBuild
    Erase tmCnfRot
    Erase tmCyfTest
    Erase tmVef
    Erase tmVpfInfo
    Erase tmLkVehInfo
    Erase tmSALink
    Erase tmVehTimes
    Erase tmRotInfo
    Erase tmAddCyf
    Erase tgCombineCrf
    Erase tgDuplCrf
    Erase smFileNames
    Erase imVefCode
    Erase imDVefCode
    Erase lmSentCifCode
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    Set ExpBkCpy = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRot_Click()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilListIndex As Integer
    Dim ilResetLast As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    Screen.MousePointer = vbHourglass
    If ckcAll.Value = vbChecked Then
        imIgnoreCkcAll = True
        'ckcAll.Value = False
        ckcAll.Value = vbUnchecked
        'DoEvents
        imIgnoreCkcAll = False
    End If
    ilResetLast = True
    ReDim imSelectPrevState(LBound(tgSortCrf) To UBound(tgSortCrf)) As Integer
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        imSelectPrevState(ilLoop) = tgSortCrf(ilLoop).iSelected
    Next ilLoop
    ilListIndex = imCurrentIndex + vbcRot.Value
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tgSortCrf) - 1 Then
        ilEndIndex = UBound(tgSortCrf) - 1
    End If
    If ((imShiftKey And 1) = 1) And (imLastIndex >= 0) Then
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tgSortCrf(ilLoop).iSelected = False
        Next ilLoop
        If imLastIndex <= ilListIndex Then
            For ilLoop = imLastIndex To ilListIndex Step 1
                tgSortCrf(ilLoop).iSelected = True
            Next ilLoop
        Else
            For ilLoop = ilListIndex To imLastIndex Step 1
                tgSortCrf(ilLoop).iSelected = True
            Next ilLoop
        End If
        ilValue = False
        If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
            llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
        Else
            llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
        End If
        llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            lbcRot.Selected(ilIndex) = tgSortCrf(ilLoop).iSelected
            ilIndex = ilIndex + 1
        Next ilLoop
        ilResetLast = False
    ElseIf ((imShiftKey And 2) = 2) Then    'Ctrl
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tgSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
            tgSortCrf(ilLoop).iSelected = False
        Next ilLoop
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tgSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    End If
    'DoEvents
    If ilResetLast Then
        imLastIndex = ilListIndex
    End If
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        If imSelectPrevState(ilLoop) <> tgSortCrf(ilLoop).iSelected Then
            If imSelectPrevState(ilLoop) Then
                tmCrf = tgSortCrf(ilLoop).tCrf
                mComputeTime False
                ilIndex = tgSortCrf(ilLoop).iCombineIndex
                Do While ilIndex >= 0
                    tmCrf = tgCombineCrf(ilIndex).tCrf
                    mComputeTime False
                    ilIndex = tgCombineCrf(ilIndex).iCombineIndex
                Loop
                ilIndex = tgSortCrf(ilLoop).iDuplIndex
                Do While ilIndex >= 0
                    tmCrf = tgDuplCrf(ilIndex).tCrf
                    mComputeTime False
                    ilIndex = tgDuplCrf(ilIndex).iDuplIndex
                Loop
            Else
                tmCrf = tgSortCrf(ilLoop).tCrf
                mComputeTime True
                ilIndex = tgSortCrf(ilLoop).iCombineIndex
                Do While ilIndex >= 0
                    tmCrf = tgCombineCrf(ilIndex).tCrf
                    mComputeTime True
                    ilIndex = tgCombineCrf(ilIndex).iCombineIndex
                Loop
                ilIndex = tgSortCrf(ilLoop).iDuplIndex
                Do While ilIndex >= 0
                    tmCrf = tgDuplCrf(ilIndex).tCrf
                    mComputeTime True
                    ilIndex = tgDuplCrf(ilIndex).iDuplIndex
                Loop
            End If
            'DoEvents
        End If
    Next ilLoop
    mSetCommands
    pbcLbcRot_Paint
    pbclbcVehicle_Paint
    DoEvents
    Screen.MousePointer = vbDefault
    imIgnoreVbcChg = False
End Sub
Private Sub lbcRot_GotFocus()
    plcCalendar.Visible = False
    'If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    '    mInitDDE
    '    mSendHelpMsg "BT"
    'End If
End Sub
Private Sub lbcRot_KeyDown(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_KeyUp(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imCurrentIndex = Y \ fgListHtArial825
    imButton = Button
    If Button = 2 Then  'Right Mouse
        imButtonIndex = imCurrentIndex + vbcRot.Value
        If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tgSortCrf) - 1) Then
            imIgnoreRightMove = True
            mShowRotInfo
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcRot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 0) Or (Y > lbcRot.Height) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcRot.Width) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If imButtonIndex <> (Y \ fgListHtArial825) + vbcRot.Value Then
            imIgnoreRightMove = True
            imButtonIndex = Y \ fgListHtArial825 + vbcRot.Value
            If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tgSortCrf) - 1) Then
                mShowRotInfo
            Else
                plcRotInfo.Visible = False
            End If
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcRot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        plcRotInfo.Visible = False
    End If
End Sub

Private Sub lbcRot_Scroll()
    pbcLbcRot_Paint
End Sub

Private Sub lbcVeh_Click()
    tmcRot.Enabled = False
    tmcRot.Enabled = True
End Sub
Private Sub lbcVeh_GotFocus()
    plcCalendar.Visible = False
    tmcRot.Enabled = False
End Sub
Private Sub lbcVeh_TopIndexChange(TopIndex As Integer)
    If tmcRot.Enabled Then
        tmcRot.Enabled = False
        tmcRot.Enabled = True
    End If
End Sub

Private Sub lbcVeh_Scroll()
    If tmcRot.Enabled Then
        tmcRot.Enabled = False
        tmcRot.Enabled = True
    End If
End Sub

Private Sub lbcVehicle_Click()
    'pbclbcVehicle_Paint  'ignore click events D.S. 11/13/02
End Sub

Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAbortTrans                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Abort Transaction and remove   *
'*                      files created                  *
'*                                                     *
'*******************************************************
Private Sub mAbortTrans()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = btrAbortTrans(hmCrf)
    On Error GoTo 0
    For ilLoop = 0 To UBound(smFileNames) - 1 Step 1
        Kill smFileNames(ilLoop)
    Next ilLoop
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
        'slStr = edcStartDate.Text
    ElseIf imDateBox = 2 Then
        'slStr = edcEndDate.Text
    ElseIf imDateBox = 3 Then
        slStr = edcTranDate.Text
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildExpTable                  *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build table from the selected   *
'*                     rotations for exporting         *
'*                                                     *
'*******************************************************
Private Sub mBuildExpTable()
'   mBuildExpTable
'
    Dim ilCrf As Integer
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVpfIndex As Integer
    Dim slKey As String
    Dim ilVpf As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim llSTStartDate As Long
    Dim llSTEndDate As Long
    Dim llRotStartTime As Long
    Dim llRotEndTime As Long
    Dim llSTStartTime As Long
    Dim llSTEndTime As Long
    Dim ilDay As Integer
    Dim slFeedDate As String
    Dim llTranDate As Long
    Dim ilTranDate0 As Integer
    Dim ilTranDate1 As Integer
    Dim slProduct As String
    Dim ilName As Integer
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilTransmit As Integer
    Dim ilVIndex As Integer
    Dim slCart As String
    Dim ilCombineIndex As Integer
    Dim ilDuplIndex As Integer
    Dim ilDone As Integer
    Dim ilSend As Integer
    Dim ilRot As Integer
    Dim ilVefCode As Integer
    Dim tlCrf As CRF
    ReDim tmRotInfo(0 To 0) As SENDROTINFO
    ReDim tmAddCyf(0 To 0) As SENDCOPYINFO
    ReDim lmRotCodeBuild(0 To 0) As Long
    slDate = edcTranDate.Text
    llTranDate = gDateValue(slDate)
    gPackDate slDate, ilTranDate0, ilTranDate1
    For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
        If tgSortCrf(ilCrf).iSelected Then
            tmCrf = tgSortCrf(ilCrf).tCrf
            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
            llRotStartDate = gDateValue(slDate)
            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
            llRotEndDate = gDateValue(slDate)
            gUnpackTimeLong tmCrf.iStartTime(0), tmCrf.iStartTime(1), False, llRotStartTime
            gUnpackTimeLong tmCrf.iEndTime(0), tmCrf.iEndTime(1), True, llRotEndTime
            ilCombineIndex = -1
            ilDuplIndex = -1
            ilDone = False
            ReDim ilVehSent(0 To 0) As Integer
            Do
                ilSend = True
                For ilRot = 0 To UBound(lmRotCodeBuild) - 1 Step 1
                    If tmCrf.lCode = lmRotCodeBuild(ilRot) Then
                        ilSend = False
                    End If
                Next ilRot
                If ilSend Then
                    lmRotCodeBuild(UBound(lmRotCodeBuild)) = tmCrf.lCode
                    ReDim Preserve lmRotCodeBuild(0 To UBound(lmRotCodeBuild) + 1) As Long
                    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                    Next ilVeh
                    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
                    'Loop thru all airing vehicles associated with selling or conventional vehicle
                    'If tmVef(ilVehIndex).sType <> "S" Then
                    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
                    '    Next ilVeh
                    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
                    'End If
                    'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                    ReDim imVefCode(0 To 0) As Integer
                    ilVefCode = tmVef(ilVehIndex).iCode
                    If tmVef(ilVehIndex).sType = "S" Then
                        'Find airing vehicle
                        If ilVpfIndex >= 0 Then
                            ilVpf = tmVpfInfo(ilVpfIndex).iFirstSALink
                            Do While ilVpf >= 0
                                imVefCode(UBound(imVefCode)) = tmSALink(ilVpf).iVefCode
                                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
                            Loop
                        End If
                    Else
                        ReDim imVefCode(0 To 1) As Integer
                        imVefCode(0) = ilVefCode
                    End If
                    For ilVeh = 0 To UBound(imVefCode) - 1 Step 1
                        ilVefCode = imVefCode(ilVeh)
                            ilSend = True
                            For ilRot = 0 To UBound(ilVehSent) - 1 Step 1
                                If ilVehSent(ilRot) = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                    ilSend = False
                                    Exit For
                                End If
                            Next ilRot
                            If ilSend Then
                                ilVehSent(UBound(ilVehSent)) = ilVefCode    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                ReDim Preserve ilVehSent(0 To UBound(ilVehSent) + 1) As Integer
                                ReDim Preserve tmRotInfo(0 To UBound(tmRotInfo) + 1) As SENDROTINFO
                                For ilName = 0 To UBound(tmVef) - 1 Step 1
                                    If tmVef(ilName).iCode = ilVefCode Then 'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                        slKey = tmVef(ilName).sName
                                        ilVpf = mFindVpfIndex(ilVefCode)    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh))
                                        'If ilVpf > 0 Then
                                        '    For ilLoop1 = LBound(tmVpfInfo(ilVpf).sVefName) To tmVpfInfo(ilVpf).iNoVefLinks - 1 Step 1
                                        '        slKey = Trim$(slKey) & " " & tmVpfInfo(ilVpf).sVefName(ilLoop1)
                                        '    Next ilLoop1
                                        'End If
                                        If ilVpf >= 0 Then
                                            ilVpf = tmVpfInfo(ilVpf).iFirstLkVehInfo
                                            Do While ilVpf >= 0
                                                slKey = Trim$(slKey) & " " & tmLkVehInfo(ilVpf).sVefName
                                                ilVpf = tmLkVehInfo(ilVpf).iNextLkVehInfo
                                            Loop
                                        End If
                                        Exit For
                                    End If
                                Next ilName
                                slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                slStr = Trim$(str$(llRotStartDate))
                                Do While Len(slStr) < 6
                                    slStr = "0" & slStr
                                Loop
                                slKey = slKey & "|" & slStr
                                For ilDay = 0 To 6 Step 1
                                    If tmCrf.sDay(ilDay) <> "N" Then
                                        slStr = Trim$(str$(ilDay))
                                        Exit For
                                    End If
                                Next ilDay
                                slKey = slKey & "|" & slStr
                                slStr = Trim$(str$(llRotStartTime))
                                Do While Len(slStr) < 6
                                    slStr = "0" & slStr
                                Loop
                                slKey = slKey & "|" & slStr
                                Select Case tmCrf.sZone
                                    Case "EST"
                                        slKey = slKey & "|1"
                                    Case "CST"
                                        slKey = slKey & "|2"
                                    Case "MST"
                                        slKey = slKey & "|3"
                                    Case "PST"
                                        slKey = slKey & "|4"
                                    Case Else
                                        slKey = slKey & "|5"
                                End Select
                                tmRotInfo(UBound(tmRotInfo) - 1).sKey = slKey
                                tmRotInfo(UBound(tmRotInfo) - 1).lCrfCode = tmCrf.lCode
                                tmRotInfo(UBound(tmRotInfo) - 1).iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                tmRotInfo(UBound(tmRotInfo) - 1).iRevised = False
                                'If ilCombineIndex = -1 Then
                                    tmRotInfo(UBound(tmRotInfo) - 1).iSortCrfIndex = ilCrf
                                'Else
                                '    tmRotInfo(UBound(tmRotInfo) - 1).iSortCrfIndex = -1
                                'End If
                                'Determine rotation status (status = 1 = send; =2=duplicate- don't send)
                                tmRotInfo(UBound(tmRotInfo) - 1).iStatus = 1
                                'Determine if this is a revision
                                tmCrfSrchKey1.sRotType = tmCrf.sRotType
                                tmCrfSrchKey1.iEtfCode = tmCrf.iEtfCode
                                tmCrfSrchKey1.iEnfCode = tmCrf.iEnfCode
                                tmCrfSrchKey1.iAdfCode = tmCrf.iAdfCode
                                tmCrfSrchKey1.lChfCode = tmCrf.lChfCode
                                tmCrfSrchKey1.iVefCode = tmCrf.iVefCode
                                tmCrfSrchKey1.iRotNo = tmCrf.iRotNo
                                ilRet = btrGetGreaterOrEqual(hmCrf, tlCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (tlCrf.sRotType = tmCrf.sRotType) And (tlCrf.iEtfCode = tmCrf.iEtfCode) And (tlCrf.iEnfCode = tmCrf.iEnfCode) And (tlCrf.iAdfCode = tmCrf.iAdfCode) And (tlCrf.lChfCode = tmCrf.lChfCode) And (tlCrf.iVefCode = tmCrf.iVefCode) And (tlCrf.iRotNo < tmCrf.iRotNo)
                                    If Trim$(tlCrf.sZone) <> "R" Then
                                        If (tlCrf.sFeedStatus = "S") And (tlCrf.sInOut = tmCrf.sInOut) And (tlCrf.ianfCode = tmCrf.ianfCode) And (tlCrf.iLen = tmCrf.iLen) Then
                                            If tlCrf.sState <> "D" Then
                                                gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slDate
                                                llSTStartDate = gDateValue(slDate)
                                                gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slDate
                                                llSTEndDate = gDateValue(slDate)
                                                If (llRotEndDate <= llSTStartDate) And (llSTEndDate >= llRotStartDate) Then
                                                    gUnpackTimeLong tlCrf.iStartTime(0), tlCrf.iStartTime(1), False, llSTStartTime
                                                    gUnpackTimeLong tlCrf.iEndTime(0), tlCrf.iEndTime(1), True, llSTEndTime
                                                    If (llRotEndTime <= llSTStartTime) And (llSTEndTime >= llRotStartTime) Then
                                                        tmRotInfo(UBound(tmRotInfo) - 1).iRevised = True
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmCrf, tlCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                    '    End If
                    Next ilVeh
                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                    tmCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        'Loop thru all airing vehicle if selling, otherwise test conventional
                        'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                        '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                        For ilVeh = 0 To UBound(imVefCode) - 1 Step 1
                            ilVefCode = imVefCode(ilVeh)
                                ilTransmit = True
                                tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                tmCyfSrchKey.iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                tmCyfSrchKey.sSource = " "
                                tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                tmCyfSrchKey.lRafCode = 0
                                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                    gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                                    If gDateValue(slFeedDate) + 90 > llTranDate Then
                                        'If airing- check if other vehicles has receive inventory
                                        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                                            If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                                If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                                                    ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                                                    ilTransmit = False
                                                    'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                                                    If ilVIndex >= 0 Then
                                                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                                                        Do While ilVIndex >= 0
                                                            tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                                            tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                                                            tmCyfSrchKey.sSource = " "
                                                            tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                                            tmCyfSrchKey.lRafCode = 0
                                                            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                                                gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                                                                If gDateValue(slFeedDate) + 90 <= llTranDate Then
                                                                    ilTransmit = True
                                                                End If
                                                            Else
                                                                ilTransmit = True
                                                                Exit For
                                                            End If
                                                        'Next ilLoop2
                                                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                                                        Loop
                                                    End If
                                                Else
                                                    ilTransmit = False
                                                End If
                                            End If
                                        Next ilLoop1
                                    End If
                                End If
                                If ilTransmit Then
                                    'Test for duplicates
                                    ilFound = False
                                    For ilLoop = LBound(tmAddCyf) To UBound(tmAddCyf) - 1 Step 1
                                        'Create records for each zone but only show cart for
                                        'one zone in the inventory area
                                        'in mComputeTime- the zone test is removed
                                        If (tmAddCyf(ilLoop).tCyf.lCifCode = tmCnf.lCifCode) And (tmAddCyf(ilLoop).tCyf.iVefCode = ilVefCode) Then   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                            If llRotStartDate < tmAddCyf(ilLoop).lRotStartDate Then
                                                tmAddCyf(ilLoop).lRotStartDate = llRotStartDate
                                            End If
                                            If llRotEndDate > tmAddCyf(ilLoop).lRotEndDate Then
                                                tmAddCyf(ilLoop).lRotEndDate = llRotEndDate
                                            End If
                                            If (tmAddCyf(ilLoop).tCyf.sTimeZone = tmCrf.sZone) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        For ilName = 0 To UBound(tmVef) - 1 Step 1
                                            If tmVef(ilName).iCode = ilVefCode Then 'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                                ilFound = True
                                                slKey = tmVef(ilName).sName
                                                ilVpf = mFindVpfIndex(ilVefCode)    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh))
                                                'If ilVpf > 0 Then
                                                '    For ilLoop1 = LBound(tmVpfInfo(ilVpf).sVefName) To tmVpfInfo(ilVpf).iNoVefLinks - 1 Step 1
                                                '        slKey = Trim$(slKey) & " " & tmVpfInfo(ilVpf).sVefName(ilLoop1)
                                                '    Next ilLoop1
                                                'End If
                                                If ilVpf >= 0 Then
                                                    ilVpf = tmVpfInfo(ilVpf).iFirstLkVehInfo
                                                    Do While ilVpf >= 0
                                                        slKey = Trim$(slKey) & " " & tmLkVehInfo(ilVpf).sVefName
                                                        ilVpf = tmLkVehInfo(ilVpf).iNextLkVehInfo
                                                    Loop
                                                End If
                                                Exit For
                                            End If
                                        Next ilName
                                        If ilFound Then
                                            ReDim Preserve tmAddCyf(0 To UBound(tmAddCyf) + 1) As SENDCOPYINFO
                                            slProduct = tgSortCrf(ilCrf).sCntrProd
                                            tmCifSrchKey.lCode = tmCnf.lCifCode
                                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'If Trim$(tmCpf.sName) <> "" Then
                                                    If (Trim$(tmCpf.sName) <> "") And (tgSpf.sUseProdSptScr <> "P") Then
                                                        slKey = slKey & "|" & tmCpf.sName
                                                        slProduct = tmCpf.sName
                                                    Else
                                                        slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                    End If
                                                Else
                                                    slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                    tmCpf.sISCI = ""
                                                    tmCpf.sCreative = ""
                                                End If
                                            Else
                                                slKey = slKey & "|" & tgSortCrf(ilCrf).sCntrProd
                                                tmCpf.sISCI = ""
                                                tmCpf.sCreative = ""
                                            End If
                                            If tmMcf.iCode <> tmCif.iMcfCode Then
                                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmMcf.sName = ""
                                                End If
                                            End If
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                                            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                                                slCart = slCart & "-" & tmCif.sCut
                                            End If
                                            'slStr = Trim$(Str$(tmCnf.lCifCode))
                                            'Do While Len(slStr) < 6
                                            '    slStr = "0" & slStr
                                            'Loop
                                            slKey = slKey & "|" & slCart    'slStr
                                            slStr = Trim$(str$(llRotStartDate))
                                            Do While Len(slStr) < 6
                                                slStr = "0" & slStr
                                            Loop
                                            slKey = slKey & "|" & slStr
                                            slStr = Trim$(str$(llRotEndDate))
                                            Do While Len(slStr) < 6
                                                slStr = "0" & slStr
                                            Loop
                                            slKey = slKey & "|" & slStr
                                            tmAddCyf(UBound(tmAddCyf) - 1).sKey = slKey
                                            tmAddCyf(UBound(tmAddCyf) - 1).sXFKey = slProduct & "|" & slCart & "|" & tmCpf.sISCI & "|" & tmCpf.sCreative
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.lCifCode = tmCnf.lCifCode
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iVefCode = ilVefCode    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sSource = " "
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.sTimeZone = tmCrf.sZone
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iFeedDate(0) = ilTranDate0
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.iFeedDate(1) = ilTranDate1
                                            tmAddCyf(UBound(tmAddCyf) - 1).tCyf.lRafCode = 0
                                            tmAddCyf(UBound(tmAddCyf) - 1).sChfProduct = Trim$(tgSortCrf(ilCrf).sCntrProd)
                                            tmAddCyf(UBound(tmAddCyf) - 1).lRotStartDate = llRotStartDate
                                            tmAddCyf(UBound(tmAddCyf) - 1).lRotEndDate = llRotEndDate
                                        End If
                                    End If
                                End If
                        '    End If
                        Next ilVeh
                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If ilCombineIndex = -1 Then
                    ilCombineIndex = tgSortCrf(ilCrf).iCombineIndex
                ElseIf ilCombineIndex >= 0 Then
                    ilCombineIndex = tgCombineCrf(ilCombineIndex).iCombineIndex
                End If
                If ilCombineIndex < 0 Then
                    ilCombineIndex = -2
                Else
                    tmCrf = tgCombineCrf(ilCombineIndex).tCrf
                End If
                If ilCombineIndex = -2 Then
                    If ilDuplIndex = -1 Then
                        ilDuplIndex = tgSortCrf(ilCrf).iDuplIndex
                    Else
                        ilDuplIndex = tgDuplCrf(ilDuplIndex).iDuplIndex
                    End If
                    If ilDuplIndex < 0 Then
                        ilDone = True
                    Else
                        tmCrf = tgDuplCrf(ilDuplIndex).tCrf
                    End If
                End If
            Loop While Not ilDone
        End If
    Next ilCrf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeTime                    *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Computer Inventory time         *
'*                                                     *
'*******************************************************
Private Sub mComputeTime(ilInc As Integer)
'   mComputeTime ilInc
'   Where:
'       ilInc(I)- True =increment time; False=Remove time
'
'       tmCrf(I)- Rotation
'
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim ilIndex As Integer
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilVpfIndex As Integer
    Dim slFeedDate As String
    Dim slDate As String
    Dim llTranDate As Long
    Dim ilTest As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slTime As String
    Dim slCode As String
    Dim ilPos As Integer
    Dim slMin As String
    Dim slSec As String
    Dim ilMin As Integer
    Dim ilSec As Integer
    Dim llTime As Long
    Dim ilVIndex As Integer
    Dim ilTransmit As Integer
    Dim ilVefCode As Integer
    ilVehIndex = -1
    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
            ilVehIndex = ilVeh
            Exit For
        End If
    Next ilVeh
    If ilVehIndex = -1 Then
        For ilVeh = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilVeh).sType = "A") Or (tmVef(ilVeh).sType = "C") Then
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            Else
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            End If
        Next ilVeh
    End If
    If ilVehIndex = -1 Then
        Exit Sub
    End If
    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
    'iGLink will contain primary airing for group vehicles or airing vehicle
    'without groups
    'iGLink for Convention is zero- replace with conventional vehicle
    'If tmVef(ilVehIndex).sType <> "S" Then
    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
    '    Next ilVeh
    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
    'End If
    ReDim imVefCode(0 To 0) As Integer
    ilVefCode = tmVef(ilVehIndex).iCode
    If tmVef(ilVehIndex).sType = "S" Then
        'Find airing vehicle
        'ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, ilVefCode)).iSAGroupNo
        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
        '            imVefCode(UBound(imVefCode)) = tgMVef(ilVeh).iCode
        '            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        '        End If
        '    End If
        'Next ilVeh
        If ilVpfIndex >= 0 Then
            ilVIndex = tmVpfInfo(ilVpfIndex).iFirstSALink
            Do While ilVIndex >= 0
                imVefCode(UBound(imVefCode)) = tmSALink(ilVIndex).iVefCode
                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
            Loop
        End If
    Else
        ReDim imVefCode(0 To 1) As Integer
        imVefCode(0) = ilVefCode
    End If
    slDate = edcTranDate.Text
    llTranDate = gDateValue(slDate)
    tmCnfSrchKey.lCrfCode = tmCrf.lCode
    tmCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
        'Loop thru all airing vehicle if selling, otherwise test conventional
        'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
        '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
        For ilVeh = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
            ilVefCode = imVefCode(ilVeh)
                ilTransmit = True
                tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                tmCyfSrchKey.iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                tmCyfSrchKey.sSource = " "
                tmCyfSrchKey.sTimeZone = tmCrf.sZone
                tmCyfSrchKey.lRafCode = 0
                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                    gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                    If gDateValue(slFeedDate) + 90 > llTranDate Then
                        'If airing- check if other vehicles has receive inventory
                        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                            If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                If (tmVef(ilLoop1).sType = "A") Or (tmVef(ilLoop1).sType = "C") Then
                                    ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                                    ilTransmit = False
                                    'For ilLoop2 = LBound(tmVpfInfo(ilVIndex).iVefLink) To tmVpfInfo(ilVIndex).iNoVefLinks - 1 Step 1
                                    If ilVIndex >= 0 Then
                                        ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                                        Do While ilVIndex >= 0
                                            tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                            tmCyfSrchKey.iVefCode = tmLkVehInfo(ilVIndex).iVefCode 'tmVpfInfo(ilVIndex).iVefLink(ilLoop2)
                                            tmCyfSrchKey.sSource = " "
                                            tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                            tmCyfSrchKey.lRafCode = 0
                                            ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet = BTRV_ERR_NONE Then   'Test Date- if 90 days- resend
                                                gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slFeedDate
                                                If gDateValue(slFeedDate) + 90 <= llTranDate Then
                                                    ilTransmit = True
                                                End If
                                            Else
                                                ilTransmit = True
                                                Exit For
                                            End If
                                        'Next ilLoop2
                                            ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                                        Loop
                                    End If
                                Else
                                    ilTransmit = False
                                End If
                            End If
                        Next ilLoop1
                    End If
                End If
                If ilTransmit Then
                    'Test for duplicates
                    ilFound = False
                    For ilLoop = LBound(tmCyfTest) To UBound(tmCyfTest) - 1 Step 1
                        'If (tmCyfTest(ilLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilLoop).iVefCode = tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) And (tmCyfTest(ilLoop).sTimeZone = tmCrf.sZone) Then
                        If (tmCyfTest(ilLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilLoop).iVefCode = ilVefCode) And (tmCyfTest(ilLoop).sTimeZone = tmCrf.sZone) Then
                            'ilFound = True
                            If Not ilInc Then
                                For ilIndex = ilLoop + 1 To UBound(tmCyfTest) - 1 Step 1
                                    'Ignore zone when counting time, when building
                                    'records in mBuildExpTable build one for each zone but one
                                    'show one piece of inventory without zone
                                    'If (tmCyfTest(ilIndex).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilIndex).iVefCode = tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) And (tmCyfTest(ilIndex).sTimeZone = tmCrf.sZone) Then
                                    If (tmCyfTest(ilIndex).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilIndex).iVefCode = ilVefCode) Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                        ilFound = True  'Found a second time- retain time
                                    End If
                                    tmCyfTest(ilIndex - 1).lCifCode = tmCyfTest(ilIndex).lCifCode
                                    tmCyfTest(ilIndex - 1).iVefCode = tmCyfTest(ilIndex).iVefCode
                                    tmCyfTest(ilIndex - 1).sSource = tmCyfTest(ilIndex).sSource
                                    tmCyfTest(ilIndex - 1).sTimeZone = tmCyfTest(ilIndex).sTimeZone
                                    tmCyfTest(ilIndex - 1).lRafCode = tmCyfTest(ilIndex).lRafCode
                                Next ilIndex
                                ReDim Preserve tmCyfTest(0 To UBound(tmCyfTest) - 1) As CYFTEST
                            Else
                                ilFound = True
                            End If
                            Exit For
                        Else
                            If (tmCyfTest(ilLoop).lCifCode = tmCnf.lCifCode) And (tmCyfTest(ilLoop).iVefCode = ilVefCode) Then  'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)) Then
                                ilFound = True  'Found a second time- retain time
                            End If
                        End If
                    Next ilLoop
                    If ilInc Then
                        ReDim Preserve tmCyfTest(0 To UBound(tmCyfTest) + 1) As CYFTEST
                        tmCyfTest(UBound(tmCyfTest) - 1).lCifCode = tmCnf.lCifCode
                        tmCyfTest(UBound(tmCyfTest) - 1).iVefCode = ilVefCode   'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh)
                        tmCyfTest(UBound(tmCyfTest) - 1).sSource = ""
                        tmCyfTest(UBound(tmCyfTest) - 1).sTimeZone = tmCrf.sZone
                        tmCyfTest(UBound(tmCyfTest) - 1).lRafCode = 0
                    End If

                    If Not ilFound Then
                        'Adjust time
                        For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
                            If tmVef(ilLoop1).iCode = ilVefCode Then    'tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) Then
                                ilVIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
                                For ilTest = 0 To lbcVehicleCode.ListCount - 1 Step 1
                                    slNameCode = lbcVehicleCode.List(ilTest)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    If ilVIndex = Val(slCode) Then
                                        'Obtain time, then adjust Vehicle|Time
                                        slNameCode = lbcVehicle.List(ilTest)
                                        ilRet = gParseItem(slNameCode, 1, "|", slName)
                                        ilRet = gParseItem(slNameCode, 2, "|", slTime)
                                        ilPos = InStr(slTime, ":")
                                        slMin = Left$(slTime, ilPos - 1)
                                        slSec = Mid$(slTime, ilPos + 1)
                                        llTime = 60 * Val(slMin) + Val(slSec)
                                        If ilInc Then
                                            llTime = llTime + tmCrf.iLen
                                        Else
                                            llTime = llTime - tmCrf.iLen
                                        End If
                                        ilSec = llTime Mod 60
                                        ilMin = llTime \ 60
                                        slTime = Trim$(str$(ilMin)) & ":" & Trim$(str$(ilSec))
                                        lbcVehicle.List(ilTest) = slName & "|" & slTime
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilTest
                                If ilFound Then
                                    Exit For
                                End If
                            End If
                        Next ilLoop1
                    End If
                End If
        '    End If
        Next ilVeh
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDuplRotation                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if this is a duplicate*
'*                     rotations                       *
'*                                                     *
'*******************************************************
Private Function mDuplRotation(ilVpfIndex As Integer, llRecPos As Long) As Integer
'   ilDupl = mDuplRotation()
'   Where:
'
'       ilVpfIndex(O)- Vpf Index
'       ilDupl(O)- 1 if this rotation should be combined with a selected rotation
'                  2 if this rotation matches a selected rotation
'                  0 if this rotation does not match a selected rotation
'
'       tmCrf(I)- Rotation to be checked
'       tmChf(I)- Contract associated with tmCrf
'
'   Note: you can't have tmCrf be "Combined" or "Matched" with more then
'         one Crf (they would have been "combined" or "Matched" previously)
'
    Dim ilRet As Integer
    Dim ilCrf As Integer
    Dim tlCrf As CRF
    Dim ilVeh As Integer
    Dim ilVehIndex As Integer
    Dim ilDVeh As Integer
    Dim ilDVehIndex As Integer
    Dim ilDVpfIndex As Integer
    Dim ilMatch As Integer
    Dim ilDay As Integer
    Dim ilTest As Integer
    Dim ilUpper As Integer
    Dim llSifCode As Long
    Dim ilVsf As Integer
    Dim slStr As String
    Dim ilDuplIndex As Integer
    Dim ilCombIndex As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim ilVIndex As Integer
    ReDim llCifCode(0 To 0) As Long
    ilVehIndex = -1
    For ilVeh = 0 To UBound(tmVef) - 1 Step 1
        If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
            ilVehIndex = ilVeh
            Exit For
        End If
    Next ilVeh
    If ilVehIndex = -1 Then
        For ilVeh = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilVeh).sType = "A") Or (tmVef(ilVeh).sType = "C") Then
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstLkVehInfo
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmLkVehInfo(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmLkVehInfo(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            Else
                ilVIndex = mFindVpfIndex(tmVef(ilVeh).iCode)
                If ilVIndex >= 0 Then
                    ilVIndex = tmVpfInfo(ilVIndex).iFirstSALink
                    Do While ilVIndex >= 0
                        If tmCrf.iVefCode = tmSALink(ilVIndex).iVefCode Then
                            ilVehIndex = ilVeh
                            Exit For
                        End If
                        ilVIndex = tmSALink(ilVIndex).iNextLkVehInfo
                    Loop
                End If
            End If
        Next ilVeh
    End If
    If ilVehIndex = -1 Then
        mDuplRotation = 0
        Exit Function
    End If
    ilVpfIndex = mFindVpfIndex(tmVef(ilVehIndex).iCode)
    'If tmVef(ilVehIndex).sType <> "S" Then
    '    For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
    '        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = 0
    '    Next ilVeh
    '    tmVpfInfo(ilVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink)) = tmVef(ilVehIndex).iCode
    'End If
    ReDim imVefCode(0 To 0) As Integer
    ilVefCode = tmVef(ilVehIndex).iCode
    If tmVef(ilVehIndex).sType = "S" Then
        'Find airing vehicle
        'ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, ilVefCode)).iSAGroupNo
        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
        '            imVefCode(UBound(imVefCode)) = tgMVef(ilVeh).iCode
        '            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        '        End If
        '    End If
        'Next ilVeh
        If ilVpfIndex >= 0 Then
            ilVpf = tmVpfInfo(ilVpfIndex).iFirstSALink
            Do While ilVpf >= 0
                imVefCode(UBound(imVefCode)) = tmSALink(ilVpf).iVefCode
                ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
            Loop
        End If
    Else
        ReDim imVefCode(0 To 1) As Integer
        imVefCode(0) = ilVefCode
    End If
    ilDuplIndex = -1
    ilCombIndex = -1
    'Get instructions
    ilUpper = 0
    tmCnfSrchKey.lCrfCode = tmCrf.lCode
    tmCnfSrchKey.iInstrNo = 0
    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
        llCifCode(ilUpper) = tmCnf.lCifCode
        ilUpper = ilUpper + 1
        ReDim Preserve llCifCode(0 To ilUpper)
        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDuplRotation = 0
    'For ilCrf = 0 To UBound(tgSortCrf) - 1 Step 1
    ilCrf = 0
    Do While ilCrf <= UBound(tgSortCrf) - 1
        tlCrf = tgSortCrf(ilCrf).tCrf
        If tlCrf.lCode <> tmCrf.lCode Then
            'Determine if sent to same vehicle
            ilMatch = True
            If tmAdf.iCode <> tmChf.iAdfCode Then
                tmAdfSrchKey.iCode = tmChf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            llSifCode = 0
            If tmChf.lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tmChf.lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While ilRet = BTRV_ERR_NONE
                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                            If tmVsf.lFSComm(ilVsf) > 0 Then
                                llSifCode = tmVsf.lFSComm(ilVsf)
                            End If
                            Exit For
                        End If
                    Next ilVsf
                    If llSifCode <> 0 Then
                        Exit Do
                    End If
                    If tmVsf.lLkVsfCode <= 0 Then
                        Exit Do
                    End If
                    tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            slStr = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)
            'If StrComp(Trim$(tmChf.sProduct), Trim$(tgSortCrf(ilCrf).sCntrProd), 1) <> 0 Then
            If StrComp(Trim$(slStr), Trim$(tgSortCrf(ilCrf).sCntrProd), 1) <> 0 Then
                ilMatch = False
            End If
            If tmCrf.iAdfCode <> tlCrf.iAdfCode Then
                ilMatch = False
            End If
            If tmCrf.sRotType <> tlCrf.sRotType Then
                ilMatch = False
            End If
            'Determine if Dates, times, days,... are the same
            If (tmCrf.iStartDate(0) <> tlCrf.iStartDate(0)) Or (tmCrf.iStartDate(1) <> tlCrf.iStartDate(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iEndDate(0) <> tlCrf.iEndDate(0)) Or (tmCrf.iEndDate(1) <> tlCrf.iEndDate(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iStartTime(0) <> tlCrf.iStartTime(0)) Or (tmCrf.iStartTime(1) <> tlCrf.iStartTime(1)) Then
                ilMatch = False
            End If
            If (tmCrf.iEndTime(0) <> tlCrf.iEndTime(0)) Or (tmCrf.iEndTime(1) <> tlCrf.iEndTime(1)) Then
                ilMatch = False
            End If
            For ilDay = 0 To 6 Step 1
                If (tmCrf.sDay(ilDay) <> tlCrf.sDay(ilDay)) Then
                    ilMatch = False
                    Exit For
                End If
            Next ilDay
            If tmCrf.sZone <> tlCrf.sZone Then
                ilMatch = False
            End If
            If tmCrf.iLen <> tlCrf.iLen Then
                ilMatch = False
            End If
            If tmCrf.sInOut <> tlCrf.sInOut Then
                ilMatch = False
            End If
            If tmCrf.ianfCode <> tlCrf.ianfCode Then
                ilMatch = False
            End If
            If ilMatch Then
                If tmCrf.iVefCode <> tlCrf.iVefCode Then
                    For ilDVeh = 0 To UBound(tmVef) - 1 Step 1
                        If tmVef(ilDVeh).iCode = tlCrf.iVefCode Then
                            ilDVehIndex = ilDVeh
                            Exit For
                        End If
                    Next ilDVeh
                    ilDVpfIndex = mFindVpfIndex(tmVef(ilDVehIndex).iCode)
                    'If tmVef(ilDVehIndex).sType <> "S" Then
                    '    For ilDVeh = LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) Step 1
                    '        tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) = 0
                    '    Next ilDVeh
                    '    tmVpfInfo(ilDVpfIndex).tVpf.iGLink(LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink)) = tmVef(ilDVehIndex).iCode
                    'End If
                    ReDim imDVefCode(0 To 0) As Integer
                    If tmVef(ilDVehIndex).sType = "S" Then
                        'Find airing vehicle
                        'ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tmVef(ilDVehIndex).iCode)).iSAGroupNo
                        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If (tgMVef(ilVeh).sType = "A") And (tgMVef(ilVeh).sState <> "D") Then
                        '        If (ilSAGroupNo = tgVpf(gVpfFind(ExpBkCpy, tgMVef(ilVeh).iCode)).iSAGroupNo) And (ilSAGroupNo <> 0) Then
                        '            imDVefCode(UBound(imDVefCode)) = tgMVef(ilVeh).iCode
                        '            ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                        '        End If
                        '    End If
                        'Next ilVeh
                        If ilDVpfIndex >= 0 Then
                            ilVpf = tmVpfInfo(ilDVpfIndex).iFirstSALink
                            Do While ilVpf >= 0
                                imDVefCode(UBound(imDVefCode)) = tmSALink(ilVpf).iVefCode
                                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                                ilVpf = tmSALink(ilVpf).iNextLkVehInfo
                            Loop
                        End If
                    Else
                        ReDim imDVefCode(0 To 1) As Integer
                        imDVefCode(0) = tmVef(ilDVehIndex).iCode
                    End If
                    'Only one vehicle has to match
                    ilMatch = False
                    'For ilVeh = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
                    '    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) <> 0 Then
                    '        For ilDVeh = LBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) To UBound(tmVpfInfo(ilDVpfIndex).tVpf.iGLink) Step 1
                    '            If tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) <> 0 Then
                    '                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilVeh) = tmVpfInfo(ilDVpfIndex).tVpf.iGLink(ilDVeh) Then
                    '                ilMatch = True
                    '                Exit For
                    '                End If
                    '            End If
                    '        Next ilDVeh
                    '        If ilMatch Then
                    '            Exit For
                    '        End If
                    '    End If
                    'Next ilVeh
                    For ilVeh = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
                        For ilDVeh = LBound(imDVefCode) To UBound(imDVefCode) - 1 Step 1
                            If imVefCode(ilVeh) = imDVefCode(ilDVeh) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilDVeh
                        If ilMatch Then
                            Exit For
                        End If
                    Next ilVeh

                End If
            End If
            If ilMatch Then
            'Test if same instructions
                ilTest = 0
                tmCnfSrchKey.lCrfCode = tlCrf.lCode
                tmCnfSrchKey.iInstrNo = 0
                ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tlCrf.lCode)
                    If ilTest >= UBound(llCifCode) Then
                        ilMatch = False
                        Exit Do
                    End If
                    If tmCnf.lCifCode <> llCifCode(ilTest) Then
                        ilMatch = False
                        Exit Do
                    End If
                    ilTest = ilTest + 1
                    ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ilMatch Then
                    If ilDuplIndex = -1 Then
                        ilDuplIndex = ilCrf
                        If tgSortCrf(ilCrf).iDuplIndex = -1 Then
                            tgSortCrf(ilCrf).iDuplIndex = UBound(tgDuplCrf)
                        Else
                            'Test if combine or superseded
                            'Assume combine
                            ilTest = tgSortCrf(ilCrf).iDuplIndex
                            Do While tgDuplCrf(ilTest).iDuplIndex >= 0
                                ilTest = tgDuplCrf(ilTest).iDuplIndex
                            Loop
                            tgDuplCrf(ilTest).iDuplIndex = UBound(tgDuplCrf)
                        End If
                        'System can't have combine and superseded
                        mDuplRotation = 2
                        tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tmChf.lCntrNo
                        tgDuplCrf(UBound(tgDuplCrf)).sVehName = tmVef(ilVehIndex).sName
                        tgDuplCrf(UBound(tgDuplCrf)).tCrf = tmCrf
                        tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = llRecPos
                        tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = -1
                        tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = ilVpfIndex
                        ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                    Else
                        'Merge
                        If tgSortCrf(ilDuplIndex).iDuplIndex = -1 Then
                            tgSortCrf(ilDuplIndex).iDuplIndex = UBound(tgDuplCrf)
                        Else
                            'Test if combine or superseded
                            'Assume combine
                            ilTest = tgSortCrf(ilDuplIndex).iDuplIndex
                            Do While tgDuplCrf(ilTest).iDuplIndex >= 0
                                ilTest = tgDuplCrf(ilTest).iDuplIndex
                            Loop
                            tgDuplCrf(ilTest).iDuplIndex = UBound(tgDuplCrf)
                        End If
                        'System can't have combine and superseded
                        mDuplRotation = 2
                        tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tgSortCrf(ilCrf).lCntrNo
                        slNameCode = tgSortCrf(ilCrf).sKey
                        ilRet = gParseItem(slNameCode, 3, "|", slName)
                        tgDuplCrf(UBound(tgDuplCrf)).sVehName = slName
                        tgDuplCrf(UBound(tgDuplCrf)).tCrf = tgSortCrf(ilCrf).tCrf
                        tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = tgSortCrf(ilCrf).lCrfRecPos
                        tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = tgSortCrf(ilCrf).iDuplIndex
                        tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = tgSortCrf(ilCrf).iVpfIndex
                        ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                        For ilTest = ilCrf To UBound(tgSortCrf) - 1 Step 1
                            tgSortCrf(ilTest) = tgSortCrf(ilTest + 1)
                        Next ilTest
                        ReDim Preserve tgSortCrf(0 To UBound(tgSortCrf) - 1) As SORTCRF
                        ilCrf = ilCrf - 1   'Test next which has been moved
                    End If
                    'Exit Function
                Else
                    'Set combine flag
                    If tmVef(ilVehIndex).sType = "S" Then
                        If ilCombIndex = -1 Then
                            ilCombIndex = ilCrf
                            If tgSortCrf(ilCrf).iCombineIndex = -1 Then
                                tgSortCrf(ilCrf).iCombineIndex = UBound(tgCombineCrf)
                            Else
                                'Test if combine or superseded
                                'Assume combine
                                ilTest = tgSortCrf(ilCrf).iCombineIndex
                                Do While tgCombineCrf(ilTest).iCombineIndex >= 0
                                    ilTest = tgCombineCrf(ilTest).iCombineIndex
                                Loop
                                tgCombineCrf(ilTest).iCombineIndex = UBound(tgCombineCrf)
                            End If
                            mDuplRotation = 1
                            tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tmChf.lCntrNo
                            slNameCode = tgSortCrf(ilCrf).sKey
                            ilRet = gParseItem(slNameCode, 3, "|", slName)
                            tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilVehIndex).sName 'slName
                            tgCombineCrf(UBound(tgCombineCrf)).tCrf = tmCrf
                            tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = llRecPos
                            tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = -1
                            tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = ilVpfIndex
                            ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                            'Exit Function
                        Else
                            'Merge
                            If tgSortCrf(ilCombIndex).iCombineIndex = -1 Then
                                tgSortCrf(ilCombIndex).iCombineIndex = UBound(tgCombineCrf)
                            Else
                                'Test if combine or superseded
                                'Assume combine
                                ilTest = tgSortCrf(ilCombIndex).iCombineIndex
                                Do While tgCombineCrf(ilTest).iCombineIndex >= 0
                                    ilTest = tgCombineCrf(ilTest).iCombineIndex
                                Loop
                                tgCombineCrf(ilTest).iCombineIndex = UBound(tgCombineCrf)
                            End If
                            'System can't have combine and superseded
                            mDuplRotation = 1
                            tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tgSortCrf(ilCrf).lCntrNo
                            tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilDVehIndex).sName
                            tgCombineCrf(UBound(tgCombineCrf)).tCrf = tgSortCrf(ilCrf).tCrf
                            tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = tgSortCrf(ilCrf).lCrfRecPos
                            tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = tgSortCrf(ilCrf).iCombineIndex
                            tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = tgSortCrf(ilCrf).iVpfIndex
                            ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                            For ilTest = ilCrf To UBound(tgSortCrf) - 1 Step 1
                                tgSortCrf(ilTest) = tgSortCrf(ilTest + 1)
                            Next ilTest
                            ReDim Preserve tgSortCrf(0 To UBound(tgSortCrf) - 1) As SORTCRF
                            ilCrf = ilCrf - 1   'Test next which has been moved
                        End If
                    End If
                End If
            End If
        End If
        ilCrf = ilCrf + 1
    Loop
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mExportLine                     *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Send line to output             *
'*                                                     *
'*******************************************************
Private Function mExportLine(slRecord As String, ilLineNo As Integer) As Integer
    Dim ilRet As Integer
    On Error GoTo mExportLineErr
    ilRet = 0
    Print #hmExport, slRecord
    If ilRet <> 0 Then
        imExporting = False
        Close #hmExport
        Screen.MousePointer = vbDefault
        ''MsgBox "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        gAutomationAlertAndLogHandler "Error writing to file" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        cmcCancel.SetFocus
        mExportLine = False
        Close #hmNewCIF
        Close #hmAllCIF
        Exit Function
    End If
    ilLineNo = ilLineNo + 1
    mExportLine = True
    Exit Function
mExportLineErr:
    ilRet = err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindVpfIndex                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find Vpf index for specified    *
'*                     vehicle                         *
'*                                                     *
'*******************************************************
Private Function mFindVpfIndex(ilVefCode As Integer) As Integer
    Dim ilLoop As Integer
    For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
        If ilVefCode = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
            mFindVpfIndex = ilLoop
            Exit Function
        End If
    Next ilLoop
    mFindVpfIndex = -1
End Function
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
    Dim slStr As String

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imIgnoreCkcAll = False
    imLastIndex = -1
    imShiftKey = 0
    imIgnoreVbcChg = False
    imTypeIndex = 0
    imTerminate = False
    ReDim tgSortCrf(0 To 0) As SORTCRF
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    
    ExpBkCpy.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone ExpBkCpy
    ExpBkCpy.Show
    'DoEvents
    imWaitCount = csiHandleValue(0, 12)
    imTimeDelay = csiHandleValue(0, 13)
    imLockValue = csiHandleValue(0, 14)
    imTranLog = csiHandleValue(0, 15)
    csiSetValue 90, 2, imLockValue, imTranLog
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imListFieldVeh(1) = 15
    imListFieldVeh(2) = 75 * igAlignCharWidth
    imListFieldVeh(3) = 90 * igAlignCharWidth
    imListFieldRot(1) = 15
    imListFieldRot(2) = 15 * igAlignCharWidth
    imListFieldRot(3) = 23 * igAlignCharWidth
    imListFieldRot(4) = 33 * igAlignCharWidth
    imListFieldRot(5) = 40 * igAlignCharWidth
    imListFieldRot(6) = 52 * igAlignCharWidth
    imListFieldRot(7) = 67 * igAlignCharWidth
    imListFieldRot(8) = 78 * igAlignCharWidth
    imListFieldRot(9) = 82 * igAlignCharWidth
    imListFieldRot(10) = 92 * igAlignCharWidth
    imListFieldRot(11) = 97 * igAlignCharWidth
    imListFieldRot(12) = 100 * igAlignCharWidth
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imDateBox = -1
    imIgnoreRightMove = False
    imButton = 0
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Chf.btr)", ExpBkCpy
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)     'Get and save CHF record length
    hmCrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Crf.Btr)", ExpBkCpy
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)     'Get and save CRF record length
    hmSif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Sif.Btr)", ExpBkCpy
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vsf.Btr)", ExpBkCpy
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmCnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cnf.Btr)", ExpBkCpy
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Mcf.Btr)", ExpBkCpy
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cif.Btr)", ExpBkCpy
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCyf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cyf.Btr)", ExpBkCpy
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Cpf.Btr)", ExpBkCpy
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmCsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Csf.Btr)", ExpBkCpy
    On Error GoTo 0
    imCsfRecLen = Len(tmCsf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Adf.Btr)", ExpBkCpy
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Anf.Btr)", ExpBkCpy
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vef.Btr)", ExpBkCpy
    On Error GoTo 0
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen-Vlf.Btr)", ExpBkCpy
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    mVehPop True
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    tmAdf.iCode = 0
    tmAnf.iCode = 0
    lbcVeh.Clear
    'ilRet = gPopUserVehicleBox(ExpBkCpy, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVeh, lbcVehCode)
    ilRet = gPopUserVehicleBox(ExpBkCpy, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVeh, tmVehCode(), smVehCodeTag)
    'mRotPop
    smTodaysDate = Format$(gNow(), "mm/dd/yy")
    edcTranDate.Text = smTodaysDate
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    cmcExport.Enabled = False
    'cmcView.Enabled = False
    imNoRotations = 0
    lacProcessing.Caption = ""
    tmcDDE.Enabled = True
    plcRotInfo.Move 1410, 4455
    'gCenterStdAlone ExpBkCpy
    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    imTerminate = True
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
'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move plcDates.Left + edcTranDate.Left, plcDates.Top + edcTranDate.Top + edcTranDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr  'Command$
'    If StrComp(slCommand, "Debug", 1) = 0 Then
'        igStdAloneMode = True 'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        imShowHelpMsg = False
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            ''MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            gAutomationAlertAndLogHandler "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ExpBkCpy, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRotPop                         *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                      Same code is in BulkFeed.Frm   *
'*                                                     *
'*******************************************************
Private Sub mRotPop()
'
'   iRet = mRotPop
'   Where:
'

    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilOffSet As Integer
    Dim llRevCntrNo As Long
    Dim slRevCntrNo As String
    Dim llRevRotNo As Long
    Dim slRevRotNo As String
    Dim llCntrNo As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilDay As Integer
    Dim ilExtLen As Integer
    Dim ilUpper As Integer
    Dim ilVehIndex As Integer
    Dim ilVeh As Integer
    Dim ilVpfIndex As Integer
    Dim ilVefSelected As Integer
    Dim ilVsf As Integer
    Dim llSifCode As Long
    ReDim tgSortCrf(0 To 0) As SORTCRF
    ilUpper = 0
    lbcRot.Clear
    pbcLbcRot_Paint
    ReDim tgDuplCrf(0 To 0) As DUPLCRF
    ReDim tgCombineCrf(0 To 0) As COMBINECRF
    ilVefSelected = False
    For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
        If lbcVeh.Selected(ilLoop) Then
            ilVefSelected = True
        End If
    Next ilLoop
    If Not ilVefSelected Then
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        vbcRot.Max = 0
        vbcRot.Value = 0
        imIgnoreVbcChg = False
       Exit Sub
    End If
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    If imTypeIndex = 0 Then
        tmCrfSrchKey1.sRotType = "A"
        tmCrfSrchKey1.iEtfCode = 0
        tmCrfSrchKey1.iEnfCode = 0
        tmCrfSrchKey1.iAdfCode = 0
        tmCrfSrchKey1.lChfCode = 0
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = 0
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        'Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC") 'Set extract limits (all records)
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Crf", "CrfFeedStatus")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "M", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "S", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "X", 1)
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddLogicConst):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        On Error GoTo mRotPopErr
        gBtrvErrorMsg ilRet, "mRotPop (btrExtAddField):" & "Crf.Btr", ExpBkCpy
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mRotPopErr
            gBtrvErrorMsg ilRet, "mRotPop (btrExtGetNextExt):" & "Clf.Btr", ExpBkCpy
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                If (Trim$(tmCrf.sZone) <> "R") And (tmCrf.sState <> "D") Then
                    ilVefSelected = False
                    For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
                        If lbcVeh.Selected(ilLoop) Then
                            slNameCode = tmVehCode(ilLoop).sKey    'lbcVehCode.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            If tmCrf.iVefCode = Val(slCode) Then
                                ilVefSelected = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If ilVefSelected Then
                        If tmChf.lCode <> tmCrf.lChfCode Then
                            tmChfSrchKey.lCode = tmCrf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mRotPopErr
                            gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Chf.Btr", ExpBkCpy
                            On Error GoTo 0
                        End If
                        If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
                            ilRet = mDuplRotation(ilVpfIndex, llRecPos)
                            For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                                If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                                    ilVehIndex = ilVeh
                                    Exit For
                                End If
                            Next ilVeh
                            If ilRet = 2 Then   'Match
                                'tgDuplCrf(UBound(tgDuplCrf)).lCntrNo = tmChf.lCntrNo
                                'tgDuplCrf(UBound(tgDuplCrf)).sVehName = tmVef(ilVehIndex).sName
                                'tgDuplCrf(UBound(tgDuplCrf)).tCrf = tmCrf
                                'tgDuplCrf(UBound(tgDuplCrf)).lCrfRecPos = llRecPos
                                'tgDuplCrf(UBound(tgDuplCrf)).iDuplIndex = -1
                                'tgDuplCrf(UBound(tgDuplCrf)).iVpfIndex = ilVpfIndex
                                'ReDim Preserve tgDuplCrf(0 To UBound(tgDuplCrf) + 1) As DUPLCRF
                            ElseIf ilRet = 1 Then   'Combine
                                'tgCombineCrf(UBound(tgCombineCrf)).lCntrNo = tmChf.lCntrNo
                                'tgCombineCrf(UBound(tgCombineCrf)).sVehName = tmVef(ilVehIndex).sName
                                'tgCombineCrf(UBound(tgCombineCrf)).tCrf = tmCrf
                                'tgCombineCrf(UBound(tgCombineCrf)).lCrfRecPos = llRecPos
                                'tgCombineCrf(UBound(tgCombineCrf)).iCombineIndex = -1
                                'tgCombineCrf(UBound(tgCombineCrf)).iVpfIndex = ilVpfIndex
                                'ReDim Preserve tgCombineCrf(0 To UBound(tgCombineCrf) + 1) As COMBINECRF
                            Else
                                llRevCntrNo = 99999999 - tmChf.lCntrNo
                                slRevCntrNo = Trim$(str$(llRevCntrNo))
                                Do While Len(slRevCntrNo) < 8
                                    slRevCntrNo = "0" & slRevCntrNo
                                Loop
                                'Scan for vehicle
                                'For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                                '    If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                                '        ilVehIndex = ilVeh
                                '        Exit For
                                '    End If
                                'Next ilVeh
                                If tmAdf.iCode <> tmCrf.iAdfCode Then
                                    tmAdfSrchKey.iCode = tmCrf.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    On Error GoTo mRotPopErr
                                    gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Adf.Btr", ExpBkCpy
                                    On Error GoTo 0
                                End If
                                slName = tmVef(ilVehIndex).sName
                                llRevRotNo = 99999 - tmCrf.iRotNo
                                slRevRotNo = Trim$(str$(llRevRotNo))
                                Do While Len(slRevRotNo) < 6
                                    slRevRotNo = "0" & slRevRotNo
                                Loop
                                ilUpper = UBound(tgSortCrf)
                                tgSortCrf(ilUpper).sKey = tmAdf.sName & "|" & slRevCntrNo & "|" & tmVef(ilVehIndex).sName & "|" & slRevRotNo
                                tgSortCrf(ilUpper).lCntrNo = tmChf.lCntrNo
                                llSifCode = 0
                                If tmChf.lVefCode < 0 Then
                                    tmVsfSrchKey.lCode = -tmChf.lVefCode
                                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Do While ilRet = BTRV_ERR_NONE
                                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                            If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                                                If tmVsf.lFSComm(ilVsf) > 0 Then
                                                    llSifCode = tmVsf.lFSComm(ilVsf)
                                                End If
                                                Exit For
                                            End If
                                        Next ilVsf
                                        If llSifCode <> 0 Then
                                            Exit Do
                                        End If
                                        If tmVsf.lLkVsfCode <= 0 Then
                                            Exit Do
                                        End If
                                        tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                End If
                                tgSortCrf(ilUpper).sCntrProd = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf) 'tmChf.sProduct
                                tgSortCrf(ilUpper).sType = tmChf.sType
                                tgSortCrf(ilUpper).lCrfRecPos = llRecPos
                                tgSortCrf(ilUpper).iSelected = False
                                tgSortCrf(ilUpper).iCombineIndex = -1
                                tgSortCrf(ilUpper).iDuplIndex = -1
                                tgSortCrf(ilUpper).iVpfIndex = ilVpfIndex
                                tgSortCrf(ilUpper).tCrf = tmCrf
                                ReDim Preserve tgSortCrf(0 To ilUpper + 1) As SORTCRF
                                ilUpper = ilUpper + 1
                            End If
                        End If
                    End If
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
        End If
        ilUpper = UBound(tgSortCrf)
        If ilUpper > 0 Then
            ArraySortTyp fnAV(tgSortCrf(), 0), ilUpper, 0, LenB(tgSortCrf(0)), 0, LenB(tgSortCrf(0).sKey), 0
        End If
        imLastIndex = -1
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        If ilUpper > vbcRot.LargeChange + 1 Then
            vbcRot.Max = ilUpper - vbcRot.LargeChange - 1
        Else
            vbcRot.Max = 0
        End If
        vbcRot.Value = 0
        imIgnoreVbcChg = False
        btrExtClear hmCrf   'Clear any previous extend operation
        For ilLoop = 0 To ilUpper - 1 Step 1
            slNameCode = tgSortCrf(ilLoop).sKey
            tmCrf = tgSortCrf(ilLoop).tCrf
            ilRet = gParseItem(slNameCode, 1, "|", slName)
            slStr = slName & "|"
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            llCntrNo = tgSortCrf(ilLoop).lCntrNo
            slStr = slStr & Trim$(str$(llCntrNo)) & "|"
            ilRet = gParseItem(slNameCode, 3, "|", slName)
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            slStr = slStr & Left$(slName, 10) & "|"
            slStr = slStr & Trim$(tmCrf.sZone) & "|"
            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slStartDate
            If slStartDate <> "" Then
                slStartDate = Left$(slStartDate, Len(slStartDate) - 3)
            End If
            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slEndDate
            If slEndDate <> "" Then
                slEndDate = Left$(slEndDate, Len(slEndDate) - 3)
            End If
            slStr = slStr & slStartDate & "-" & slEndDate & "|"
            gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slStartTime
            gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slEndTime
            slStr = slStr & slStartTime & "-" & slEndTime & "|"
            For ilDay = 0 To 6 Step 1
                slStr = slStr & tmCrf.sDay(ilDay) '& "|"
            Next ilDay
            slStr = slStr & "|"
            slStr = slStr & Trim$(str$(tmCrf.iLen)) & "|"
            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
                If tmAnf.iCode <> tmCrf.ianfCode Then
                    tmAnfSrchKey.iCode = tmCrf.ianfCode
                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mRotPopErr
                    gBtrvErrorMsg ilRet, "mRotPop (btrGetEqual):" & "Anf.Btr", ExpBkCpy
                    On Error GoTo 0
                End If
                If tmCrf.sInOut = "O" Then
                    slName = "O" & Trim$(tmAnf.sName)
                Else
                    slName = Trim$(tmAnf.sName)
                End If
            Else
                slName = "All avails"
            End If
            slStr = slStr & slName & "|"
            Select Case tmCrf.sRotType
                Case "A"
                    slStr = slStr & "CS " & "|"
                Case "O"
                    slStr = slStr & "OBB" & "|"
                Case "C"
                    slStr = slStr & "CBB" & "|"
                Case "E"
                    slStr = slStr & "ABB" & "|"
                Case Else
                    slStr = slStr & " |"
            End Select
            If tmCrf.lCsfCode > 0 Then
                slStr = slStr & "C"
            Else
                slStr = slStr & " "
            End If
            If lbcRot.ListCount < vbcRot.LargeChange + 1 Then
                lbcRot.AddItem slStr
            End If
            tgSortCrf(ilLoop).sKey = slStr
        Next ilLoop
    End If
    pbcLbcRot_Paint
    If ckcAll.Value = vbChecked Then
        'ckcAll_Click True
        ckcAll_Click
    End If
    Exit Sub
mRotPopErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set button state               *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim slTranDate As String
    Dim ilLoop As Integer
    Dim ilSelected As Integer
    ilSelected = False
    For ilLoop = 0 To UBound(tgSortCrf) - 1 Step 1
        If tgSortCrf(ilLoop).iSelected Then
            ilSelected = True
            Exit For
        End If
    Next ilLoop
    slTranDate = edcTranDate.Text
    If (ilSelected) And (gValidDate(slTranDate)) Then
        cmcExport.Enabled = True
    Else
        cmcExport.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowRotInfo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Dulpicate and Combined    *
'*                      rotations                      *
'*                                                     *
'*******************************************************
Private Sub mShowRotInfo()
    Dim ilShow As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilButtonIndex As Integer
    ilButtonIndex = imButtonIndex
    If (imButtonIndex < 0) And (imButtonIndex > UBound(tgSortCrf) - 1) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    For ilLoop = 0 To 4 Step 1
        lacRotInfo(ilLoop).Caption = ""
    Next ilLoop
    ilShow = 0
    lacRotInfo(ilShow).Caption = "No Matching or Combinations"
    ilIndex = tgSortCrf(ilButtonIndex).iCombineIndex
    Do While ilIndex >= 0
        lacRotInfo(ilShow).Caption = "Combining with Rotation #:" & str$(tgCombineCrf(ilIndex).tCrf.iRotNo) & str$(tgCombineCrf(ilIndex).lCntrNo) & " " & Trim$(tgCombineCrf(ilIndex).sVehName)
        ilShow = ilShow + 1
        If ilShow > 4 Then
            Exit Sub
        End If
        ilIndex = tgCombineCrf(ilIndex).iCombineIndex
    Loop
    ilIndex = tgSortCrf(ilButtonIndex).iDuplIndex
    Do While ilIndex >= 0
        lacRotInfo(ilShow).Caption = "Matching Rotation #:" & str$(tgDuplCrf(ilIndex).tCrf.iRotNo) & str$(tgDuplCrf(ilIndex).lCntrNo) & " " & Trim$(tgDuplCrf(ilIndex).sVehName)
        ilShow = ilShow + 1
        If ilShow > 4 Then
            Exit Sub
        End If
        ilIndex = tgDuplCrf(ilIndex).iDuplIndex
    Loop
    DoEvents
    If (imButtonIndex < 0) And (imButtonIndex > UBound(tgSortCrf) - 1) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    If (tgSpf.sUseProdSptScr <> "P") Then
        If ilShow > 0 Then
            lacRotInfo(ilShow).Caption = "Product: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        Else
            lacRotInfo(1).Caption = "Product: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        End If
    Else
        If ilShow > 0 Then
            lacRotInfo(ilShow).Caption = "Short Title: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        Else
            lacRotInfo(1).Caption = "Short Title: " & Trim$(tgSortCrf(ilButtonIndex).sCntrProd)
        End If
    End If
    plcRotInfo.Visible = True
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
    csiSetValue imWaitCount, imTimeDelay, imLockValue, imTranLog

    Screen.MousePointer = vbDefault
    sgDoneMsg = ""
    igManUnload = YES
    Unload ExpBkCpy
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicles to generate *
'*                     bulk feed                       *
'*                                                     *
'*******************************************************
Private Sub mVehPop(ilfirstTime As Integer)
'
'   tmVef will contain the primary vehicle (primary vehicle is
'        the one within the group that is earliest alphabetically)
'        Find no group defined for the vehicle, then tmVef contain
'        the vehicle and tmVpfInfo will contain no links
'   tmVpfInfo will contain the other vehicles grouped with the primary
'   tVpf.iGLink will only contain primary and vehicles without groups only
'               the other vehicles within a group are removed from iGLink
'
'   For Prime remove 24-Hour Format from link table as the spots only
'   map into 12m-6am and rotation should not be generated for 24-hour format
'
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slChar As String
    Dim slPrefix As String
    Dim slSecond As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim ilOk As Integer
    Dim ilLetter As Integer
    Dim ilFound As Integer
    Dim ilUse As Integer
    Dim ilVpfIndex As Integer
    Dim ilLoop1 As Integer
    Dim ilLoop2 As Integer
    Dim ilNextLk As Integer
    Dim slDate As String
    Dim ilPrevIndex As Integer
    Dim tlVef As VEF
    Dim tlSVef As VEF
    lbcVehicle.Clear
    lbcVehicleCode.Clear
    pbclbcVehicle_Paint
    DoEvents
    If ilfirstTime Then

        imVefRecLen = Len(tlVef)
        'Determine which vehicles are to be combined (same bulk group Number)
        ReDim tmVpfInfo(LBound(tgVpf) To LBound(tgVpf)) As VPFINFO
        ReDim tmLkVehInfo(0 To 0) As LKVEHINFO
        For ilLoop = LBound(tgVpf) To UBound(tgVpf) Step 1
            tmVefSrchKey.iCode = tgVpf(ilLoop).iVefKCode
            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ((tlVef.sType = "A") Or (tlVef.sType = "C") Or (tlVef.sType = "S")) And (tgVpf(ilLoop).sGGroupNo <> "-") Then
                ilFound = False
                If tgVpf(ilLoop).sGGroupNo <> " " Then  'Group number only defined for airing and conventional vehicles
                    For ilTest = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
                        If tmVpfInfo(ilTest).tVpf.sGGroupNo = tgVpf(ilLoop).sGGroupNo Then
                            ilFound = True
                            'Save primary alphabetically
                            tmVefSrchKey.iCode = tmVpfInfo(ilTest).tVpf.iVefKCode
                            ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If StrComp(tlSVef.sName, tlVef.sName, 1) <= 0 Then
                                'If tmVpfInfo(ilTest).iNoVefLinks <= UBound(tmVpfInfo(ilTest).iVefLink) Then
                                '    tmVpfInfo(ilTest).iVefLink(tmVpfInfo(ilTest).iNoVefLinks) = tgVpf(ilLoop).iVefKCode
                                '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                'End If
                                tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                                ilIndex = UBound(tmLkVehInfo)
                                If ilNextLk = -1 Then
                                    tmVpfInfo(ilTest).iFirstLkVehInfo = ilIndex
                                Else
                                    Do
                                        If tmLkVehInfo(ilNextLk).iNextLkVehInfo = -1 Then
                                            tmLkVehInfo(ilNextLk).iNextLkVehInfo = ilIndex
                                            Exit Do
                                        End If
                                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                    Loop
                                End If
                                tmLkVehInfo(ilIndex).iVefCode = tlVef.iCode
                                tmLkVehInfo(ilIndex).sVefName = tlVef.sName
                                tmLkVehInfo(ilIndex).iNextLkVehInfo = -1
                                ReDim Preserve tmLkVehInfo(0 To ilIndex + 1) As LKVEHINFO
                            Else
                                'Switch places
                                'If tmVpfInfo(ilTest).iNoVefLinks <= UBound(tmVpfInfo(ilTest).iVefLink) Then
                                '    tmVpfInfo(ilTest).iVefLink(tmVpfInfo(ilTest).iNoVefLinks) = tmVpfInfo(ilTest).tVpf.iVefKCode
                                '    tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                'End If
                                tmVpfInfo(ilTest).iNoVefLinks = tmVpfInfo(ilTest).iNoVefLinks + 1
                                ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                                ilIndex = UBound(tmLkVehInfo)
                                If ilNextLk = -1 Then
                                    tmVpfInfo(ilTest).iFirstLkVehInfo = ilIndex
                                Else
                                    Do
                                        If tmLkVehInfo(ilNextLk).iNextLkVehInfo = -1 Then
                                            tmLkVehInfo(ilNextLk).iNextLkVehInfo = ilIndex
                                            Exit Do
                                        End If
                                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                                    Loop
                                End If
                                tmLkVehInfo(ilIndex).iVefCode = tlSVef.iCode
                                tmLkVehInfo(ilIndex).sVefName = tlSVef.sName
                                tmLkVehInfo(ilIndex).iNextLkVehInfo = -1
                                ReDim Preserve tmLkVehInfo(0 To ilIndex + 1) As LKVEHINFO
                                tmVpfInfo(ilTest).tVpf = tgVpf(ilLoop)
                            End If
                            Exit For
                        End If
                    Next ilTest
                End If
                If Not ilFound Then
                    ilIndex = UBound(tmVpfInfo)
                    tmVpfInfo(ilIndex).tVpf = tgVpf(ilLoop)
                    'If StrComp(Trim$(tlVef.sName), "Prime", 1) = 0 Then
                    '    'Remove 24-Hour format from tgVpf.iGLink
                    '    For ilTest = LBound(tgVpf(ilLoop).iGLink) To UBound(tgVpf(ilLoop).iGLink) Step 1
                    '        If tgVpf(ilLoop).iGLink(ilTest) > 0 Then
                    '            tmVefSrchKey.iCode = tgVpf(ilTest).iVefKCode
                    '            ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    '            If StrComp(Trim$(tlSVef.sName), "24-HOUR FORMATS", 1) = 0 Then
                    '                tmVpfInfo(ilIndex).tVpf.iGLink(ilTest) = 0
                    '                For ilMove = ilTest To UBound(tgVpf(ilLoop).iGLink) - 1 Step 1
                    '                    tmVpfInfo(ilIndex).tVpf.iGLink(ilMove) = tmVpfInfo(ilIndex).tVpf.iGLink(ilMove + 1)
                    '                Next ilMove
                    '                Exit For
                    '                tmVpfInfo(ilIndex).tVpf.iGLink(UBound(tgVpf(ilLoop).iGLink)) = 0
                    '            End If
                    '        End If
                    '    Next ilTest
                    'End If
                    'tmVpfInfo(ilIndex).iNoVefLinks = LBound(tmVpfInfo(ilTest).iVefLink)
                    'tmVpfInfo(ilIndex).iVefLink(LBound(tmVpfInfo(ilTest).iVefLink)) = 0
                    'tmVpfInfo(ilIndex).sVefName(LBound(tmVpfInfo(ilTest).iVefLink)) = ""
                    tmVpfInfo(ilIndex).iNoVefLinks = 0
                    tmVpfInfo(ilIndex).iFirstLkVehInfo = -1
                    tmVpfInfo(ilIndex).iFirstSALink = -1
                    'tmVpfInfo(ilIndex).iVefLink(LBound(tmVpfInfo(ilTest).iVefLink)) = 0
                    'tmVpfInfo(ilIndex).sVefName(LBound(tmVpfInfo(ilTest).iVefLink)) = ""
                    ReDim Preserve tmVpfInfo(LBound(tgVpf) To ilIndex + 1) As VPFINFO
                End If
            End If
        Next ilLoop
        'Order groups alphabetically
        For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
            'If tmVpfInfo(ilLoop).iNoVefLinks > LBound(tmVpfInfo(ilLoop).iVefLink) + 1 Then
            '    'Sort names- only the first two for now
            '    tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink))
            '    ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '    tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1)
            '    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            '    If StrComp(tlSVef.sName, tlVef.sName, 1) > 0 Then
            '        'Switch names
            '        tmVefSrchKey.iCode = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink))
            '        tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink)) = tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1)
            '        tmVpfInfo(ilLoop).iVefLink(LBound(tmVpfInfo(ilLoop).iVefLink) + 1) = tmVefSrchKey.iCode
            '    End If
            'End If
            If tmVpfInfo(ilLoop).iNoVefLinks > 1 Then
                ilIndex = tmVpfInfo(ilLoop).iFirstLkVehInfo
                ilNextLk = tmLkVehInfo(ilIndex).iNextLkVehInfo
                Do While ilNextLk >= 0
                    Do
                        tmVefSrchKey.iCode = tmLkVehInfo(ilIndex).iVefCode
                        ilRet = btrGetEqual(hmVef, tlSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        tmVefSrchKey.iCode = tmLkVehInfo(ilNextLk).iVefCode
                        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If StrComp(tlSVef.sName, tlVef.sName, 1) > 0 Then
                            'Switch names
                            tmVefSrchKey.iCode = tmLkVehInfo(ilIndex).iVefCode
                            slName = tmLkVehInfo(ilIndex).sVefName
                            tmLkVehInfo(ilIndex).iVefCode = tmLkVehInfo(ilNextLk).iVefCode
                            tmLkVehInfo(ilIndex).sVefName = tmLkVehInfo(ilNextLk).sVefName
                            tmLkVehInfo(ilNextLk).iVefCode = tmVefSrchKey.iCode
                            tmLkVehInfo(ilNextLk).sVefName = slName
                        End If
                        ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                    Loop While ilNextLk >= 0
                    ilIndex = tmLkVehInfo(ilIndex).iNextLkVehInfo
                    ilNextLk = tmLkVehInfo(ilIndex).iNextLkVehInfo
                Loop
            End If
        Next ilLoop
        ReDim tmVef(0 To 0) As VEF
        imVefRecLen = Len(tmVef(0))
        ilUpper = 0
        ilRet = btrGetFirst(hmVef, tmVef(ilUpper), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ilUse = False
            For ilTest = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
                If tmVpfInfo(ilTest).tVpf.iVefKCode = tmVef(ilUpper).iCode Then
                    ilUse = True
                    Exit For
                End If
                ilFound = False
                'For ilIndex = LBound(tmVpfInfo(ilTest).iVefLink) To tmVpfInfo(ilTest).iNoVefLinks - 1 Step 1
                '    If tmVpfInfo(ilTest).iVefLink(ilIndex) = tmVef(ilUpper).iCode Then
                '        tmVpfInfo(ilTest).sVefName(ilIndex) = Trim$(tmVef(ilUpper).sName)
                '        ilFound = True
                '        Exit For
                '    End If
                'Next ilIndex
                ilNextLk = tmVpfInfo(ilTest).iFirstLkVehInfo
                Do While ilNextLk >= 0
                    If tmLkVehInfo(ilNextLk).iVefCode = tmVef(ilUpper).iCode Then
                        ilFound = True
                        Exit For
                    End If
                    ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                Loop
                If ilFound Then
                    Exit For
                End If
            Next ilTest
            If ilUse Then
                ilUpper = ilUpper + 1
                ReDim Preserve tmVef(0 To ilUpper) As VEF
            End If
            ilRet = btrGetNext(hmVef, tmVef(ilUpper), imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        'Build Selling to Airing link table
        ReDim tmSALink(0 To 0) As LKVEHINFO
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilLoop).sType = "S") Then
                'slDate = Format$(Now, "m/d/yy")
                'gBuildLinkArray hmVlf, tmVef(ilLoop), slDate, imDVefCode()
                slDate = Format$(gNow(), "m/d/yy")
                For ilUse = 0 To 8 Step 1
                    gBuildLinkArray hmVlf, tmVef(ilLoop), slDate, imDVefCode()
                    If UBound(imDVefCode) > LBound(imDVefCode) Then
                        Exit For
                    End If
                    slDate = gIncOneWeek(slDate)
                Next ilUse
                'slDate = gObtainPrevMonday(slDate)
                'llDate = gDateValue(slDate)
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ReDim imDVefCode(0 To 0) As Integer
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpBkCpy, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                'llDate = llDate + 5
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpBkCpy, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                'llDate = llDate + 1
                'ReDim tmTVlf(1 To 1) As VLF
                'gObtainVlf "S", hmVlf, tmVef(ilLoop).iCode, llDate, tmTVlf()
                'ilPrevAir = -1
                'For ilLoop1 = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
                '    If ilPrevAir <> tmTVlf(ilLoop1).iAirCode Then
                '        ilFound = False
                '        For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                '            If tmTVlf(ilLoop1).iAirCode = imDVefCode(ilTest) Then
                '                ilFound = True
                '                Exit For
                '            End If
                '        Next ilTest
                '        If Not ilFound Then
                '            If tgVpf(gVpfFind(ExpBkCpy, tmTVlf(ilLoop1).iAirCode)).sGGroupNo <> "-" Then
                '                imDVefCode(UBound(imDVefCode)) = tmTVlf(ilLoop1).iAirCode
                '                ReDim Preserve imDVefCode(0 To UBound(imDVefCode) + 1) As Integer
                '            End If
                '        End If
                '        ilPrevAir = tmTVlf(ilLoop1).iAirCode
                '    End If
                'Next ilLoop1
                ilVpfIndex = mFindVpfIndex(tmVef(ilLoop).iCode)
                ilIndex = UBound(tmSALink)
                ilPrevIndex = -1
                For ilTest = 0 To UBound(imDVefCode) - 1 Step 1
                    If tgVpf(gVpfFind(ExpBkCpy, imDVefCode(ilTest))).sGGroupNo <> "-" Then
                        If ilPrevIndex >= 0 Then
                            tmSALink(ilPrevIndex).iNextLkVehInfo = ilIndex
                        Else
                            tmVpfInfo(ilVpfIndex).iFirstSALink = ilIndex
                        End If
                        tmSALink(ilIndex).iVefCode = imDVefCode(ilTest)
                        tmSALink(ilIndex).sVefName = ""
                        tmSALink(ilIndex).iNextLkVehInfo = -1
                        ReDim Preserve tmSALink(0 To ilIndex + 1) As LKVEHINFO
                        ilPrevIndex = ilIndex
                        ilIndex = ilIndex + 1
                    End If
                Next ilTest
            End If
        Next ilLoop
        ReDim imDVefCode(0 To 0) As Integer
        'Generate names for file prefix (only for conventional and airing)
        For ilLoop = 0 To UBound(tmVef) - 1 Step 1
            If (tmVef(ilLoop).sType = "A") Or (tmVef(ilLoop).sType = "C") Then
                slName = Trim$(tmVef(ilLoop).sName)
                slPrefix = Left$(slName, 1)
                ilOk = False
                'Try first letter plus letter after blank or -
                For ilIndex = 2 To Len(slName) Step 1
                    slChar = Mid$(slName, ilIndex, 1)
                    If (slChar = " ") Or (slChar = "-") Then
                        If slPrefix <> "X" Then
                            ilOk = True
                            slSecond = Mid$(slName, ilIndex + 1, 1)
                            If ((Asc(slSecond) >= Asc("0")) And (Asc(slSecond) <= Asc("9"))) Or ((Asc(slSecond) >= Asc("A")) And (Asc(slSecond) <= Asc("Z"))) Then
                                slPrefix = slPrefix & slSecond
                                For ilTest = 0 To ilLoop - 1 Step 1
                                    If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                        ilOk = False
                                        Exit For
                                    End If
                                Next ilTest
                            Else
                                ilOk = False
                            End If
                        End If
                        If ilOk Then
                            tmVef(ilTest).sCodeStn = slPrefix
                            Exit For
                        Else
                            slPrefix = Left$(slName, 1)
                        End If
                    End If
                Next ilIndex
                If Not ilOk Then
                    'Try first two letters of the name
                    slPrefix = Left$(slName, 2)
                    If slPrefix <> "X" Then
                        ilOk = True
                        For ilTest = 0 To ilLoop - 1 Step 1
                            If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                ilOk = False
                                Exit For
                            End If
                        Next ilTest
                    End If
                    If ilOk Then
                        tmVef(ilTest).sCodeStn = slPrefix
                    Else
                        For ilLetter = Asc("A") To Asc("Z") Step 1
                            slPrefix = Left$(slName, 1) & Chr(ilLetter)
                            If slPrefix <> "X" Then
                                ilOk = True
                                For ilTest = 0 To ilLoop - 1 Step 1
                                    If Trim$(tmVef(ilTest).sCodeStn) = slPrefix Then
                                        ilOk = False
                                        Exit For
                                    End If
                                Next ilTest
                            End If
                            If ilOk Then
                                tmVef(ilTest).sCodeStn = slPrefix
                                Exit For
                            End If
                        Next ilLetter
                    End If
                End If
            End If
        Next ilLoop
        ''Adjust links for selling only points to primary airing.
        'For ilLoop1 = 0 To UBound(tmVef) - 1 Step 1
        '    If (tmVef(ilLoop1).sType = "S") Then
        '        ilVpfIndex = mFindVpfIndex(tmVef(ilLoop1).iCode)
        '        If ilVpfIndex >= 0 Then
        '            For ilLoop2 = LBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
        '                ilFound = False
        '                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) > 0 Then
        '                    For ilLoop = LBound(tmVpfInfo) To UBound(tmVpfInfo) - 1 Step 1
        '                        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
        '                            ilFound = True
        '                        End If
        '                        If Not ilFound Then
        '                            For ilIndex = LBound(tmVpfInfo(ilLoop).iVefLink) To tmVpfInfo(ilLoop).iNoVefLinks - 1 Step 1
        '                                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).iVefLink(ilIndex) Then
        '                                    tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop2) = tmVpfInfo(ilLoop).tVpf.iVefKCode
        '                                    ilFound = True
        '                                End If
        '                            Next ilIndex
        '                        End If
        '                        If ilFound Then
        '                            'Remove from links that are associated with this vehicle
        '                            For ilLoop3 = ilLoop2 + 1 To UBound(tmVpfInfo(ilVpfIndex).tVpf.iGLink) Step 1
        '                                ilFound = False
        '                                If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) > 0 Then
        '                                    If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = tmVpfInfo(ilLoop).tVpf.iVefKCode Then
        '                                        tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = 0
        '                                    End If
        '                                    For ilIndex = LBound(tmVpfInfo(ilLoop).iVefLink) To tmVpfInfo(ilLoop).iNoVefLinks - 1 Step 1
        '                                        If tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = tmVpfInfo(ilLoop).iVefLink(ilIndex) Then
        '                                            tmVpfInfo(ilVpfIndex).tVpf.iGLink(ilLoop3) = 0
        '                                        End If
        '                                    Next ilIndex
        '                                End If
        '                            Next ilLoop3
        '                            Exit For
        '                        End If
        '                    Next ilLoop
        '                End If
        '            Next ilLoop2
        '        End If
        '    End If
        'Next ilLoop1
    Else
        For ilLoop = 0 To lbcVeh.ListCount - 1 Step 1
            lbcVeh.Selected(ilLoop) = False
            tmcRot.Enabled = False
        Next ilLoop
    End If
    'Only show Airing and Conventional
    For ilLoop2 = 0 To UBound(tmVef) - 1 Step 1
        If (tmVef(ilLoop2).sType = "A") Or (tmVef(ilLoop2).sType = "C") Then
            ilVpfIndex = mFindVpfIndex(tmVef(ilLoop2).iCode)
            If ilVpfIndex >= 0 Then
                slName = Trim$(tmVef(ilLoop2).sName)
                'For ilLoop1 = LBound(tmVpfInfo(ilVpfIndex).sVefName) To tmVpfInfo(ilVpfIndex).iNoVefLinks - 1 Step 1
                '    slName = slName & " " & Trim$(tmVpfInfo(ilVpfIndex).sVefName(ilLoop1))
                'Next ilLoop1
                ilNextLk = tmVpfInfo(ilVpfIndex).iFirstLkVehInfo
                Do While ilNextLk >= 0
                    slName = slName & " " & Trim$(tmLkVehInfo(ilNextLk).sVefName)
                    ilNextLk = tmLkVehInfo(ilNextLk).iNextLkVehInfo
                Loop
                lbcVehicleCode.AddItem slName & "|" & "0:0" & "\" & Trim$(str$(ilVpfIndex))
            End If
        End If
    Next ilLoop2
    DoEvents
    For ilLoop1 = 0 To lbcVehicleCode.ListCount - 1 Step 1
        slNameCode = lbcVehicleCode.List(ilLoop1)
        ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get application name
        lbcVehicle.AddItem slName
    Next ilLoop1
    pbclbcVehicle_Paint
    DoEvents
End Sub

Private Sub lbcVehicle_Scroll()
    pbclbcVehicle_Paint
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
                    'edcStartDate.Text = Format$(llDate, "m/d/yy")
                    'edcStartDate.SelStart = 0
                    'edcStartDate.SelLength = Len(edcStartDate.Text)
                    'imBypassFocus = True
                    'edcStartDate.SetFocus
                    'Exit Sub
                ElseIf imDateBox = 2 Then
                    'edcEndDate.Text = Format$(llDate, "m/d/yy")
                    'edcEndDate.SelStart = 0
                    'edcEndDate.SelLength = Len(edcEndDate.Text)
                    'imBypassFocus = True
                    'edcEndDate.SetFocus
                    'Exit Sub
                ElseIf imDateBox = 3 Then
                    edcTranDate.Text = Format$(llDate, "m/d/yy")
                    edcTranDate.SelStart = 0
                    edcTranDate.SelLength = Len(edcTranDate.Text)
                    imBypassFocus = True
                    edcTranDate.SetFocus
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
        'edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        'edcEndDate.SetFocus
    ElseIf imDateBox = 3 Then
        edcTranDate.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcLbcRot_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRotEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 10) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilRotEnd = lbcRot.TopIndex + lbcRot.Height \ fgListHtArial825
    If ilRotEnd > lbcRot.ListCount Then
        ilRotEnd = lbcRot.ListCount
    End If
    If lbcRot.ListCount <= lbcRot.Height \ fgListHtArial825 Then
        llWidth = lbcRot.Width - 30
    Else
        llWidth = lbcRot.Width - igScrollBarWidth - 30
    End If
    pbcLbcRot.Width = llWidth
    pbcLbcRot.Cls
    llFgColor = pbcLbcRot.ForeColor
    For ilLoop = lbcRot.TopIndex To ilRotEnd - 1 Step 1
        pbcLbcRot.ForeColor = llFgColor
        If lbcRot.MultiSelect = 0 Then
            If lbcRot.ListIndex = ilLoop Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        Else
            If lbcRot.Selected(ilLoop) Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        End If
        slStr = lbcRot.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To 11 Step 1
            pbcLbcRot.CurrentX = imListFieldRot(ilField)
            pbcLbcRot.CurrentY = (ilLoop - lbcRot.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            gAdjShowLen pbcLbcRot, slStr, imListFieldRot(ilField + 1) - imListFieldRot(ilField)
            pbcLbcRot.Print slStr
            ilFieldIndex = ilFieldIndex + 1
        Next ilField
        pbcLbcRot.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub pbclbcVehicle_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilVehiclesEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 1) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilVehiclesEnd = lbcVehicle.TopIndex + lbcVehicle.Height \ fgListHtArial825
    If ilVehiclesEnd > lbcVehicle.ListCount Then
        ilVehiclesEnd = lbcVehicle.ListCount
    End If
    If lbcVehicle.ListCount <= lbcVehicle.Height \ fgListHtArial825 Then
        llWidth = lbcVehicle.Width - 30
    Else
        llWidth = lbcVehicle.Width - igScrollBarWidth - 30
    End If
    pbcLbcVehicle.Width = llWidth
    pbcLbcVehicle.Cls
    llFgColor = pbcLbcVehicle.ForeColor
    For ilLoop = lbcVehicle.TopIndex To ilVehiclesEnd - 1 Step 1
        pbcLbcVehicle.ForeColor = llFgColor
        If lbcVehicle.MultiSelect = 0 Then
            If lbcVehicle.ListIndex = ilLoop Then
                gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehicle.ForeColor = vbWhite
            End If
        Else
            If lbcVehicle.Selected(ilLoop) Then
                gPaintArea pbcLbcVehicle, CSng(0), CSng((ilLoop - lbcVehicle.TopIndex) * fgListHtArial825), CSng(pbcLbcVehicle.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcVehicle.ForeColor = vbWhite
            End If
        End If
        slStr = lbcVehicle.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To UBound(slFields) Step 1
            pbcLbcVehicle.CurrentX = imListFieldVeh(ilField)
            pbcLbcVehicle.CurrentY = (ilLoop - lbcVehicle.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            gAdjShowLen pbcLbcVehicle, slStr, imListFieldVeh(ilField + 1) - imListFieldVeh(ilField)
            pbcLbcVehicle.Print slStr
            ilFieldIndex = ilFieldIndex + 1
        Next ilField
        pbcLbcVehicle.ForeColor = llFgColor
    Next ilLoop

End Sub

Private Sub tmcDDE_Timer()
    tmcDDE.Enabled = False
    'mInitDDE
    'mSendHelpMsg "BT"
End Sub
Private Sub tmcRot_Timer()
    tmcRot.Enabled = False
    Screen.MousePointer = vbHourglass
    ReDim tmCyfTest(0 To 0) As CYFTEST  'Save each Cyf to be sent
    'mVehPop False
    mRotPop
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcRot_Change()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tgSortCrf) - 1 Then
        ilEndIndex = UBound(tgSortCrf) - 1
    End If
    ilValue = False
    If UBound(tgSortCrf) < vbcRot.LargeChange + 1 Then
        llRg = CLng(UBound(tgSortCrf) - 1) * &H10000 Or 0
    Else
        llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
    End If
    llRet = SendMessageByNum(lbcRot.HWnd, LB_SELITEMRANGE, ilValue, llRg)
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.List(ilIndex) = tgSortCrf(ilLoop).sKey
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.Selected(ilIndex) = tgSortCrf(ilLoop).iSelected
        ilIndex = ilIndex + 1
    Next ilLoop
    pbcLbcRot_Paint
    imIgnoreVbcChg = False
End Sub
Private Sub vbcRot_Scroll()
    vbcRot_Change
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Export Bulk Copy"
End Sub

Private Sub mExportCopyHeader(ilLineNo As Integer, ilNoCopyLines As Integer, ilPageNo As Integer, slBlank As String, slCopyHeader As String, slVehName As String, slStnCode As String, slFeedNo As String, slTranDate As String, slPrevNewInv As String)
    Dim slStr As String
    
    If ilLineNo + ilNoCopyLines > 52 Then
        If ilPageNo = 0 Then
            'slRecord = ""
            If Not mExportLine(slBlank, ilLineNo) Then
                Exit Sub
            End If
        Else
            slCopyHeader = Chr(12)  'Form Feed
            If Not mExportLine(slCopyHeader, ilLineNo) Then
                Exit Sub
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        slCopyHeader = " "
        Do While Len(slCopyHeader) < 35
            slCopyHeader = slCopyHeader & " "
        Loop
        slCopyHeader = slCopyHeader & Trim$(tgSpf.sGClient)
        If Not mExportLine(slCopyHeader, ilLineNo) Then
            Exit Sub
        End If
        slCopyHeader = " "
        Do While Len(slCopyHeader) < 35
            slCopyHeader = slCopyHeader & " "
        Loop
        slCopyHeader = slCopyHeader & slVehName
        If Not mExportLine(slCopyHeader, ilLineNo) Then
            Exit Sub
        End If
        slCopyHeader = " "
        Do While Len(slCopyHeader) < 35
            slCopyHeader = slCopyHeader & " "
        Loop
        slCopyHeader = slCopyHeader & "Commercial Feed # " & slStnCode & "-" & slFeedNo
        If Not mExportLine(slCopyHeader, ilLineNo) Then
            Exit Sub
        End If
        slCopyHeader = " "
        Do While Len(slCopyHeader) < 35
            slCopyHeader = slCopyHeader & " "
        Loop
        slCopyHeader = slCopyHeader & slTranDate & "  "
        slCopyHeader = slCopyHeader & "Page:"
        slStr = Trim$(str$(ilPageNo))
        Do While Len(slStr) < 5
            slStr = " " & slStr
        Loop
        slCopyHeader = slCopyHeader & slStr
        If Not mExportLine(slCopyHeader, ilLineNo) Then
            Exit Sub
        End If
        slCopyHeader = ""
        If Not mExportLine(slBlank, ilLineNo) Then
            Exit Sub
        End If
        If Not mExportLine(slBlank, ilLineNo) Then
            Exit Sub
        End If
        If (slPrevNewInv <> "") And (ilNoCopyLines = 1) Then
            If Not mExportLine(slPrevNewInv, ilLineNo) Then
                Exit Sub
            End If
            slCopyHeader = "     "
            slCopyHeader = slCopyHeader & "------------------------------"
            If Not mExportLine(slCopyHeader, ilLineNo) Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub mExportRotHeader(ilLineNo As Integer, ilPageNo As Integer, slRecord As String, slVehName As String, slTranDate As String)
    Dim slStr As String
    If ilLineNo >= 52 Then
        If ilPageNo > 0 Then
            slRecord = Chr(12)  'Form Feed
            If Not mExportLine(slRecord, ilLineNo) Then
                Exit Sub
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        slRecord = "-"
        Do While Len(slRecord) < 78
            slRecord = slRecord & "-"
        Loop
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        slRecord = " "
        Do While Len(slRecord) < 68
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & "Page:"
        slStr = Trim$(str$(ilPageNo))
        Do While Len(slStr) < 5
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        slRecord = ""
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        slRecord = Trim$(tgSpf.sGClient) & " " & slVehName & " Network Feed Instructions " & slTranDate
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        slRecord = ""
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
        If Not mExportLine(slRecord, ilLineNo) Then
            Exit Sub
        End If
    End If
End Sub

Private Function mExportXFerHeader(ilLineNo As Integer, ilPageNo As Integer, slTranDate As String, slRecord As String) As Boolean
    Dim slStr As String
    
    mExportXFerHeader = True
    If ilLineNo >= 48 Then
        If ilPageNo = 0 Then
            slRecord = ""
            If Not mExportLine(slRecord, ilLineNo) Then
                mExportXFerHeader = False
                Exit Function
            End If
        Else
            slRecord = Chr(12)  'Form Feed
            If Not mExportLine(slRecord, ilLineNo) Then
                mExportXFerHeader = False
                Exit Function
            End If
        End If
        ilPageNo = ilPageNo + 1
        ilLineNo = 0
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & Trim$(tgSpf.sGClient)
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & "Cross Reference"
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = " "
        Do While Len(slRecord) < 35
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & slTranDate & "  "
        slRecord = slRecord & "Page:"
        slStr = Trim$(str$(ilPageNo))
        Do While Len(slStr) < 5
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = ""
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
        slRecord = "Short Title"
        Do While Len(slRecord) < 20
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Cart"
        Do While Len(slRecord) < 31
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " ISCI"
        Do While Len(slRecord) < 52
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Creative Title"
        Do While Len(slRecord) < 83
            slRecord = slRecord & " "
        Loop
        slRecord = slRecord & " Vehicle"
        If Not mExportLine(slRecord, ilLineNo) Then
            mExportXFerHeader = False
            Exit Function
        End If
    End If
    mExportXFerHeader = True
End Function

Private Sub mExportSendMsg(slMsgFileName As String, slMsgFile As String, slMsgLine As String, ilMsgType As Integer, ilLineNo As Integer, ilNoCopyLines As Integer, ilPageNo As Integer, slBlank As String, slCopyHeader As String, slVehName As String, slStnCode As String, slFeedNo As String, slTranDate As String, slPrevNewInv As String, slRecord As String)
    Dim ilRet As Integer
    Dim ilPos As Integer
    ilRet = 0
    'On Error GoTo mExportSendMsgErr:
    'hmMsg = FreeFile
    slMsgFile = sgExportPath & slMsgFileName
    'Open slMsgFile For Input Access Read As hmMsg
    ilRet = gFileOpen(slMsgFile, "Input Access Read", hmMsg)
    If ilRet = 0 Then
        err.Clear
        Do
            'On Error GoTo mExportSendMsgErr:
            Line Input #hmMsg, slMsgLine
            On Error GoTo 0
            ilRet = err.Number
            If (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            End If
            If Len(slMsgLine) > 0 Then
                If (Asc(slMsgLine) = 26) Then    'Ctrl Z
                    Exit Do
                End If
                ilPos = InStr(UCase$(slMsgLine), "XX/XX/XXXX")
                If ilPos > 0 Then
                    Mid$(slMsgLine, ilPos) = slTranDate
                End If
            End If
            If Not mExportLine(slMsgLine, ilLineNo) Then
                Exit Sub
            End If
            If ilMsgType = 0 Then
                '6/2/16: Replaced GoSub
                'GoSub cmcExportCopyHeader
                mExportCopyHeader ilLineNo, ilNoCopyLines, ilPageNo, slBlank, slCopyHeader, slVehName, slStnCode, slFeedNo, slTranDate, slPrevNewInv
            ElseIf ilMsgType = 1 Then
                '6/2/16: Replace GoSub
                'GoSub cmcExportRotHeader
                mExportRotHeader ilLineNo, ilPageNo, slRecord, slVehName, slTranDate
            End If
        Loop
        Close hmMsg
    End If
    Exit Sub
'mExportSendMsgErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
